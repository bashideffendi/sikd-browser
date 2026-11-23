"""
SIKD/DJPK Client — Akses data APBD dari DJPK Kemenkeu
Data publik tanpa autentikasi.
"""

import requests
import xml.etree.ElementTree as ET
import re

BASE_URL = "https://djpk.kemenkeu.go.id/portal"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/html, */*',
}

_session = requests.Session()
_session.headers.update(HEADERS)

MSG_DJPK_TOLAK = ('Data tidak tersedia di DJPK untuk kombinasi tahun, provinsi, dan pemda ini, '
                  'atau server DJPK sedang bermasalah. Coba ganti pilihan atau ulangi beberapa saat lagi.')
MSG_DJPK_DOWN = 'Server DJPK tidak merespons. Coba lagi beberapa saat lagi.'


class DJPKError(Exception):
    """Gagal ambil data dari DJPK. `str(e)` aman ditampilkan ke user; detail teknis ada di `detail`."""

    def __init__(self, message, detail=''):
        super().__init__(message)
        self.detail = detail


def _get(url, **kwargs):
    """GET ke DJPK; semua kegagalan jaringan/HTTP diubah jadi DJPKError dengan pesan yang ramah."""
    try:
        r = _session.get(url, **kwargs)
        r.raise_for_status()
        return r
    except requests.HTTPError as e:
        raise DJPKError(MSG_DJPK_TOLAK, detail=str(e)) from e
    except requests.RequestException as e:
        raise DJPKError(MSG_DJPK_DOWN, detail=str(e)) from e


def _kode_nama_list(r):
    """Respons DJPK {kode: nama} -> [{kode, nama}]."""
    try:
        data = r.json()
    except ValueError as e:
        raise DJPKError(MSG_DJPK_TOLAK, detail=f'respons bukan JSON: {e}') from e
    if not isinstance(data, dict):
        raise DJPKError(MSG_DJPK_TOLAK, detail=f'format respons tak terduga: {type(data).__name__}')
    return [{"kode": str(kode), "nama": str(nama)} for kode, nama in data.items()]


def get_provinsi(tahun):
    """Get list of provinces for a given year."""
    return _kode_nama_list(_get(f"{BASE_URL}/provinsi/{tahun}", timeout=15))


def get_pemda(provinsi_kode, tahun):
    """Get list of local governments for a province and year."""
    return _kode_nama_list(_get(f"{BASE_URL}/pemda/{provinsi_kode}/{tahun}", timeout=15))


def get_apbd(tahun, provinsi='--', pemda='--', periode='12'):
    """
    Get APBD data. Returns parsed data from XML spreadsheet.
    periode: 1-12 (bulan realisasi)
    provinsi/pemda: kode or '--' for all
    """
    url = f"{BASE_URL}/csv_apbd"
    params = {
        'type': 'apbd',
        'periode': periode,
        'tahun': tahun,
        'provinsi': provinsi,
        'pemda': pemda,
    }

    r = _get(url, params=params, timeout=30)

    content = r.text.strip()
    if not content or '<Workbook' not in content:
        return {'error': 'Data tidak tersedia untuk parameter yang dipilih', 'rows': []}

    return _parse_xml_spreadsheet(content)


def get_apbd_compare(tahun_list, provinsi, pemda, periode='12'):
    """Get APBD data for multiple years for comparison."""
    results = {}
    for tahun in tahun_list:
        try:
            data = get_apbd(tahun, provinsi, pemda, periode)
            results[str(tahun)] = data
        except DJPKError as e:
            results[str(tahun)] = {'error': str(e), 'rows': []}
    return results


# DJPK memberi tepat 1 spasi di depan SEMUA akun, jadi level tidak bisa dibaca dari indentasi.
# Level ditentukan dari nama akun (berlaku untuk format lama 2011-an dan format Permendagri 77/2020).
_LEVEL0 = re.compile(r'^(pendapatan|belanja|pembiayaan)( daerah)?$', re.I)
_LEVEL1 = re.compile(
    r'^(pad|tkdd|dana perimbangan|pendapatan lainnya|lain-lain pendapatan( daerah)?( yang sah)?'
    r'|belanja (tidak langsung|langsung|operasi|lainnya|transfer)'
    r'|(penerimaan|pengeluaran) pembiayaan( daerah)?)$', re.I)


def _level_akun(akun, seen_l1):
    """0 = kelompok (Pendapatan/Belanja/Pembiayaan), 1 = subtotal (PAD, TKDD, ...), 2 = rincian."""
    if _LEVEL0.match(akun):
        seen_l1.clear()
        return 0
    key = akun.lower()
    # Nama subtotal yang muncul lagi di kelompok yang sama = rinciannya (mis. "Pendapatan Lainnya" 2011)
    if _LEVEL1.match(akun) and key not in seen_l1:
        seen_l1.add(key)
        return 1
    return 2


def _parse_xml_spreadsheet(xml_text):
    """Parse XML Spreadsheet format from DJPK into structured data."""
    # Fix namespace issues
    xml_text = re.sub(r'xmlns="[^"]*"', '', xml_text)
    xml_text = re.sub(r'ss:', '', xml_text)

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return {'error': 'Gagal parse data XML', 'rows': []}

    rows = []
    all_rows = root.findall('.//Row')

    if not all_rows:
        return {'error': 'Tidak ada data', 'rows': []}

    # First row is header
    headers = []
    for cell in all_rows[0].findall('Cell'):
        data_el = cell.find('Data')
        headers.append(data_el.text.strip() if data_el is not None and data_el.text else '')

    # Data rows
    for row_el in all_rows[1:]:
        cells = row_el.findall('Cell')
        row = {}
        for i, cell in enumerate(cells):
            data_el = cell.find('Data')
            if data_el is not None and data_el.text:
                val = data_el.text.strip()
                data_type = cell.get('Type', '') or (data_el.get('Type', ''))
                if data_type == 'Number':
                    try:
                        val = float(val)
                    except ValueError:
                        pass
                if i < len(headers) and headers[i]:
                    row[headers[i]] = val
                else:
                    row[f'col_{i}'] = val
            elif i < len(headers) and headers[i]:
                row[headers[i]] = ''

        if row:
            rows.append(row)

    # Categorize rows
    categorized = []
    seen_l1 = set()
    for row in rows:
        akun = str(row.get('Akun', '')).strip()
        level = _level_akun(akun, seen_l1)
        anggaran = row.get('Anggaran', 0)
        realisasi = row.get('Realisasi', 0)
        persentase = row.get('Persentase', '')

        anggaran_num = anggaran if isinstance(anggaran, (int, float)) else 0
        realisasi_num = realisasi if isinstance(realisasi, (int, float)) else 0

        # Skip duplikat: kalau nama dan nilai sama persis dengan baris sebelumnya, skip
        if categorized:
            prev = categorized[-1]
            if (prev['akun'] == akun and
                prev['anggaran'] == anggaran_num and
                prev['realisasi'] == realisasi_num):
                continue

        categorized.append({
            'akun': akun,
            'level': level,
            'anggaran': anggaran_num,
            'realisasi': realisasi_num,
            'persentase': str(persentase),
        })

    return {
        'headers': headers,
        'rows': categorized,
        'total_rows': len(categorized),
    }


def close():
    _session.close()

"""
SIKD Browser — Flask Web App
Jelajahi data APBD dari DJPK Kemenkeu
"""

import os
import io
import re
import atexit
from flask import Flask, request, jsonify, render_template, send_file
import sikd_client
from sikd_client import DJPKError

app = Flask(__name__, static_folder='static', static_url_path='/static')

TAHUN_RE = re.compile(r'^\d{4}$')
KODE_RE = re.compile(r'^\d{2}$')
MAKS_TAHUN_BANDING = 10
MSG_UMUM = 'Terjadi kesalahan di server. Coba lagi beberapa saat lagi.'


def _pesan_tahun(t):
    return f'Tahun "{t[:10]}" tidak valid. Tulis 4 digit, pisahkan dengan koma (misal 2023,2024,2025).'


def _validasi(tahun_list, provinsi, pemda, periode):
    """Cek parameter sebelum diteruskan ke DJPK. Return pesan error atau None."""
    if not tahun_list:
        return 'Tahun harus diisi (misal 2023,2024,2025).'
    if len(tahun_list) > MAKS_TAHUN_BANDING:
        return f'Maksimal {MAKS_TAHUN_BANDING} tahun sekali banding.'
    for t in tahun_list:
        if not TAHUN_RE.match(t):
            return _pesan_tahun(t)
    if not (periode.isdigit() and 1 <= int(periode) <= 12):
        return 'Periode harus angka bulan 1 sampai 12.'
    if provinsi != '--' and not KODE_RE.match(provinsi):
        return 'Kode provinsi tidak valid.'
    if pemda != '--' and not KODE_RE.match(pemda):
        return 'Kode pemda tidak valid.'
    # DJPK balas 500 kalau pemda diisi tapi provinsi "semua"
    if pemda != '--' and provinsi == '--':
        return 'Pilih provinsi dulu sebelum memilih pemda.'
    return None


def _parse_tahun_list(s):
    return [t.strip() for t in s.split(',') if t.strip()]


def _bad_request(msg):
    return jsonify({'error': msg}), 400


def _djpk_gagal(e):
    """DJPK menolak / tidak merespons. Detail teknis cuma ke log, user dapat pesan ramah."""
    app.logger.warning('DJPK gagal: %s', e.detail)
    return jsonify({'error': str(e)}), 502


def _error_umum():
    app.logger.exception('Error tak terduga')
    return jsonify({'error': MSG_UMUM}), 500


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/provinsi')
def api_provinsi():
    tahun = request.args.get('tahun', '2025')
    if not TAHUN_RE.match(tahun):
        return _bad_request(_pesan_tahun(tahun))
    try:
        return jsonify(sikd_client.get_provinsi(tahun))
    except DJPKError as e:
        return _djpk_gagal(e)
    except Exception:
        return _error_umum()


@app.route('/api/pemda')
def api_pemda():
    provinsi = request.args.get('provinsi', '')
    tahun = request.args.get('tahun', '2025')
    if not KODE_RE.match(provinsi):
        return _bad_request('Pilih provinsi dulu.')
    if not TAHUN_RE.match(tahun):
        return _bad_request(_pesan_tahun(tahun))
    try:
        return jsonify(sikd_client.get_pemda(provinsi, tahun))
    except DJPKError as e:
        return _djpk_gagal(e)
    except Exception:
        return _error_umum()


@app.route('/api/apbd')
def api_apbd():
    tahun = request.args.get('tahun', '2025')
    provinsi = request.args.get('provinsi', '--')
    pemda = request.args.get('pemda', '--')
    periode = request.args.get('periode', '12')
    err = _validasi([tahun], provinsi, pemda, periode)
    if err:
        return _bad_request(err)
    try:
        return jsonify(sikd_client.get_apbd(tahun, provinsi, pemda, periode))
    except DJPKError as e:
        return _djpk_gagal(e)
    except Exception:
        return _error_umum()


@app.route('/api/compare')
def api_compare():
    tahun_list = _parse_tahun_list(request.args.get('tahun', ''))  # 2023,2024,2025
    provinsi = request.args.get('provinsi', '--')
    pemda = request.args.get('pemda', '--')
    periode = request.args.get('periode', '12')
    err = _validasi(tahun_list, provinsi, pemda, periode)
    if err:
        return _bad_request(err)
    try:
        return jsonify(sikd_client.get_apbd_compare(tahun_list, provinsi, pemda, periode))
    except Exception:
        return _error_umum()


@app.route('/api/download-excel')
def api_download_excel():
    """Download data APBD sebagai file Excel asli."""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    tahun = request.args.get('tahun', '2025')
    provinsi = request.args.get('provinsi', '--')
    pemda = request.args.get('pemda', '--')
    periode = request.args.get('periode', '12')
    mode = request.args.get('mode', 'single')  # 'single' atau 'compare'
    tahun_list_str = request.args.get('tahun_list', '')
    filename = re.sub(r'[^A-Za-z0-9_-]', '_', request.args.get('filename', ''))[:120] or 'APBD'

    is_compare = mode == 'compare'
    tahun_list = _parse_tahun_list(tahun_list_str) if is_compare else [tahun]
    err = _validasi(tahun_list, provinsi, pemda, periode)
    if err:
        return _bad_request(err)

    # Ambil data dulu; kalau gagal/kosong, kirim pesan JSON (bukan file Excel kosong / halaman 500 polos)
    try:
        if is_compare:
            all_data = sikd_client.get_apbd_compare(tahun_list, provinsi, pemda, periode)
            if not any(d.get('rows') for d in all_data.values()):
                pesan = next((d['error'] for d in all_data.values() if d.get('error')), None)
                return jsonify({'error': pesan or 'Data tidak tersedia untuk tahun-tahun yang dipilih.'}), 404
        else:
            data = sikd_client.get_apbd(tahun, provinsi, pemda, periode)
            if not data.get('rows'):
                return jsonify({'error': data.get('error') or 'Data tidak tersedia untuk parameter yang dipilih.'}), 404
    except DJPKError as e:
        return _djpk_gagal(e)
    except Exception:
        return _error_umum()

    wb = openpyxl.Workbook()
    ws = wb.active

    # Styles
    hdr_fill = PatternFill(start_color="4c6ef5", end_color="4c6ef5", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    lvl0_fill = PatternFill(start_color="DBE4FF", end_color="DBE4FF", fill_type="solid")
    lvl0_font = Font(bold=True, size=10, name="Arial")
    lvl1_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    lvl1_font = Font(bold=True, size=10, name="Arial")
    data_font = Font(size=10, name="Arial")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    if is_compare:
        ws.title = "Perbandingan APBD"
        years = sorted(all_data.keys())

        # Collect all unique akun
        all_akun = []
        akun_seen = set()
        for y in years:
            for r in all_data[y].get('rows', []):
                if r['akun'] not in akun_seen:
                    akun_seen.add(r['akun'])
                    all_akun.append({'akun': r['akun'], 'level': r['level']})

        # Header
        headers = ['Akun']
        for y in years:
            headers += [f'Anggaran {y}', f'Realisasi {y}', f'% {y}']
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.fill, c.font, c.border = hdr_fill, hdr_font, thin_border
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Rows
        for i, a in enumerate(all_akun, 2):
            ws.cell(row=i, column=1, value=a['akun']).border = thin_border
            col = 2
            for y in years:
                row_data = next((r for r in all_data[y].get('rows', []) if r['akun'] == a['akun']), None)
                if row_data:
                    ws.cell(row=i, column=col, value=row_data.get('anggaran', 0)).border = thin_border
                    ws.cell(row=i, column=col+1, value=row_data.get('realisasi', 0)).border = thin_border
                    try:
                        pct = float(row_data.get('persentase', '0'))
                        ws.cell(row=i, column=col+2, value=pct/100).border = thin_border
                        ws.cell(row=i, column=col+2).number_format = '0.00%'
                    except (ValueError, TypeError):
                        ws.cell(row=i, column=col+2, value=row_data.get('persentase', '')).border = thin_border
                else:
                    for offset in range(3):
                        ws.cell(row=i, column=col+offset, value='-').border = thin_border
                col += 3

            # Apply styling based on level
            for col_idx in range(1, len(headers)+1):
                cell = ws.cell(row=i, column=col_idx)
                cell.font = data_font
                if a['level'] == 0:
                    cell.fill = lvl0_fill
                    cell.font = lvl0_font
                elif a['level'] == 1:
                    cell.fill = lvl1_fill
                    cell.font = lvl1_font
                if col_idx >= 2 and col_idx % 3 != 1:  # Number columns
                    cell.number_format = '#,##0'

        # Column widths
        ws.column_dimensions['A'].width = 50
        for i in range(len(years)):
            col_letter_base = openpyxl.utils.get_column_letter(2 + i*3)
            ws.column_dimensions[col_letter_base].width = 18
            ws.column_dimensions[openpyxl.utils.get_column_letter(3 + i*3)].width = 18
            ws.column_dimensions[openpyxl.utils.get_column_letter(4 + i*3)].width = 10

    else:
        # Single year mode
        ws.title = f"APBD {tahun}"
        rows = data.get('rows', [])

        headers = ['Akun', 'Anggaran', 'Realisasi', '%']
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=col_idx, value=h)
            c.fill, c.font, c.border = hdr_fill, hdr_font, thin_border
            c.alignment = Alignment(horizontal='center', vertical='center')

        for i, r in enumerate(rows, 2):
            ws.cell(row=i, column=1, value=r.get('akun', '')).border = thin_border
            ws.cell(row=i, column=2, value=r.get('anggaran', 0)).border = thin_border
            ws.cell(row=i, column=3, value=r.get('realisasi', 0)).border = thin_border
            try:
                pct = float(r.get('persentase', '0'))
                ws.cell(row=i, column=4, value=pct/100).border = thin_border
                ws.cell(row=i, column=4).number_format = '0.00%'
            except (ValueError, TypeError):
                ws.cell(row=i, column=4, value=r.get('persentase', '')).border = thin_border

            # Styling by level
            for col_idx in range(1, 5):
                cell = ws.cell(row=i, column=col_idx)
                cell.font = data_font
                if r.get('level') == 0:
                    cell.fill = lvl0_fill
                    cell.font = lvl0_font
                elif r.get('level') == 1:
                    cell.fill = lvl1_fill
                    cell.font = lvl1_font
                if col_idx in (2, 3):
                    cell.number_format = '#,##0'

        ws.column_dimensions['A'].width = 50
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 10

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{filename}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


atexit.register(sikd_client.close)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)

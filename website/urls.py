from flask import Blueprint, render_template, request, send_file
import json
from .utils import is_supported_file, is_valid_jadwal, find_asisten, find_patners, all_schedules, generate_excel

urls = Blueprint('web', __name__)


@urls.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@urls.route('/proses', methods=['POST'])
def proses():
    file = request.files.get('file')
    npm_asisten = request.form.get('npm_asisten', '').strip()

    if not file or not file.filename:
        return render_template('jadwal.html', isDocx=False, error_msg='File belum dipilih.')

    is_supported = is_supported_file(file)
    if not is_supported:
        return render_template(
            'jadwal.html',
            isDocx=False,
            error_msg='Format file tidak didukung. Gunakan file .pdf, .docx, atau .xlsx',
        )

    is_valid = is_valid_jadwal(file)
    if not is_valid:
        return render_template(
            'jadwal.html',
            isDocx=True,
            isValid=False,
            error_msg='File yang dimasukkan bukan Jadwal Mengawas Asisten atau tabel tidak terbaca.',
        )

    schedules = all_schedules(file)
    cek_asisten = find_asisten(schedules, npm_asisten)
    target_asisten = next(
        (
            a
            for a in schedules
            if str(a.get('npm', '')).strip().lower() == npm_asisten.lower()
        ),
        None,
    )
    jadwal = target_asisten['jadwal'] if target_asisten else []

    if not cek_asisten or not jadwal:
        return render_template(
            'jadwal.html',
            isDocx=True,
            isValid=True,
            cekAsisten=False,
            error_msg=f"Asisten dengan NPM '{npm_asisten}' tidak ditemukan di file jadwal.",
        )

    patners = find_patners(schedules, npm_asisten)

    def filled(x):
        return (x or '').strip() not in ('', '-')

    total_sesi = sum(sum(1 for s in j['sesi'] if filled(s)) for j in jadwal)
    return render_template(
        'jadwal.html',
        isDocx=True,
        isValid=True,
        cekAsisten=True,
        jadwal=jadwal or [],
        patners=patners or [],
        nama=(target_asisten.get('nama', '') if target_asisten else '').title(),
        npm=npm_asisten,
        total_sesi=total_sesi,
    )


@urls.route('/download_excel', methods=['POST'])
def download_excel():
    jadwal = json.loads(request.form['jadwal'])
    patners = json.loads(request.form['patners'])
    nama = request.form['nama']

    output = generate_excel(jadwal, patners, nama)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"Jadwal_{nama.title().replace(' ', '_')}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

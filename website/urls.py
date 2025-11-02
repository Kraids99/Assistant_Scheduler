from flask import Blueprint, render_template, request, send_file, session, redirect, url_for
from io import BytesIO
import json
from .utils import is_docx, is_valid_jadwal, read_file, find_asisten, find_patners, all_schedules, generate_excel
from docx import Document

urls = Blueprint('web', __name__)

@urls.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@urls.route('/proses', methods=['POST'])
def proses():
    file = request.files['file']
    
    nama_asisten = request.form.get('nama_asisten', '').strip().lower()

    isDocx = is_docx(file)

    if not isDocx:
        return render_template('jadwal.html', isDocx=False, error_msg="Format file tidak didukung. Gunakan file .docx")

    isValid = is_valid_jadwal(file)

    if not isValid:
        return render_template('jadwal.html', isDocx=True, isValid=False, error_msg="File yang dimasukkan bukan Jadwal Mengawas Asisten.")

    docx_file = read_file(file)

    cekAsisten = find_asisten(docx_file, nama_asisten)
    schedules = all_schedules(docx_file)
    jadwal = next((a["jadwal"] for a in schedules if a["nama"].lower() == nama_asisten), [])

    if not cekAsisten or not jadwal:
        return render_template('jadwal.html', isDocx=True, isValid=True, cekAsisten=False, error_msg=f"Asisten bernama '{nama_asisten.title()}' tidak ditemukan di file jadwal.")

    patners = find_patners(schedules, nama_asisten)
    def filled(x): 
        return (x or '').strip() not in ('', '-')

    total_sesi = sum(sum(1 for s in j["sesi"] if filled(s)) for j in jadwal)
    return render_template('jadwal.html', isDocx=True, isValid=True, cekAsisten=True ,jadwal=jadwal or [],patners=patners or [],nama=nama_asisten.title(), total_sesi=total_sesi)

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
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
from flask import Blueprint, render_template, request, send_file, session
import os
from .utils import read_file, find_asisten, find_patners, all_schedules
from werkzeug.utils import secure_filename

urls = Blueprint('web', __name__)

UPLOAD_FOLDER = 'uploads/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@urls.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@urls.route('/proses', methods=['POST'])
def proses():
    file = request.files['file']

    nama_asisten = request.form.get('nama_asisten', '').strip().lower()

    docx_file = read_file(file)
    cekAsisten = find_asisten(docx_file, nama_asisten)
    schedules = all_schedules(docx_file)
    jadwal = next((a["jadwal"] for a in schedules if a["nama"].lower() == nama_asisten.lower()), [])

    patners = find_patners(schedules, nama_asisten)
    
    return render_template('index.html',cekAsisten=cekAsisten,jadwal=jadwal,patners=patners,nama=nama_asisten.title())
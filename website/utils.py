from docx import Document
import io
from io import BytesIO
import re
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def read_file(input_file):
    name = input_file.filename.lower()

    if name.endswith('.docx'):
        input_file.seek(0)
        return Document(io.BytesIO(input_file.read()))
    raise ValueError("Format file tidak didukung.")

def find_asisten(input_file, nama_asisten):
    for table in input_file.tables:
        for row in table.rows:
            row_text = " ".join(cell.text.strip().lower() for cell in row.cells)
            if nama_asisten.lower() in row_text:
                return True
    return False

bulan_pattern = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|Jan|Feb|Mar|Apr|Mei|Jun|Jul|Agu|Sep|Okt|Nov|Des)", re.IGNORECASE)

def all_schedules(input_file):
    all_schedules = []
    bulan_pattern = re.compile(
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|"
        r"Jan|Feb|Mar|Apr|Mei|Jun|Jul|Agu|Sep|Okt|Nov|Des)",
        re.IGNORECASE
    )

    for table in input_file.tables:
        # cari baris tanggal
        tanggal_header = []
        for r in table.rows:
            cells = [c.text.strip() for c in r.cells]
            if any(bulan_pattern.search(c) for c in cells):
                tanggal_header = [t for t in cells if bulan_pattern.search(t)]
                tanggal_header = list(dict.fromkeys(tanggal_header))
                break

        if not tanggal_header:
            continue  # lewati tabel tanpa tanggal

        # proses semua baris nama asisten
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 6 or not cells[1] or cells[1].lower() == "nama":
                continue

            nama = cells[1].lower()
            ruangan_data = cells[5:]
            if not ruangan_data:
                continue

            per_tanggal = len(ruangan_data) // len(tanggal_header)
            jadwal = []

            for i, tanggal in enumerate(tanggal_header):
                start = i * per_tanggal
                end = start + 4
                sesi_data = ruangan_data[start:end]
                sesi_fix = [s if s else "-" for s in sesi_data]
                jadwal.append({
                    "tanggal": tanggal,
                    "sesi": sesi_fix
                })

            all_schedules.append({
                "nama": nama,
                "jadwal": jadwal
            })

    return all_schedules

def find_patners(all_schedules, target_name):
    target = next((a for a in all_schedules if a["nama"].lower() == target_name.lower()), None)
    if not target:
        return []

    hasil = []
    for j in target["jadwal"]:
        tanggal = j["tanggal"]
        for sesi_idx, ruang in enumerate(j["sesi"]):
            if ruang == "-" or not ruang:
                continue

            patners = []
            # cek asisten lain
            for other in all_schedules:
                if other["nama"].lower() == target_name.lower():
                    continue
                for oj in other["jadwal"]:
                    if oj["tanggal"] == tanggal and oj["sesi"][sesi_idx] == ruang:
                        patners.append(other["nama"])
                        break

            hasil.append({
                "ruangan": f"{tanggal} | Sesi {sesi_idx+1} | {ruang}",
                "patners": patners if patners else ["-"]
            })

    return hasil

def generate_excel(jadwal, patners, nama):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Jadwal Ngawas {nama[0].title()}"

    fill_header = PatternFill("solid", fgColor="DCE6F1")
    fill_sesi = PatternFill("solid", fgColor="FCE4D6")
    border_medium = Border(
        left=Side(style="medium", color="000000"),
        right=Side(style="medium", color="000000"),
        top=Side(style="medium", color="000000"),
        bottom=Side(style="medium", color="000000")
    )
    center = Alignment(horizontal="center", vertical="center")

    # === Dinamis: jumlah kolom berdasarkan tanggal ===
    total_tanggal = len(jadwal)
    total_kolom = total_tanggal + 1  # kolom pertama = Sesi/Tanggal
    last_col_letter = get_column_letter(total_kolom)

    # === Judul utama ===
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = f"Jadwal Ngawas {nama.title()}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A1"].alignment = center

    # === Header jadwal ===
    ws["A2"] = "Sesi / Tanggal"
    ws["A2"].fill = fill_header
    ws["A2"].font = Font(bold=True)
    ws["A2"].alignment = center
    ws["A2"].border = border_medium

    for i, j in enumerate(jadwal, start=2):
        c = ws.cell(row=2, column=i, value=j["tanggal"])
        c.fill = fill_header
        c.font = Font(bold=True)
        c.alignment = center
        c.border = border_medium

    # === Isi jadwal ===
    for s_idx in range(4):
        sesi_cell = ws.cell(row=3 + s_idx, column=1, value=f"{s_idx + 1}")
        sesi_cell.fill = fill_sesi
        sesi_cell.border = border_medium
        sesi_cell.alignment = center
        for c_idx, j in enumerate(jadwal, start=2):
            val = j["sesi"][s_idx]
            cell = ws.cell(row=3 + s_idx, column=c_idx, value=val)
            cell.border = border_medium
            cell.alignment = center

    # === Tabel Patners ===
    start_row = 9
    ws.merge_cells(f"A{start_row}:D{start_row}")
    ws[f"A{start_row}"] = "Daftar Patners"
    ws[f"A{start_row}"].font = Font(bold=True, size=14)
    ws[f"A{start_row}"].alignment = center

    ws[f"A{start_row+1}"] = "Ruangan"
    ws.merge_cells(f"B{start_row+1}:D{start_row+1}")
    ws[f"B{start_row+1}"] = "Patners"
    for col in ["A", "B", "c", "D"]:
        cell = ws[f"{col}{start_row+1}"]
        cell.fill = fill_header
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.border = border_medium

    for i, p in enumerate(patners, start=start_row + 2):
        ws[f"A{i}"] = p["ruangan"]
        ws[f"A{i}"].fill = fill_sesi
        ws.merge_cells(f"B{i}:D{i}")
        ws[f"B{i}"] = ", ".join(p["patners"]) if p["patners"][0] != "-" else "-"
        for col in ["A", "B", "C", "D"]:
            cell = ws[f"{col}{i}"]
            cell.border = border_medium
            cell.alignment = center

    # === Lebar kolom otomatis ===
    for col in range(1, total_kolom + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 20

    ws.column_dimensions["A"].width = 30

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# def schedule_table(input_file, nama_asisten):
#     """
#     Membaca tabel jadwal ngawas (format UTS besar).
#     Output contoh:
#     [{'tanggal': '20-Oct', 'sesi': ['3327', '-', '-', '-']}, ...]
#     """
#     result = []
#     target_row = None
#     tanggal_header = []

#     for table in input_file.tables:
#         # cari baris yang mengandung semua tanggal
#         for r in table.rows:
#             cells = [c.text.strip() for c in r.cells]
#             if any(bulan_pattern.search(c) for c in cells):
#                 # ambil semua tanggal unik
#                 tanggal_header = [t for t in cells if bulan_pattern.search(t)]
#                 tanggal_header = list(dict.fromkeys(tanggal_header))
#                 break

#         # cari baris nama asisten
#         for row in table.rows:
#             cells = [c.text.strip() for c in row.cells]
#             if nama_asisten.lower() in " ".join(cells).lower():
#                 target_row = cells
#                 break

#         if target_row:
#             break

#     if not target_row or not tanggal_header:
#         return []

#     # hitung offset kolom tanggal
#     # di file kamu: kolom 0–4 = No, Nama, NPM, Prodi → mulai tanggal di kolom ke-5
#     start_col = 5  
#     ruangan_data = target_row[start_col:]

#     # Bersihkan tanggal duplikat & kosong
#     tanggal_header = [t for t in tanggal_header if t.strip()]
#     tanggal_header = list(dict.fromkeys(tanggal_header))

#     # Bagi data per tanggal (4 sesi per tanggal)
#     per_tanggal = len(ruangan_data) // len(tanggal_header)
#     for i, tanggal in enumerate(tanggal_header):
#         start = i * per_tanggal
#         end = start + 4
#         sesi_data = ruangan_data[start:end]
#         sesi_fix = [s if s else "-" for s in sesi_data]
#         result.append({"tanggal": tanggal, "sesi": sesi_fix})

#     return result
from docx import Document
import io
import re

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
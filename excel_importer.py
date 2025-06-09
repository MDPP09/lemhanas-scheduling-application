import pandas as pd
from datetime import datetime
import sqlite3 # Diperlukan untuk create_table di bagian __main__ untuk pengujian

# Fungsi import_activities_from_excel
def import_activities_from_excel(file_path):
    imported_count = 0
    failed_count = 0
    errors = []

    try:
        df = pd.read_excel(file_path, header=0, dtype=str)

        # Mapping header Excel ke nama kolom database kita
        column_mapping = {
            'TANGGAL': 'tanggal_kegiatan',
            # 'WAKTU' sekarang akan diurai menjadi dua kolom DB: waktu_mulai_kegiatan dan waktu_akhir_kegiatan
            'KEGIATAN': 'uraian_kegiatan',
            'TEMPAT/RUANGAN': 'tempat_ruangan',
            'PIMPINAN': 'pimpinan',
            'PELAKSANA/PESERTA': 'daftar_peserta',
            'TGL INPUT': 'tanggal_input',
            'WKT INPUT': 'waktu_input',
            'PIC': 'narahubung',
            'KONTAK PERSON': 'kontak_person'
        }

        for row_idx, row_series in df.iterrows():
            excel_row_number = row_idx + 2 

            data = {}
            for excel_col, db_col in column_mapping.items():
                cell_value = str(row_series.get(excel_col, '')).strip()
                if cell_value.lower() == 'nan':
                    cell_value = ''
                data[db_col] = cell_value

            # --- Tangani kolom WAKTU dari Excel untuk waktu_mulai_kegiatan dan waktu_akhir_kegiatan ---
            # Asumsi: Kolom 'WAKTU' di Excel berisi "HH:MM - HH:MM"
            excel_time_raw = str(row_series.get('WAKTU', '')).strip()
            
            if not excel_time_raw: # Skip if no time data
                errors.append(f"Baris {excel_row_number}: Kolom WAKTU kosong. Baris dilewati.")
                failed_count += 1
                continue

            try:
                # Pisahkan string waktu "HH:MM - HH:MM"
                start_time_str, end_time_str = excel_time_raw.split('-')
                data['waktu_mulai_kegiatan'] = start_time_str.strip()
                data['waktu_akhir_kegiatan'] = end_time_str.strip()
            except ValueError:
                errors.append(f"Baris {excel_row_number}: Format WAKTU '{excel_time_raw}' tidak valid. Harap gunakan 'HH:MM - HH:MM'.")
                failed_count += 1
                continue

            # Lewati baris jika kolom kegiatan utama (tanggal, waktu_mulai, waktu_akhir, uraian) kosong
            if not all([data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], data['waktu_akhir_kegiatan'], data['uraian_kegiatan']]):
                errors.append(f"Baris {excel_row_number}: Data utama (Tanggal, Waktu, Uraian) tidak lengkap.")
                failed_count += 1
                continue

            # Auto-fill 'tanggal_input' dan 'waktu_input' jika tidak ada di Excel
            if not data.get('tanggal_input'):
                data['tanggal_input'] = datetime.now().strftime('%Y-%m-%d')
            if not data.get('waktu_input'):
                data['waktu_input'] = datetime.now().strftime('%H:%M')

            # --- Validasi dan Konversi Format Tanggal Kegiatan ---
            try:
                # Asumsi: Format tanggal di Excel adalah DD-MM-YYYY (misal: 01-05-2025)
                tanggal_kegiatan_obj = datetime.strptime(data['tanggal_kegiatan'], '%d-%m-%Y').date()
                data['tanggal_kegiatan'] = tanggal_kegiatan_obj.strftime('%Y-%m-%d')
            except ValueError as ve:
                errors.append(f"Baris {excel_row_number}: Format TANGGAL '{data.get('tanggal_kegiatan')}' tidak valid. (Harap pakai DD-MM-YYYY). Error: {ve}")
                failed_count += 1
                continue 
            
            # --- Validasi Format Waktu Mulai dan Waktu Akhir ---
            try:
                start_time_obj = datetime.strptime(data['waktu_mulai_kegiatan'], '%H:%M').time()
                end_time_obj = datetime.strptime(data['waktu_akhir_kegiatan'], '%H:%M').time()
                if start_time_obj >= end_time_obj:
                    errors.append(f"Baris {excel_row_number}: Waktu Mulai ({data['waktu_mulai_kegiatan']}) harus lebih awal dari Waktu Akhir ({data['waktu_akhir_kegiatan']}).")
                    failed_count += 1
                    continue
            except ValueError as ve:
                errors.append(f"Baris {excel_row_number}: Format Waktu Mulai atau Waktu Akhir tidak valid. (Harap pakai HH:MM). Error: {ve}")
                failed_count += 1
                continue

            from db_handler import add_activity 
            success, message = add_activity(data)
            
            if success:
                imported_count += 1
            else:
                failed_count += 1
                errors.append(f"Baris {excel_row_number}: {message} (Kegiatan: {data.get('uraian_kegiatan', 'N/A')})")
    
    except FileNotFoundError:
        errors.append("File Excel tidak ditemukan.")
    except pd.errors.EmptyDataError:
        errors.append("File Excel kosong atau tidak memiliki data yang valid.")
    except Exception as e:
        errors.append(f"Terjadi kesalahan saat membaca file Excel: {e}")
    
    return imported_count, failed_count, errors

if __name__ == '__main__':
    import os
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Ganti dengan path ke file Excel yang Anda ingin uji
    test_excel_path = os.path.join(current_dir, 'Jadwal Giat Pimpinan Lemhannas RI.xlsx - JUNI 2025.xlsx')

    conn = sqlite3.connect("schedule.db")
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Kegiatan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_kegiatan TEXT NOT NULL,
            waktu_mulai_kegiatan TEXT NOT NULL,
            waktu_akhir_kegiatan TEXT NOT NULL,
            uraian_kegiatan TEXT,
            tempat_ruangan TEXT,
            pimpinan TEXT,
            daftar_peserta TEXT,
            tanggal_input TEXT,
            waktu_input TEXT,
            narahubung TEXT,
            kontak_person TEXT
        )
    ''')
    conn.commit()
    conn.close()

    print(f"Mencoba mengimpor dari: {test_excel_path}")
    imported, failed, errs = import_activities_from_excel(test_excel_path)
    print(f"Import Selesai. Berhasil: {imported}, Gagal: {failed}")
    if errs:
        print("Detail Error:")
        for err in errs:
            print(err)
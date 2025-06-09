import sqlite3
from datetime import datetime, timedelta

DATABASE_NAME = "schedule.db"

def connect_db():
    conn = sqlite3.connect(DATABASE_NAME)
    return conn

def create_table():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Kegiatan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_kegiatan TEXT NOT NULL,
            waktu_mulai_kegiatan TEXT NOT NULL,  -- Kolom baru
            waktu_akhir_kegiatan TEXT NOT NULL,  -- Kolom baru
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

# Fungsi bantu untuk memeriksa tumpang tindih waktu
def is_time_overlap(start_time1_str, end_time1_str, start_time2_str, end_time2_str):
    fmt = '%H:%M'
    start1 = datetime.strptime(start_time1_str, fmt).time()
    end1 = datetime.strptime(end_time1_str, fmt).time()
    start2 = datetime.strptime(start_time2_str, fmt).time()
    end2 = datetime.strptime(end_time2_str, fmt).time()

    # Normalisasi waktu akhir jika melewati tengah malam (misal 23:00 - 01:00)
    # Untuk kesederhanaan, kita asumsikan kegiatan tidak melewati tengah malam.
    # Jika perlu melewati tengah malam, logika ini harus lebih kompleks.
    if start1 > end1: end1 = datetime.strptime("23:59", fmt).time() # Untuk mencegah masalah waktu akhir lebih kecil dari mulai
    if start2 > end2: end2 = datetime.strptime("23:59", fmt).time()

    # Cek tumpang tindih: [start1, end1) dan [start2, end2)
    # Tumpang tindih jika (start1 < end2 AND start2 < end1)
    return start1 < end2 and start2 < end1

def validate_activity_overlap(activity_date, new_start_time, new_end_time, new_pimpinan_raw, new_participants_raw, current_activity_id=None):
    conn = connect_db()
    cursor = conn.cursor()

    # Normalize new pimpinan and participants
    new_pimpinan = set(p.strip().lower() for p in new_pimpinan_raw.split(',') if p.strip())
    new_participants = set(p.strip().lower() for p in new_participants_raw.split(',') if p.strip())

    # Ambil semua kegiatan pada tanggal yang sama
    query = "SELECT id, waktu_mulai_kegiatan, waktu_akhir_kegiatan, pimpinan, daftar_peserta FROM Kegiatan WHERE tanggal_kegiatan = ?"
    params = [activity_date]
    
    if current_activity_id is not None:
        query += " AND id != ?" # Kecualikan kegiatan yang sedang diedit
        params.append(current_activity_id)

    cursor.execute(query, params)
    clashing_activities = cursor.fetchall()
    conn.close()

    for activity_row in clashing_activities:
        existing_id, existing_start_time, existing_end_time, existing_pimpinan_raw, existing_participants_raw = activity_row

        # Normalize existing pimpinan and participants
        existing_pimpinan = set(p.strip().lower() for p in existing_pimpinan_raw.split(',') if p.strip())
        existing_participants = set(p.strip().lower() for p in existing_participants_raw.split(',') if p.strip())

        # Check for time overlap
        if is_time_overlap(new_start_time, new_end_time, existing_start_time, existing_end_time):
            # Check for pimpinan overlap
            if not new_pimpinan.isdisjoint(existing_pimpinan):
                return False, f"Pimpinan '{new_pimpinan_raw}' sudah terjadwal pada waktu tersebut ({existing_start_time}-{existing_end_time})."
            
            # Check for participants overlap (if any common participants)
            if not new_participants.isdisjoint(existing_participants):
                return False, f"Beberapa peserta sudah terjadwal pada waktu tersebut ({existing_start_time}-{existing_end_time})."
    
    return True, "" # No overlap found

def add_activity(data):
    conn = connect_db()
    cursor = conn.cursor()

    # Validasi tumpang tindih sebelum insert
    is_valid, message = validate_activity_overlap(
        data['tanggal_kegiatan'],
        data['waktu_mulai_kegiatan'],
        data['waktu_akhir_kegiatan'],
        data['pimpinan'],
        data['daftar_peserta']
    )
    if not is_valid:
        conn.close()
        return False, message

    try:
        cursor.execute('''
            INSERT INTO Kegiatan (
                tanggal_kegiatan, waktu_mulai_kegiatan, waktu_akhir_kegiatan, uraian_kegiatan,
                tempat_ruangan, pimpinan, daftar_peserta, tanggal_input, waktu_input,
                narahubung, kontak_person
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], data['waktu_akhir_kegiatan'],
            data['uraian_kegiatan'], data['tempat_ruangan'], data['pimpinan'],
            data['daftar_peserta'], data['tanggal_input'], data['waktu_input'],
            data['narahubung'], data['kontak_person']
        ))
        conn.commit()
        conn.close()
        return True, "Kegiatan berhasil ditambahkan!"
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat menambahkan kegiatan: {e}"

def get_activity_by_id(activity_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Kegiatan WHERE id = ?", (activity_id,))
    activity = cursor.fetchone()
    conn.close()
    return activity

def update_activity(activity_id, data):
    conn = connect_db()
    cursor = conn.cursor()

    # Validasi tumpang tindih sebelum update
    is_valid, message = validate_activity_overlap(
        data['tanggal_kegiatan'],
        data['waktu_mulai_kegiatan'],
        data['waktu_akhir_kegiatan'],
        data['pimpinan'],
        data['daftar_peserta'],
        current_activity_id=activity_id # Penting: Kecualikan kegiatan yang sedang diedit
    )
    if not is_valid:
        conn.close()
        return False, message

    try:
        cursor.execute('''
            UPDATE Kegiatan SET
                tanggal_kegiatan = ?,
                waktu_mulai_kegiatan = ?,
                waktu_akhir_kegiatan = ?,
                uraian_kegiatan = ?,
                tempat_ruangan = ?,
                pimpinan = ?,
                daftar_peserta = ?,
                narahubung = ?,
                kontak_person = ?
            WHERE id = ?
        ''', (
            data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], data['waktu_akhir_kegiatan'],
            data['uraian_kegiatan'], data['tempat_ruangan'], data['pimpinan'],
            data['daftar_peserta'], data['narahubung'], data['kontak_person'],
            activity_id
        ))
        conn.commit()
        conn.close()
        return True, "Kegiatan berhasil diperbarui!"
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat memperbarui kegiatan: {e}"

def get_all_activities():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Kegiatan ORDER BY tanggal_kegiatan, waktu_mulai_kegiatan")
    activities = cursor.fetchall()
    conn.close()
    return activities

def delete_activity(activity_id):
    conn = connect_db()
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM Kegiatan WHERE id = ?", (activity_id,))
        conn.commit()
        conn.close()
        return True, "Kegiatan berhasil dihapus."
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat menghapus kegiatan: {e}"

if __name__ == '__main__':
    create_table()
    print("Database and table 'Kegiatan' created/checked.")
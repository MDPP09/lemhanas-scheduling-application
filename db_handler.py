import sqlite3
from datetime import datetime, timedelta
import random # Untuk warna acak awal

DATABASE_NAME = "schedule.db"

def connect_db():
    conn = sqlite3.connect(DATABASE_NAME)
    # Configure row_factory to return rows as dictionaries for easier column access
    conn.row_factory = sqlite3.Row
    return conn

def create_table():
    conn = connect_db()
    cursor = conn.cursor()

    # Tabel Pimpinan (BARU)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Pimpinan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT NOT NULL UNIQUE,
            warna TEXT NOT NULL
        )
    ''')

    # Tabel Kegiatan (Diubah: pimpinan -> id_pimpinan)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Kegiatan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal_kegiatan TEXT NOT NULL,
            waktu_mulai_kegiatan TEXT NOT NULL,
            waktu_akhir_kegiatan TEXT NOT NULL,
            uraian_kegiatan TEXT,
            tempat_ruangan TEXT,
            id_pimpinan INTEGER, -- Diubah dari TEXT ke INTEGER
            daftar_peserta TEXT,
            tanggal_input TEXT,
            waktu_input TEXT,
            narahubung TEXT,
            kontak_person TEXT,
            FOREIGN KEY (id_pimpinan) REFERENCES Pimpinan(id) ON DELETE SET NULL
        )
    ''')
    conn.commit()
    conn.close()

# --- Fungsi Manajemen Pimpinan (BARU) ---
def get_all_pimpinan():
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Pimpinan ORDER BY nama")
    pimpinan_list = cursor.fetchall()
    conn.close()
    return pimpinan_list

def add_pimpinan(nama_pimpinan):
    conn = connect_db()
    cursor = conn.cursor()
    try:
        # Generate random hex color
        color = '#%06x' % random.randint(0, 0xFFFFFF)
        cursor.execute("INSERT INTO Pimpinan (nama, warna) VALUES (?, ?)", (nama_pimpinan, color))
        conn.commit()
        conn.close()
        return True, "Pimpinan berhasil ditambahkan!", cursor.lastrowid
    except sqlite3.IntegrityError:
        conn.close()
        return False, "Nama pimpinan sudah ada.", None
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat menambahkan pimpinan: {e}", None

def delete_pimpinan(pimpinan_id):
    conn = connect_db()
    cursor = conn.cursor()
    try:
        # Check if there are any activities associated with this pimpinan
        cursor.execute("SELECT COUNT(*) FROM Kegiatan WHERE id_pimpinan = ?", (pimpinan_id,))
        if cursor.fetchone()[0] > 0:
            # If activities are found, set id_pimpinan to NULL before deleting pimpinan
            cursor.execute("UPDATE Kegiatan SET id_pimpinan = NULL WHERE id_pimpinan = ?", (pimpinan_id,))
            conn.commit() # Commit the update before deleting pimpinan
            
        cursor.execute("DELETE FROM Pimpinan WHERE id = ?", (pimpinan_id,))
        conn.commit()
        conn.close()
        return True, "Pimpinan berhasil dihapus!"
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat menghapus pimpinan: {e}"

def get_pimpinan_by_id(pimpinan_id):
    conn = connect_db()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Pimpinan WHERE id = ?", (pimpinan_id,))
    pimpinan = cursor.fetchone()
    conn.close()
    return pimpinan

def update_pimpinan_color(pimpinan_id, new_color):
    conn = connect_db()
    cursor = conn.cursor()
    try:
        cursor.execute("UPDATE Pimpinan SET warna = ? WHERE id = ?", (new_color, pimpinan_id))
        conn.commit()
        conn.close()
        return True, "Warna pimpinan berhasil diperbarui!"
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat memperbarui warna pimpinan: {e}"

# --- Fungsi Bantu Validasi Waktu ---
def is_time_overlap(start_time1_str, end_time1_str, start_time2_str, end_time2_str):
    fmt = '%H:%M'
    start1 = datetime.strptime(start_time1_str, fmt).time()
    end1 = datetime.strptime(end_time1_str, fmt).time()
    start2 = datetime.strptime(start_time2_str, fmt).time()
    end2 = datetime.strptime(end_time2_str, fmt).time()

    # Handle cases where end time is on the next day (e.g., 23:00 - 01:00)
    # For simplicity, we assume activities don't span across midnight in the current context.
    # If they do, the logic here needs to be more complex, possibly involving datetime objects with dates.
    # For now, if end time is before start time, treat it as end of day for overlap check
    if end1 < start1:
        end1 = datetime.strptime("23:59", fmt).time() # Treat as end of day for overlap check
    if end2 < start2:
        end2 = datetime.strptime("23:59", fmt).time()

    return start1 < end2 and start2 < end1

def validate_activity_overlap(activity_date, new_start_time, new_end_time, id_pimpinan, new_participants_raw, current_activity_id=None):
    conn = connect_db()
    cursor = conn.cursor()

    # Get pimpinan name for validation message (if needed)
    pimpinan_name = ""
    if id_pimpinan: # Ensure id_pimpinan is not None/empty
        pimpinan_row = get_pimpinan_by_id(id_pimpinan)
        if pimpinan_row:
            pimpinan_name = pimpinan_row['nama']
    
    new_participants = set(p.strip().lower() for p in new_participants_raw.split(',') if p.strip())

    # Ambil semua kegiatan pada tanggal yang sama
    query = "SELECT id, waktu_mulai_kegiatan, waktu_akhir_kegiatan, id_pimpinan, daftar_peserta FROM Kegiatan WHERE tanggal_kegiatan = ?"
    params = [activity_date]
    
    if current_activity_id is not None:
        query += " AND id != ?" # Kecualikan kegiatan yang sedang diedit
        params.append(current_activity_id)

    cursor.execute(query, params)
    clashing_activities = cursor.fetchall()
    conn.close()

    for activity_row in clashing_activities:
        existing_id = activity_row['id']
        existing_start_time = activity_row['waktu_mulai_kegiatan']
        existing_end_time = activity_row['waktu_akhir_kegiatan']
        existing_id_pimpinan = activity_row['id_pimpinan']
        existing_participants_raw = activity_row['daftar_peserta']

        existing_participants = set(p.strip().lower() for p in existing_participants_raw.split(',') if p.strip())

        # Check for time overlap
        if is_time_overlap(new_start_time, new_end_time, existing_start_time, existing_end_time):
            # Check for pimpinan overlap (if same pimpinan)
            if id_pimpinan is not None and existing_id_pimpinan == id_pimpinan:
                return False, f"Pimpinan '{pimpinan_name}' sudah terjadwal pada waktu tersebut ({existing_start_time}-{existing_end_time})."
            
            # Check for participants overlap (if any common participants)
            if new_participants and not new_participants.isdisjoint(existing_participants):
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
        data.get('id_pimpinan'),
        data['daftar_peserta']
    )
    if not is_valid:
        conn.close()
        return False, message

    try:
        cursor.execute('''
            INSERT INTO Kegiatan (
                tanggal_kegiatan, waktu_mulai_kegiatan, waktu_akhir_kegiatan, uraian_kegiatan,
                tempat_ruangan, id_pimpinan, daftar_peserta, tanggal_input, waktu_input,
                narahubung, kontak_person
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], data['waktu_akhir_kegiatan'],
            data['uraian_kegiatan'], data['tempat_ruangan'], data.get('id_pimpinan'),
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
    # Join with Pimpinan table to get pimpinan name and color
    cursor.execute('''
        SELECT
            K.id, K.tanggal_kegiatan, K.waktu_mulai_kegiatan, K.waktu_akhir_kegiatan,
            K.uraian_kegiatan, K.tempat_ruangan, K.id_pimpinan,
            K.daftar_peserta, K.tanggal_input, K.waktu_input, K.narahubung, K.kontak_person,
            P.nama AS pimpinan_nama, P.warna AS pimpinan_warna
        FROM Kegiatan AS K
        LEFT JOIN Pimpinan AS P ON K.id_pimpinan = P.id
        WHERE K.id = ?
    ''', (activity_id,))
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
        data.get('id_pimpinan'),
        data['daftar_peserta'],
        current_activity_id=activity_id
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
                id_pimpinan = ?,
                daftar_peserta = ?,
                narahubung = ?,
                kontak_person = ?
            WHERE id = ?
        ''', (
            data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], data['waktu_akhir_kegiatan'],
            data['uraian_kegiatan'], data['tempat_ruangan'], data.get('id_pimpinan'),
            data['daftar_peserta'], data['narahubung'], data['kontak_person'],
            activity_id
        ))
        conn.commit()
        conn.close()
        return True, "Kegiatan berhasil diperbarui!"
    except sqlite3.Error as e:
        conn.close()
        return False, f"Error saat memperbarui kegiatan: {e}"

# Fungsi untuk mengambil semua kegiatan, ditambah informasi pimpinan
def get_all_activities(id_pimpinan_filter=None):
    conn = connect_db()
    cursor = conn.cursor()
    
    query = '''
        SELECT
            K.id, K.tanggal_kegiatan, K.waktu_mulai_kegiatan, K.waktu_akhir_kegiatan,
            K.uraian_kegiatan, K.tempat_ruangan, P.nama AS pimpinan_nama,
            K.daftar_peserta, K.tanggal_input, K.waktu_input, K.narahubung, K.kontak_person,
            P.id AS id_pimpinan, P.warna AS pimpinan_warna
        FROM Kegiatan AS K
        LEFT JOIN Pimpinan AS P ON K.id_pimpinan = P.id
    '''
    params = []

    if id_pimpinan_filter is not None:
        query += " WHERE K.id_pimpinan = ?"
        params.append(id_pimpinan_filter)

    query += " ORDER BY K.tanggal_kegiatan, K.waktu_mulai_kegiatan"
    
    # DEBUG PRINT: Cetak query dan parameter
    print(f"DEBUG DB: Executing query: {query}")
    print(f"DEBUG DB: With parameters: {params}")

    cursor.execute(query, params)
    activities = cursor.fetchall() # Returns Row objects (dictionary-like)
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
    print("Database and tables 'Kegiatan' and 'Pimpinan' created/checked.")
    # Add some dummy pimpinan for testing
    add_pimpinan("Gubernur")
    add_pimpinan("Sekretaris")
    add_pimpinan("Kadiv Umum")
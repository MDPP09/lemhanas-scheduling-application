import customtkinter as ctk
from datetime import datetime, timedelta
from db_handler import create_table, add_activity, get_all_activities, delete_activity, get_activity_by_id, update_activity
from excel_importer import import_activities_from_excel
from tkinter import filedialog, messagebox
from tkcalendar import Calendar, DateEntry
import threading
import time

# CustomTkinter appearance settings
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Aplikasi Penjadwalan Kegiatan Instansi")
        self.geometry("1400x800")

        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        create_table()

        self.create_widgets()
        self.notification_thread = None
        self._notified_activities = set()
        self.start_notification_checker()
        self.update_calendar_markers()

    def create_widgets(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Menu Aplikasi", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.add_activity_button = ctk.CTkButton(self.sidebar_frame, text="Tambah Kegiatan", command=self.open_add_activity_form)
        self.add_activity_button.grid(row=1, column=0, padx=20, pady=10)

        self.import_excel_button = ctk.CTkButton(self.sidebar_frame, text="Import dari Excel", command=self.import_excel_dialog)
        self.import_excel_button.grid(row=2, column=0, padx=20, pady=10)

        self.refresh_button = ctk.CTkButton(self.sidebar_frame, text="Refresh Jadwal & Kalender", command=self.refresh_all)
        self.refresh_button.grid(row=3, column=0, padx=20, pady=10)

        self.main_content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_content_frame.grid_rowconfigure(1, weight=1)
        self.main_content_frame.grid_columnconfigure(0, weight=1)

        self.calendar_frame = ctk.CTkFrame(self.main_content_frame)
        self.calendar_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.calendar_frame.grid_columnconfigure(0, weight=1)

        self.calendar_label = ctk.CTkLabel(self.calendar_frame, text="Kalender Kegiatan", font=ctk.CTkFont(size=20, weight="bold"))
        self.calendar_label.pack(pady=(0, 10))

        self.calendar = Calendar(self.calendar_frame, selectmode='day',
                                 date_pattern='y-mm-dd',
                                 font="Arial 12",
                                 tooltipdelay=100,
                                 showweeknumbers=False,
                                 cursor="hand1")
        self.calendar.pack(expand=True, fill="both")
        self.calendar.bind("<<CalendarSelected>>", self.on_date_selected)

        self.activity_list_frame = ctk.CTkFrame(self.main_content_frame)
        self.activity_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.activity_list_frame.grid_rowconfigure(1, weight=1)
        self.activity_list_frame.grid_columnconfigure(0, weight=1)

        self.activity_list_label = ctk.CTkLabel(self.activity_list_frame, text="Kegiatan pada Tanggal Dipilih:", font=ctk.CTkFont(size=18, weight="bold"))
        self.activity_list_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.activity_display_frame = ctk.CTkScrollableFrame(self.activity_list_frame)
        self.activity_display_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.activity_display_frame.grid_columnconfigure(0, weight=1)

        self.activity_labels = []

        self.on_date_selected(event=None)

    def on_date_selected(self, event):
        selected_date = self.calendar.get_date()
        self.load_activities_for_date(selected_date)

    def load_activities_for_date(self, date_str):
        for widget in self.activity_display_frame.winfo_children():
            widget.destroy()
        self.activity_labels = []

        all_activities = get_all_activities()
        filtered_activities = [act for act in all_activities if act[1] == date_str]

        self.activity_list_label.configure(text=f"Kegiatan pada Tanggal: {date_str}")

        if not filtered_activities:
            no_data_label = ctk.CTkLabel(self.activity_display_frame, text="Tidak ada kegiatan untuk tanggal ini.", font=ctk.CTkFont(size=16))
            no_data_label.grid(row=0, column=0, padx=20, pady=20)
            return

        # Headers disesuaikan untuk waktu mulai dan akhir
        headers = ["ID", "Tgl Kegiatan", "Mulai", "Akhir", "Uraian", "Tempat", "Pimpinan", "Peserta", "Narahubung", "Kontak", "Aksi"]
        # Indeks kolom DB: (id, tanggal_kegiatan, waktu_mulai_kegiatan, waktu_akhir_kegiatan, uraian_kegiatan, tempat_ruangan, pimpinan, daftar_peserta, tanggal_input, waktu_input, narahubung, kontak_person)
        display_indices = [0, 1, 2, 3, 4, 5, 6, 7, 10, 11] # Sesuaikan indeks kolom DB

        for col_idx, header_text in enumerate(headers):
            header_label = ctk.CTkLabel(self.activity_display_frame, text=header_text, font=ctk.CTkFont(weight="bold"))
            header_label.grid(row=0, column=col_idx, padx=5, pady=5, sticky="w")
            if header_text == "Uraian":
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=3)
            elif header_text == "Aksi":
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=0, minsize=140)
            else:
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=1)

        for row_idx, activity in enumerate(filtered_activities):
            displayed_data = [activity[i] for i in display_indices]
            for col_idx, item in enumerate(displayed_data):
                data_label = ctk.CTkLabel(self.activity_display_frame, text=str(item))
                data_label.grid(row=row_idx + 1, column=col_idx, padx=5, pady=2, sticky="w")
                self.activity_labels.append(data_label)

            action_frame = ctk.CTkFrame(self.activity_display_frame, fg_color="transparent")
            action_frame.grid(row=row_idx + 1, column=len(headers) - 1, padx=5, pady=2, sticky="ew")
            action_frame.grid_columnconfigure(0, weight=1)
            action_frame.grid_columnconfigure(1, weight=1)

            edit_button = ctk.CTkButton(action_frame, text="Edit",
                                        command=lambda a_id=activity[0]: self.open_edit_activity_form(a_id),
                                        width=60, fg_color="gray", hover_color="#696969")
            edit_button.grid(row=0, column=0, padx=(0, 5), sticky="w")

            delete_button = ctk.CTkButton(action_frame, text="Hapus",
                                        command=lambda a_id=activity[0]: self.confirm_delete_activity(a_id),
                                        width=60, fg_color="red", hover_color="#8b0000")
            delete_button.grid(row=0, column=1, sticky="w")


    def update_calendar_markers(self):
        self.calendar.calevent_remove('all')

        all_activities = get_all_activities()
        activity_dates = {}
        for activity in all_activities:
            date_str = activity[1]
            if date_str not in activity_dates:
                activity_dates[date_str] = []
            activity_dates[date_str].append(activity)

        for date_str, activities in activity_dates.items():
            try:
                dt_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                tooltip_text = "Kegiatan:\n" + "\n".join([f"- {a[2]} - {a[3]} ({a[4]})" for a in activities]) # Waktu_mulai - waktu_akhir (uraian)
                self.calendar.calevent_create(dt_obj, tooltip_text, 'activity')
            except ValueError:
                print(f"Invalid date format found in DB: {date_str}")
        self.calendar.tag_config('activity', background='lightblue', foreground='black')


    def open_add_activity_form(self):
        add_form = AddActivityForm(self)
        self.wait_window(add_form)
        self.refresh_all()

    def open_edit_activity_form(self, activity_id):
        activity_data = get_activity_by_id(activity_id)
        if activity_data:
            edit_form = EditActivityForm(self, activity_data)
            self.wait_window(edit_form)
            self.refresh_all()
        else:
            messagebox.showerror("Error", "Kegiatan tidak ditemukan.")

    def confirm_delete_activity(self, activity_id):
        if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus kegiatan ID {activity_id}?"):
            success, message = delete_activity(activity_id)
            if success:
                messagebox.showinfo("Berhasil", message)
                self.refresh_all()
            else:
                messagebox.showerror("Error", message)

    def import_excel_dialog(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            imported, failed, errors = import_activities_from_excel(file_path)
            if imported > 0:
                messagebox.showinfo("Import Selesai", f"Berhasil mengimpor {imported} kegiatan.\nGagal: {failed}")
                if errors:
                    error_msg = "\n".join(errors)
                    messagebox.showwarning("Peringatan Impor", f"Beberapa kegiatan gagal diimpor:\n{error_msg}")
            else:
                messagebox.showwarning("Import Gagal", f"Tidak ada kegiatan yang berhasil diimpor. Total gagal: {failed}")
                if errors:
                    error_msg = "\n".join(errors)
                    messagebox.showerror("Error Impor", f"Detail Error:\n{error_msg}")
            self.refresh_all()

    def refresh_all(self):
        self.load_activities_for_date(self.calendar.get_date())
        self.update_calendar_markers()
        self._notified_activities.clear()

    def start_notification_checker(self):
        if self.notification_thread is None or not self.notification_thread.is_alive():
            self.notification_thread = threading.Thread(target=self._check_for_notifications, daemon=True)
            self.notification_thread.start()

    def _check_for_notifications(self):
        while True:
            self.check_upcoming_activities()
            time.sleep(60)

    def check_upcoming_activities(self):
        now = datetime.now()
        all_activities = get_all_activities()

        for activity in all_activities:
            try:
                # activity tuple: (id, tanggal_kegiatan, waktu_mulai_kegiatan, waktu_akhir_kegiatan, ...)
                activity_start_datetime_str = f"{activity[1]} {activity[2]}"
                activity_start_dt_obj = datetime.strptime(activity_start_datetime_str, '%Y-%m-%d %H:%M')

                # Gunakan waktu mulai kegiatan untuk notifikasi
                notification_time = activity_start_dt_obj - timedelta(minutes=15)

                if now >= notification_time and now < activity_start_dt_obj:
                    if activity[0] not in self._notified_activities:
                        self.show_notification(activity)
                        self._notified_activities.add(activity[0])

            except ValueError as e:
                print(f"Error parsing date/time for activity ID {activity[0]}: {e}")
            except Exception as e:
                print(f"An unexpected error occurred during notification check: {e}")

    def show_notification(self, activity_data):
        self.after(0, lambda: self._display_notification_popup(activity_data))

    def _display_notification_popup(self, activity_data):
        activity_id = activity_data[0]
        activity_uraian = activity_data[4] # Indeks berubah
        activity_start_time = activity_data[2] # Waktu mulai
        activity_end_time = activity_data[3]   # Waktu akhir
        activity_place = activity_data[5] # Indeks berubah

        messagebox.showinfo(
            "Peringatan Kegiatan Mendatang!",
            f"Kegiatan: {activity_uraian}\n"
            f"Waktu: {activity_start_time} - {activity_end_time}\n"
            f"Tempat: {activity_place}\n\n"
            f"Kegiatan akan segera dimulai!"
        )

    def on_closing(self):
        self.destroy()


class AddActivityForm(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Tambah Kegiatan Baru")
        self.geometry("600x680") # Sesuaikan tinggi
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.create_form_widgets()

    def create_form_widgets(self):
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        labels = [
            "Tanggal Kegiatan (YYYY-MM-DD):",
            "Waktu MULAI (HH:MM):",     # Label baru
            "Waktu AKHIR (HH:MM):",     # Label baru
            "Uraian Kegiatan:",
            "Tempat Ruangan:",
            "Pimpinan:",
            "Daftar Peserta/Jumlah:",
            "Narahubung:",
            "Kontak Person:",
        ]

        self.entries = {}
        # Mapping label ke key dictionary entries
        label_to_entry_key = {
            "Tanggal Kegiatan (YYYY-MM-DD):": 'tanggal_kegiatan',
            "Waktu MULAI (HH:MM):": 'waktu_mulai_kegiatan',
            "Waktu AKHIR (HH:MM):": 'waktu_akhir_kegiatan',
            "Uraian Kegiatan:": 'uraian_kegiatan',
            "Tempat Ruangan:": 'tempat_ruangan',
            "Pimpinan:": 'pimpinan',
            "Daftar Peserta/Jumlah:": 'daftar_peserta',
            "Narahubung:": 'narahubung',
            "Kontak Person:": 'kontak_person',
        }

        for i, label_text in enumerate(labels):
            label = ctk.CTkLabel(self.frame, text=label_text)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            entry_key = label_to_entry_key[label_text]

            if entry_key == 'tanggal_kegiatan':
                self.entries[entry_key] = DateEntry(self.frame, width=20, background='darkblue', foreground='white', bordercolor='darkblue',
                                                            headersbackground='darkblue', headersforeground='white', selectbackground='lightblue',
                                                            selectforeground='black', normalbackground='lightgray', normalforeground='black',
                                                            locale='id_ID', date_pattern='yyyy-mm-dd')
                self.entries[entry_key].grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key].set_date(datetime.now().date())
            else:
                entry = ctk.CTkEntry(self.frame, width=300)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = entry
        
        # Set default values for time inputs
        if 'waktu_mulai_kegiatan' in self.entries:
            self.entries['waktu_mulai_kegiatan'].insert(0, datetime.now().strftime('%H:%M'))
        if 'waktu_akhir_kegiatan' in self.entries:
            # Default waktu akhir bisa 1 jam setelah waktu mulai
            default_end_time = (datetime.now() + timedelta(hours=1)).strftime('%H:%M')
            self.entries['waktu_akhir_kegiatan'].insert(0, default_end_time)


        self.add_button = ctk.CTkButton(self.frame, text="Simpan Kegiatan", command=self.save_activity)
        self.add_button.grid(row=len(labels), column=0, columnspan=2, pady=20)
    
    def save_activity(self):
        data = {
            'tanggal_kegiatan': self.entries['tanggal_kegiatan'].get_date().strftime('%Y-%m-%d'),
            'waktu_mulai_kegiatan': self.entries['waktu_mulai_kegiatan'].get(),
            'waktu_akhir_kegiatan': self.entries['waktu_akhir_kegiatan'].get(),
            'uraian_kegiatan': self.entries['uraian_kegiatan'].get(),
            'tempat_ruangan': self.entries['tempat_ruangan'].get(),
            'pimpinan': self.entries['pimpinan'].get(),
            'daftar_peserta': self.entries['daftar_peserta'].get(),
            'tanggal_input': datetime.now().strftime('%Y-%m-%d'),
            'waktu_input': datetime.now().strftime('%H:%M'),
            'narahubung': self.entries['narahubung'].get(),
            'kontak_person': self.entries['kontak_person'].get()
        }

        # Basic validation
        if not all([data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], 
                      data['waktu_akhir_kegiatan'], data['uraian_kegiatan']]):
            messagebox.showerror("Input Error", "Tanggal Kegiatan, Waktu Mulai, Waktu Akhir, dan Uraian Kegiatan tidak boleh kosong.")
            return
        
        try:
            datetime.strptime(data['tanggal_kegiatan'], '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Input Error", "Format Tanggal Kegiatan salah. Gunakan YYYY-MM-DD.")
            return
        
        # Validate time formats and ensure start time is before end time
        try:
            start_time_obj = datetime.strptime(data['waktu_mulai_kegiatan'], '%H:%M').time()
            end_time_obj = datetime.strptime(data['waktu_akhir_kegiatan'], '%H:%M').time()
            if start_time_obj >= end_time_obj:
                messagebox.showerror("Input Error", "Waktu Mulai harus lebih awal dari Waktu Akhir.")
                return
        except ValueError:
            messagebox.showerror("Input Error", "Format Waktu Mulai atau Waktu Akhir salah. Gunakan HH:MM.")
            return

        success, message = add_activity(data)
        if success:
            messagebox.showinfo("Berhasil", message)
            self.destroy()
        else:
            messagebox.showerror("Error", message)

    def on_closing(self):
        self.master.focus_set()
        self.destroy()

class EditActivityForm(ctk.CTkToplevel):
    def __init__(self, master, activity_data):
        super().__init__(master)
        self.title("Edit Kegiatan")
        self.geometry("600x680") # Sesuaikan tinggi
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.activity_id = activity_data[0]
        self.activity_data = activity_data # (id, tanggal_kegiatan, waktu_mulai, waktu_akhir, uraian, tempat, pimpinan, peserta, tgl_input, wkt_input, narahubung, kontak_person)

        self.create_form_widgets()
        self.load_activity_data()

    def create_form_widgets(self):
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        labels = [
            "Tanggal Kegiatan (YYYY-MM-DD):",
            "Waktu MULAI (HH:MM):",
            "Waktu AKHIR (HH:MM):",
            "Uraian Kegiatan:",
            "Tempat Ruangan:",
            "Pimpinan:",
            "Daftar Peserta/Jumlah:",
            "Narahubung:",
            "Kontak Person:",
        ]

        self.entries = {}
        label_to_entry_key = {
            "Tanggal Kegiatan (YYYY-MM-DD):": 'tanggal_kegiatan',
            "Waktu MULAI (HH:MM):": 'waktu_mulai_kegiatan',
            "Waktu AKHIR (HH:MM):": 'waktu_akhir_kegiatan',
            "Uraian Kegiatan:": 'uraian_kegiatan',
            "Tempat Ruangan:": 'tempat_ruangan',
            "Pimpinan:": 'pimpinan',
            "Daftar Peserta/Jumlah:": 'daftar_peserta',
            "Narahubung:": 'narahubung',
            "Kontak Person:": 'kontak_person',
        }

        for i, label_text in enumerate(labels):
            label = ctk.CTkLabel(self.frame, text=label_text)
            label.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            
            entry_key = label_to_entry_key[label_text]

            if entry_key == 'tanggal_kegiatan':
                self.entries[entry_key] = DateEntry(self.frame, width=20, background='darkblue', foreground='white', bordercolor='darkblue',
                                                            headersbackground='darkblue', headersforeground='white', selectbackground='lightblue',
                                                            selectforeground='black', normalbackground='lightgray', normalforeground='black',
                                                            locale='id_ID', date_pattern='yyyy-mm-dd')
                self.entries[entry_key].grid(row=i, column=1, padx=10, pady=5, sticky="ew")
            else:
                entry = ctk.CTkEntry(self.frame, width=300)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = entry
        
        self.save_button = ctk.CTkButton(self.frame, text="Update Kegiatan", command=self.update_activity)
        self.save_button.grid(row=len(labels), column=0, columnspan=2, pady=20)

    def load_activity_data(self):
        # activity_data tuple (id, tanggal_kegiatan, waktu_mulai, waktu_akhir, uraian, tempat, pimpinan, peserta, tgl_input, wkt_input, narahubung, kontak_person)
        self.entries['tanggal_kegiatan'].set_date(datetime.strptime(self.activity_data[1], '%Y-%m-%d').date())
        self.entries['waktu_mulai_kegiatan'].insert(0, self.activity_data[2])
        self.entries['waktu_akhir_kegiatan'].insert(0, self.activity_data[3])
        self.entries['uraian_kegiatan'].insert(0, self.activity_data[4])
        self.entries['tempat_ruangan'].insert(0, self.activity_data[5])
        self.entries['pimpinan'].insert(0, self.activity_data[6])
        self.entries['daftar_peserta'].insert(0, self.activity_data[7])
        self.entries['narahubung'].insert(0, self.activity_data[10])
        self.entries['kontak_person'].insert(0, self.activity_data[11])

    def update_activity(self):
        data = {
            'tanggal_kegiatan': self.entries['tanggal_kegiatan'].get_date().strftime('%Y-%m-%d'),
            'waktu_mulai_kegiatan': self.entries['waktu_mulai_kegiatan'].get(),
            'waktu_akhir_kegiatan': self.entries['waktu_akhir_kegiatan'].get(),
            'uraian_kegiatan': self.entries['uraian_kegiatan'].get(),
            'tempat_ruangan': self.entries['tempat_ruangan'].get(),
            'pimpinan': self.entries['pimpinan'].get(),
            'daftar_peserta': self.entries['daftar_peserta'].get(),
            'narahubung': self.entries['narahubung'].get(),
            'kontak_person': self.entries['kontak_person'].get()
        }

        # Basic validation
        if not all([data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], 
                      data['waktu_akhir_kegiatan'], data['uraian_kegiatan']]):
            messagebox.showerror("Input Error", "Tanggal Kegiatan, Waktu Mulai, Waktu Akhir, dan Uraian Kegiatan tidak boleh kosong.")
            return
        
        try:
            datetime.strptime(data['tanggal_kegiatan'], '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Input Error", "Format Tanggal Kegiatan salah. Gunakan YYYY-MM-DD.")
            return

        # Validate time formats and ensure start time is before end time
        try:
            start_time_obj = datetime.strptime(data['waktu_mulai_kegiatan'], '%H:%M').time()
            end_time_obj = datetime.strptime(data['waktu_akhir_kegiatan'], '%H:%M').time()
            if start_time_obj >= end_time_obj:
                messagebox.showerror("Input Error", "Waktu Mulai harus lebih awal dari Waktu Akhir.")
                return
        except ValueError:
            messagebox.showerror("Input Error", "Format Waktu Mulai atau Waktu Akhir salah. Gunakan HH:MM.")
            return

        success, message = update_activity(self.activity_id, data)
        if success:
            messagebox.showinfo("Berhasil", message)
            self.destroy()
        else:
            messagebox.showerror("Error", message)

    def on_closing(self):
        self.master.focus_set()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
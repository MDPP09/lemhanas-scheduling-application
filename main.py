from tkcalendar import Calendar, DateEntry
import customtkinter as ctk
from datetime import datetime, timedelta, date
import calendar as py_calendar
from tkinter import filedialog, messagebox, colorchooser
import threading
import time
from PIL import ImageTk
import os

# Import functions from db_handler (updated)
from db_handler import (
    create_table, add_activity, get_all_activities, delete_activity,
    get_activity_by_id, update_activity,
    get_all_pimpinan, add_pimpinan, delete_pimpinan, get_pimpinan_by_id,
    update_pimpinan_color
)
# Import from excel_importer
from excel_importer import import_activities_from_excel

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Aplikasi Penjadwalan Kegiatan Lemhanas")
        self.geometry("1400x800")
        self.iconpath = ImageTk.PhotoImage(file=os.path.join("assets","Logo_Lembaga_Ketahanan_Nasional.png"))
        self.wm_iconbitmap()
        self.iconphoto(False, self.iconpath)

        self.grid_columnconfigure(0, weight=0) # Sidebar
        self.grid_columnconfigure(1, weight=1) # Main content
        self.grid_rowconfigure(0, weight=1)

        create_table() # Ensure database tables exist
        self.pimpinan_data = {} # To store pimpinan ID -> name mapping
        self.pimpinan_colors = {} # To store pimpinan ID -> color mapping
        self._load_pimpinan_data() # Load pimpinan at startup

        self.current_filter_id_pimpinan = None # Default: show all pimpinan

        self.create_widgets()
        self.notification_thread = None
        self._notified_activities = set()
        self.start_notification_checker()
        
        # Initial load: display activities for the currently selected date (default: today)
        self.load_activities_for_date(self.calendar.get_date())
        self.update_calendar_markers() # Initial markers for activities

    def _load_pimpinan_data(self):
        self.pimpinan_data = {}
        self.pimpinan_colors = {}
        all_pimpinan = get_all_pimpinan()
        for p in all_pimpinan:
            self.pimpinan_data[p['id']] = p['nama']
            self.pimpinan_colors[p['id']] = p['warna']

    def create_widgets(self):
        # Sidebar Frame
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1) # Adjust row configure for new buttons

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="Menu Aplikasi", font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.add_activity_button = ctk.CTkButton(self.sidebar_frame, text="Tambah Kegiatan", command=self.open_add_activity_form)
        self.add_activity_button.grid(row=1, column=0, padx=20, pady=10)

        self.import_excel_button = ctk.CTkButton(self.sidebar_frame, text="Import dari Excel", command=self.import_excel_dialog)
        self.import_excel_button.grid(row=2, column=0, padx=20, pady=10)

        self.manage_pimpinan_button = ctk.CTkButton(self.sidebar_frame, text="Kelola Pimpinan", command=self.open_manage_pimpinan_form)
        self.manage_pimpinan_button.grid(row=3, column=0, padx=20, pady=10)
        
        # Filter Section
        self.filter_label = ctk.CTkLabel(self.sidebar_frame, text="Filter Pimpinan:", font=ctk.CTkFont(size=14, weight="bold"))
        self.filter_label.grid(row=4, column=0, padx=20, pady=(10, 5), sticky="w")

        pimpinan_names_for_filter = ["Semua Pimpinan"] + [p['nama'] for p in get_all_pimpinan()]
        self.pimpinan_filter_combobox = ctk.CTkComboBox(self.sidebar_frame, values=pimpinan_names_for_filter,
                                                        command=self._apply_pimpinan_filter)
        self.pimpinan_filter_combobox.set("Semua Pimpinan")
        self.pimpinan_filter_combobox.grid(row=5, column=0, padx=20, pady=5, sticky="ew")

        self.refresh_button = ctk.CTkButton(self.sidebar_frame, text="Refresh Jadwal & Kalender", command=self.refresh_all)
        self.refresh_button.grid(row=6, column=0, padx=20, pady=10)

        # Main Content Frame
        self.main_content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.main_content_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_content_frame.grid_rowconfigure(1, weight=1) # Calendar takes minimal, activity list takes the rest
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
        # Bind events to update the daily activity view
        self.calendar.bind("<<CalendarSelected>>", self.on_date_selected)
        # We also need to refresh markers when month changes, but not the activity list below
        self.calendar.bind("<<CalendarMonthChanged>>", self.on_month_changed) 

        # Bottom section for Activities on Selected Date (REVERTED TO OLDER STRUCTURE)
        self.activity_list_frame = ctk.CTkFrame(self.main_content_frame)
        self.activity_list_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.activity_list_frame.grid_rowconfigure(1, weight=1) # For scrollable frame
        self.activity_list_frame.grid_columnconfigure(0, weight=1)

        self.activity_list_label = ctk.CTkLabel(self.activity_list_frame, text="Kegiatan pada Tanggal Dipilih:", font=ctk.CTkFont(size=18, weight="bold"))
        self.activity_list_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

        self.activity_display_frame = ctk.CTkScrollableFrame(self.activity_list_frame)
        self.activity_display_frame.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.activity_display_frame.grid_columnconfigure(0, weight=1)

        self.activity_labels = [] # To store references to activity data labels


    # --- Filter Pimpinan Logic ---
    def _apply_pimpinan_filter(self, selected_pimpinan_name):
        self.current_filter_id_pimpinan = None
        if selected_pimpinan_name != "Semua Pimpinan":
            # Compare names in a case-insensitive manner
            for p_id, p_name in self.pimpinan_data.items():
                if p_name.lower() == selected_pimpinan_name.lower():
                    self.current_filter_id_pimpinan = p_id
                    break
        self.refresh_all() # Refresh the view with the new filter

    # --- New/Modified Functions for Calendar View & Activity List ---
    def on_month_changed(self, event=None):
        # When month changes, only update calendar markers, not the daily activity list below
        self.update_calendar_markers()

    def on_date_selected(self, event=None):
        # When a date is selected, load activities for that specific date
        selected_date = self.calendar.get_date()
        self.load_activities_for_date(selected_date)


    def load_activities_for_date(self, date_str):
        # Clear existing labels
        for widget in self.activity_display_frame.winfo_children():
            widget.destroy()
        self.activity_labels = []

        all_activities = get_all_activities(id_pimpinan_filter=self.current_filter_id_pimpinan) # Apply filter
        filtered_activities = [act for act in all_activities if act['tanggal_kegiatan'] == date_str] # Access by key

        self.activity_list_label.configure(text=f"Kegiatan pada Tanggal: {date_str}")

        if not filtered_activities:
            no_data_label = ctk.CTkLabel(self.activity_display_frame, text="Tidak ada kegiatan untuk tanggal ini.", font=ctk.CTkFont(size=16))
            no_data_label.grid(row=0, column=0, padx=20, pady=20)
            return

        # Create header row
        headers = ["ID", "Waktu", "Uraian Kegiatan", "Tempat", "Pimpinan", "Aksi"]
        
        # Configure columns for header (adjust width weights for buttons)
        for col_idx, header_text in enumerate(headers):
            header_label = ctk.CTkLabel(self.activity_display_frame, text=header_text, font=ctk.CTkFont(weight="bold"))
            header_label.grid(row=0, column=col_idx, padx=5, pady=5, sticky="w")
            if header_text == "Uraian Kegiatan":
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=3)
            elif header_text == "Aksi":
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=0, minsize=140) # Space for 2 buttons
            else:
                self.activity_display_frame.grid_columnconfigure(col_idx, weight=1)

        # Sort activities by start time
        filtered_activities.sort(key=lambda x: datetime.strptime(x['waktu_mulai_kegiatan'], '%H:%M').time())

        for row_idx, activity in enumerate(filtered_activities):
            # activity is a Row object (dictionary-like)
            activity_id = activity['id']
            activity_time_range = f"{activity['waktu_mulai_kegiatan']} - {activity['waktu_akhir_kegiatan']}"
            activity_uraian = activity['uraian_kegiatan']
            activity_tempat = activity['tempat_ruangan']
            pimpinan_name = activity['pimpinan_nama']
            pimpinan_color = activity['pimpinan_warna']

            # Display cells
            id_label = ctk.CTkLabel(self.activity_display_frame, text=str(activity_id))
            id_label.grid(row=row_idx + 1, column=0, padx=5, pady=2, sticky="w")

            time_label = ctk.CTkLabel(self.activity_display_frame, text=activity_time_range)
            time_label.grid(row=row_idx + 1, column=1, padx=5, pady=2, sticky="w")

            uraian_label = ctk.CTkLabel(self.activity_display_frame, text=activity_uraian)
            uraian_label.grid(row=row_idx + 1, column=2, padx=5, pady=2, sticky="w")

            tempat_label = ctk.CTkLabel(self.activity_display_frame, text=activity_tempat)
            tempat_label.grid(row=row_idx + 1, column=3, padx=5, pady=2, sticky="w")

            pimpinan_label = ctk.CTkLabel(self.activity_display_frame, text=pimpinan_name,
                                           fg_color=pimpinan_color, text_color="black", corner_radius=5)
            pimpinan_label.grid(row=row_idx + 1, column=4, padx=5, pady=2, sticky="w")


            # Action buttons (Edit and Delete)
            action_frame = ctk.CTkFrame(self.activity_display_frame, fg_color="transparent")
            action_frame.grid(row=row_idx + 1, column=5, padx=5, pady=2, sticky="ew")
            action_frame.grid_columnconfigure(0, weight=1)
            action_frame.grid_columnconfigure(1, weight=1)

            edit_button = ctk.CTkButton(action_frame, text="Edit",
                                        command=lambda a_id=activity_id: self.open_edit_activity_form(a_id),
                                        width=60, fg_color="gray", hover_color="#696969")
            edit_button.grid(row=0, column=0, padx=(0, 5), sticky="w")

            delete_button = ctk.CTkButton(action_frame, text="Hapus",
                                        command=lambda a_id=activity_id: self.confirm_delete_activity(a_id),
                                        width=60, fg_color="red", hover_color="#8b0000")
            delete_button.grid(row=0, column=1, sticky="w")

    # The show_activity_context_menu will no longer be called as buttons are direct
    # but I'll keep it just in case you want to re-enable it for some reason.
    def show_activity_context_menu(self, event, activity_id):
        # Create a Toplevel window for the context menu
        menu = ctk.CTkToplevel(self)
        menu.overrideredirect(True) # Make it behave like a popup menu

        # Position the menu at the mouse click location
        menu.geometry(f"+{event.x_root}+{event.y_root}")
        
        # Edit button
        edit_button = ctk.CTkButton(menu, text="Edit", 
                                    command=lambda: (menu.destroy(), self.open_edit_activity_form(activity_id)))
        edit_button.pack(fill="x", padx=5, pady=2)

        # Delete button
        delete_button = ctk.CTkButton(menu, text="Hapus", 
                                      command=lambda: (menu.destroy(), self.confirm_delete_activity(activity_id)),
                                      fg_color="red", hover_color="#8b0000")
        delete_button.pack(fill="x", padx=5, pady=2)

        # Bind a click outside to close the menu
        def close_menu_on_click_outside(event):
            if not (menu.winfo_exists() and menu.winfo_containing(event.x_root, event.y_root) == menu):
                menu.destroy()
        self.bind("<Button-1>", close_menu_on_click_outside, add="+")
        menu.bind("<Button-1>", lambda e: "break", add="+") # Prevent click on menu from closing it immediately
        
        menu.focus_set()


    def update_calendar_markers(self):
        self.calendar.calevent_remove('all')

        all_activities = get_all_activities(id_pimpinan_filter=self.current_filter_id_pimpinan) # Apply filter to markers
        
        activity_dates = {}
        for activity in all_activities:
            date_str = activity['tanggal_kegiatan']
            if date_str not in activity_dates:
                activity_dates[date_str] = []
            activity_dates[date_str].append(activity)

        for date_str, activities in activity_dates.items():
            try:
                dt_obj = datetime.strptime(date_str, '%Y-%m-%d').date()
                tooltip_text = "Kegiatan:\n" + "\n".join([f"- {a['waktu_mulai_kegiatan']}-{a['waktu_akhir_kegiatan']} {a['uraian_kegiatan']}" for a in activities]) 
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
        self._load_pimpinan_data() # Reload pimpinan data
        # Refresh the activity list for the currently selected date
        self.load_activities_for_date(self.calendar.get_date())
        self.update_calendar_markers() # Refresh markers (will use the current filter)
        self._notified_activities.clear()

    # --- Manage Pimpinan Form (BARU) ---
    def open_manage_pimpinan_form(self):
        manage_form = ManagePimpinanForm(self)
        self.wait_window(manage_form)
        self.refresh_all() # Refresh after managing pimpinan

    # Notification part is mostly the same, only activity data access changes from tuple to dict
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
        all_activities = get_all_activities() # Get all activities, no filter for notifications

        for activity in all_activities:
            try:
                # activity is now a Row object (dictionary-like)
                activity_start_datetime_str = f"{activity['tanggal_kegiatan']} {activity['waktu_mulai_kegiatan']}"
                activity_start_dt_obj = datetime.strptime(activity_start_datetime_str, '%Y-%m-%d %H:%M')

                notification_time = activity_start_dt_obj - timedelta(minutes=15)

                if now >= notification_time and now < activity_start_dt_obj:
                    if activity['id'] not in self._notified_activities:
                        self.show_notification(activity)
                        self._notified_activities.add(activity['id'])

            except ValueError as e:
                print(f"Error parsing date/time for activity ID {activity.get('id', 'N/A')}: {e}")
            except Exception as e:
                print(f"An unexpected error occurred during notification check: {e}")

    def show_notification(self, activity_data):
        self.after(0, lambda: self._display_notification_popup(activity_data))

    def _display_notification_popup(self, activity_data):
        # activity_data is a Row object
        activity_id = activity_data['id']
        activity_uraian = activity_data['uraian_kegiatan']
        activity_start_time = activity_data['waktu_mulai_kegiatan']
        activity_end_time = activity_data['waktu_akhir_kegiatan']
        activity_place = activity_data['tempat_ruangan']

        messagebox.showinfo(
            "Peringatan Kegiatan Mendatang!",
            f"Kegiatan: {activity_uraian}\n"
            f"Waktu: {activity_start_time} - {activity_end_time}\n"
            f"Tempat: {activity_place}\n\n"
            f"Kegiatan akan segera dimulai!"
        )

    def on_closing(self):
        self.destroy()


# --- Form untuk Tambah Kegiatan (Modifikasi untuk ComboBox Pimpinan) ---
class AddActivityForm(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Tambah Kegiatan Baru")
        self.geometry("600x720") # Tinggi disesuaikan
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master_app = master # Reference to the main App instance
        self.pimpinan_options = self.master_app.pimpinan_data # ID -> Name
        self.create_form_widgets()

    def create_form_widgets(self):
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        labels = [
            "Tanggal Kegiatan (YYYY-MM-DD):",
            "Waktu MULAI (HH:MM):",
            "Waktu AKHIR (HH:MM):",
            "Uraian Kegiatan:",
            "Tempat Ruangan:",
            "Pimpinan:", # Label untuk ComboBox
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
            "Pimpinan:": 'id_pimpinan', # Key untuk ComboBox
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
            elif entry_key == 'id_pimpinan': # Custom for Pimpinan ComboBox
                pimpinan_names = sorted(list(self.pimpinan_options.values()))
                self.pimpinan_combobox = ctk.CTkComboBox(self.frame, values=["Pilih Pimpinan"] + pimpinan_names,
                                                            command=self._update_pimpinan_selection)
                self.pimpinan_combobox.set("Pilih Pimpinan")
                self.pimpinan_combobox.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = None # Will store ID directly in save_activity
            else:
                entry = ctk.CTkEntry(self.frame, width=300)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = entry
        
        if 'waktu_mulai_kegiatan' in self.entries:
            self.entries['waktu_mulai_kegiatan'].insert(0, datetime.now().strftime('%H:%M'))
        if 'waktu_akhir_kegiatan' in self.entries:
            # Default waktu akhir bisa 1 jam setelah waktu mulai
            default_end_time = (datetime.now() + timedelta(hours=1)).strftime('%H:%M')
            self.entries['waktu_akhir_kegiatan'].insert(0, default_end_time)


        self.add_button = ctk.CTkButton(self.frame, text="Simpan Kegiatan", command=self.save_activity)
        self.add_button.grid(row=len(labels), column=0, columnspan=2, pady=20)
    
    def _update_pimpinan_selection(self, selected_name):
        self.selected_pimpinan_id = None
        for p_id, p_name in self.pimpinan_options.items():
            if p_name == selected_name:
                self.selected_pimpinan_id = p_id
                break

    def save_activity(self):
        data = {
            'tanggal_kegiatan': self.entries['tanggal_kegiatan'].get_date().strftime('%Y-%m-%d'),
            'waktu_mulai_kegiatan': self.entries['waktu_mulai_kegiatan'].get(),
            'waktu_akhir_kegiatan': self.entries['waktu_akhir_kegiatan'].get(),
            'uraian_kegiatan': self.entries['uraian_kegiatan'].get(),
            'tempat_ruangan': self.entries['tempat_ruangan'].get(),
            'id_pimpinan': self.selected_pimpinan_id, # Ambil ID pimpinan
            'daftar_peserta': self.entries['daftar_peserta'].get(),
            'tanggal_input': datetime.now().strftime('%Y-%m-%d'),
            'waktu_input': datetime.now().strftime('%H:%M'),
            'narahubung': self.entries['narahubung'].get(),
            'kontak_person': self.entries['kontak_person'].get()
        }

        # Basic validation
        if not all([data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], 
                      data['waktu_akhir_kegiatan'], data['uraian_kegiatan'], data['id_pimpinan']]): # id_pimpinan juga wajib
            messagebox.showerror("Input Error", "Tanggal Kegiatan, Waktu Mulai, Waktu Akhir, Uraian Kegiatan, dan Pimpinan tidak boleh kosong.")
            return
        
        try:
            datetime.strptime(data['tanggal_kegiatan'], '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Input Error", "Format Tanggal Kegiatan salah. Gunakan मीडियावर-MM-DD.")
            return
        
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
        self.master_app.focus_set()
        self.destroy()

# --- Form untuk Edit Kegiatan (Modifikasi untuk ComboBox Pimpinan) ---
class EditActivityForm(ctk.CTkToplevel):
    def __init__(self, master, activity_data):
        super().__init__(master)
        self.title("Edit Kegiatan")
        self.geometry("600x720") # Tinggi disesuaikan
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master_app = master
        self.activity_id = activity_data['id']
        self.activity_data = activity_data # This is a Row object (dictionary-like)

        self.pimpinan_options = self.master_app.pimpinan_data # ID -> Name
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
            "Pimpinan:": 'id_pimpinan',
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
            elif entry_key == 'id_pimpinan': # Custom for Pimpinan ComboBox
                pimpinan_names = sorted(list(self.pimpinan_options.values()))
                self.pimpinan_combobox = ctk.CTkComboBox(self.frame, values=pimpinan_names,
                                                            command=self._update_pimpinan_selection)
                self.pimpinan_combobox.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = None # Will store ID directly in update_activity
            else:
                entry = ctk.CTkEntry(self.frame, width=300)
                entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                self.entries[entry_key] = entry
        
        self.save_button = ctk.CTkButton(self.frame, text="Update Kegiatan", command=self.update_activity_data)
        self.save_button.grid(row=len(labels), column=0, columnspan=2, pady=20)

    def _update_pimpinan_selection(self, selected_name):
        self.selected_pimpinan_id = None
        for p_id, p_name in self.pimpinan_options.items():
            if p_name == selected_name:
                self.selected_pimpinan_id = p_id
                break

    def load_activity_data(self):
        self.entries['tanggal_kegiatan'].set_date(datetime.strptime(self.activity_data['tanggal_kegiatan'], '%Y-%m-%d').date())
        self.entries['waktu_mulai_kegiatan'].insert(0, self.activity_data['waktu_mulai_kegiatan'])
        self.entries['waktu_akhir_kegiatan'].insert(0, self.activity_data['waktu_akhir_kegiatan'])
        self.entries['uraian_kegiatan'].insert(0, self.activity_data['uraian_kegiatan'])
        self.entries['tempat_ruangan'].insert(0, self.activity_data['tempat_ruangan'])
        
        # Set ComboBox value based on current pimpinan
        current_pimpinan_id = self.activity_data['id_pimpinan']
        if current_pimpinan_id in self.pimpinan_options:
            self.pimpinan_combobox.set(self.pimpinan_options[current_pimpinan_id])
            self.selected_pimpinan_id = current_pimpinan_id # Initialize selected_pimpinan_id
        else:
            self.pimpinan_combobox.set("Pilih Pimpinan")
            self.selected_pimpinan_id = None


        self.entries['daftar_peserta'].insert(0, self.activity_data['daftar_peserta'])
        self.entries['narahubung'].insert(0, self.activity_data['narahubung'])
        self.entries['kontak_person'].insert(0, self.activity_data['kontak_person'])

    def update_activity_data(self): # Renamed to avoid confusion with db_handler.update_activity
        data = {
            'tanggal_kegiatan': self.entries['tanggal_kegiatan'].get_date().strftime('%Y-%m-%d'),
            'waktu_mulai_kegiatan': self.entries['waktu_mulai_kegiatan'].get(),
            'waktu_akhir_kegiatan': self.entries['waktu_akhir_kegiatan'].get(),
            'uraian_kegiatan': self.entries['uraian_kegiatan'].get(),
            'tempat_ruangan': self.entries['tempat_ruangan'].get(),
            'id_pimpinan': self.selected_pimpinan_id,
            'daftar_peserta': self.entries['daftar_peserta'].get(),
            'narahubung': self.entries['narahubung'].get(),
            'kontak_person': self.entries['kontak_person'].get()
        }

        if not all([data['tanggal_kegiatan'], data['waktu_mulai_kegiatan'], 
                      data['waktu_akhir_kegiatan'], data['uraian_kegiatan'], data['id_pimpinan']]):
            messagebox.showerror("Input Error", "Tanggal Kegiatan, Waktu Mulai, Waktu Akhir, Uraian Kegiatan, dan Pimpinan tidak boleh kosong.")
            return
        
        try:
            datetime.strptime(data['tanggal_kegiatan'], '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Input Error", "Format Tanggal Kegiatan salah. Gunakan मीडियावर-MM-DD.")
            return

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
        self.master_app.focus_set()
        self.destroy()

# --- Form untuk Kelola Pimpinan (BARU) ---
class ManagePimpinanForm(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Kelola Pimpinan")
        self.geometry("500x400")
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.master_app = master
        self.create_widgets()
        self.load_pimpinan_list()

    def create_widgets(self):
        self.frame = ctk.CTkFrame(self)
        self.frame.pack(padx=20, pady=20, fill="both", expand=True)

        self.frame.grid_columnconfigure(0, weight=1)

        self.add_pimpinan_label = ctk.CTkLabel(self.frame, text="Tambah Pimpinan Baru:", font=ctk.CTkFont(weight="bold"))
        self.add_pimpinan_label.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        self.new_pimpinan_entry = ctk.CTkEntry(self.frame, placeholder_text="Nama Pimpinan")
        self.new_pimpinan_entry.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.add_pimpinan_button = ctk.CTkButton(self.frame, text="Tambah", command=self.add_new_pimpinan)
        self.add_pimpinan_button.grid(row=1, column=1, padx=10, pady=5)

        self.pimpinan_list_label = ctk.CTkLabel(self.frame, text="Daftar Pimpinan:", font=ctk.CTkFont(weight="bold"))
        self.pimpinan_list_label.grid(row=2, column=0, padx=10, pady=(15, 5), sticky="w")

        self.pimpinan_list_frame = ctk.CTkScrollableFrame(self.frame, height=200)
        self.pimpinan_list_frame.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="nsew")
        self.pimpinan_list_frame.grid_columnconfigure(0, weight=1) # Name
        self.pimpinan_list_frame.grid_columnconfigure(1, weight=0) # Color display
        self.pimpinan_list_frame.grid_columnconfigure(2, weight=0) # Color picker button
        self.pimpinan_list_frame.grid_columnconfigure(3, weight=0) # Delete button

    def load_pimpinan_list(self):
        for widget in self.pimpinan_list_frame.winfo_children():
            widget.destroy()
        
        pimpinan_list = get_all_pimpinan() # Get current list of pimpinan
        
        if not pimpinan_list:
            no_pimpinan_label = ctk.CTkLabel(self.pimpinan_list_frame, text="Belum ada pimpinan.")
            no_pimpinan_label.grid(row=0, column=0, padx=10, pady=10)
            return

        # Header for the list
        ctk.CTkLabel(self.pimpinan_list_frame, text="Nama Pimpinan", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.pimpinan_list_frame, text="Warna", font=ctk.CTkFont(weight="bold")).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ctk.CTkLabel(self.pimpinan_list_frame, text="Aksi", font=ctk.CTkFont(weight="bold")).grid(row=0, column=2, columnspan=2, padx=5, pady=2, sticky="w")


        for i, pimpinan in enumerate(pimpinan_list):
            pimpinan_id = pimpinan['id']
            pimpinan_nama = pimpinan['nama']
            pimpinan_warna = pimpinan['warna']

            # Name Label
            name_label = ctk.CTkLabel(self.pimpinan_list_frame, text=pimpinan_nama, anchor="w")
            name_label.grid(row=i+1, column=0, padx=5, pady=2, sticky="ew")

            # Color Display
            color_display = ctk.CTkLabel(self.pimpinan_list_frame, text="", bg_color=pimpinan_warna, width=30, height=20)
            color_display.grid(row=i+1, column=1, padx=5, pady=2, sticky="w")

            # Change Color Button
            change_color_btn = ctk.CTkButton(self.pimpinan_list_frame, text="Ganti Warna", width=80,
                                              command=lambda pid=pimpinan_id, c_display=color_display: self.change_pimpinan_color(pid, c_display))
            change_color_btn.grid(row=i+1, column=2, padx=5, pady=2, sticky="w")

            # Delete Button
            delete_btn = ctk.CTkButton(self.pimpinan_list_frame, text="Hapus", width=60, fg_color="red", hover_color="#8b0000",
                                        command=lambda pid=pimpinan_id: self.confirm_delete_pimpinan(pid))
            delete_btn.grid(row=i+1, column=3, padx=5, pady=2, sticky="w")

    def add_new_pimpinan(self):
        new_name = self.new_pimpinan_entry.get().strip()
        if new_name:
            success, message, new_id = add_pimpinan(new_name)
            if success:
                messagebox.showinfo("Berhasil", message)
                self.new_pimpinan_entry.delete(0, ctk.END)
                self.load_pimpinan_list() # Refresh the list
            else:
                messagebox.showerror("Error", message)
        else:
            messagebox.showwarning("Input Kosong", "Nama pimpinan tidak boleh kosong.")

    def change_pimpinan_color(self, pimpinan_id, color_display_widget):
        current_color_str = color_display_widget.cget("bg_color")
        color_code, hex_color = colorchooser.askcolor(color=current_color_str, title="Pilih Warna Pimpinan")
        if hex_color: # hex_color will be None if user cancels
            success, message = update_pimpinan_color(pimpinan_id, hex_color)
            if success:
                color_display_widget.configure(bg_color=hex_color) # Update UI immediately
                messagebox.showinfo("Berhasil", message)
                # No need to reload pimpinan list, just update the widget
            else:
                messagebox.showerror("Error", message)

    def confirm_delete_pimpinan(self, pimpinan_id):
        pimpinan_info = get_pimpinan_by_id(pimpinan_id)
        if pimpinan_info:
            pimpinan_name = pimpinan_info['nama']
            if messagebox.askyesno("Konfirmasi Hapus", f"Apakah Anda yakin ingin menghapus pimpinan '{pimpinan_name}'? \n\nSemua kegiatan yang terkait dengan pimpinan ini TIDAK AKAN terhapus, tetapi akan menjadi tidak terhubung dengan pimpinan manapun."):
                success, message = delete_pimpinan(pimpinan_id)
                if success:
                    messagebox.showinfo("Berhasil", message)
                    self.load_pimpinan_list() # Refresh the list
                else:
                    messagebox.showerror("Error", message)
        else:
            messagebox.showerror("Error", "Pimpinan tidak ditemukan.")


    def on_closing(self):
        self.master_app.focus_set()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
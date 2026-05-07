import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw
from datetime import datetime
import threading
import time
import os
from pystray import Icon as TrayIcon, Menu as TrayMenu, MenuItem as TrayItem
from PIL import Image, ImageDraw

class UltimateAutoSaver:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Auto-Save Controller")
        self.root.geometry("400x520")
        self.root.resizable(False, False)
        
        self.file_paths = []
        self.is_running = False
        self.after_id = None
        
        # --- UI LAYOUT ---
        tk.Label(root, text="Excel Auto-Save Controller", font=("Arial", 14, "bold")).pack(pady=10)

        file_frame = tk.Frame(root)
        file_frame.pack(pady=5)
        self.add_btn = tk.Button(file_frame, text="➕ Add Files", command=self.add_files, width=15)
        self.add_btn.grid(row=0, column=0, padx=5)
        self.remove_btn = tk.Button(file_frame, text="➖ Remove Selected", command=self.remove_files, width=15)
        self.remove_btn.grid(row=0, column=1, padx=5)

        self.file_listbox = tk.Listbox(root, width=55, height=8, selectmode=tk.MULTIPLE)
        self.file_listbox.pack(pady=10, padx=20)

        timer_frame = tk.Frame(root)
        timer_frame.pack(pady=5)
        
        tk.Label(timer_frame, text="Min:").grid(row=0, column=0)
        self.min_entry = tk.Entry(timer_frame, width=5, justify='center')
        self.min_entry.insert(0, "1")
        self.min_entry.grid(row=0, column=1, padx=5)

        tk.Label(timer_frame, text="Sec:").grid(row=0, column=2)
        self.sec_entry = tk.Entry(timer_frame, width=5, justify='center')
        self.sec_entry.insert(0, "0")
        self.sec_entry.grid(row=0, column=3, padx=5)

        self.status_label = tk.Label(root, text="Status: IDLE", fg="gray", font=("Arial", 9, "italic"))
        self.status_label.pack(pady=10)

        self.start_btn = tk.Button(root, text="START AUTO-SAVE", command=self.start_saving, 
                                   bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), width=25)
        self.start_btn.pack(pady=5)

        self.stop_btn = tk.Button(root, text="STOP", command=self.stop_saving, 
                                  state="disabled", bg="#f44336", fg="white", 
                                  disabledforeground="#ffffff", font=("Arial", 10, "bold"), width=25)
        self.stop_btn.pack(pady=5)

        self.tray_btn = tk.Button(root, text="Minimize to Tray", command=self.hide_to_tray, 
                                  bg="#2196F3", fg="white", width=25)
        self.tray_btn.pack(pady=20)

    def create_image(self):
        image = Image.new('RGB', (64, 64), color=(33, 150, 243))
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill=(255, 255, 255))
        return image

    def hide_to_tray(self):
        self.root.withdraw()
        menu = TrayMenu(TrayItem('Show App', self.show_window), TrayItem('Exit', self.quit_app))
        self.icon = TrayIcon("ExcelAutoSave", self.create_image(), "Excel Auto-Saver", menu)
        threading.Thread(target=self.icon.run, daemon=True).start()

    def show_window(self):
        if hasattr(self, 'icon'): self.icon.stop()
        self.root.after(0, self.root.deiconify)

    def quit_app(self):
        if hasattr(self, 'icon'): self.icon.stop()
        self.root.quit()

    def add_files(self):
        files = filedialog.askopenfilenames(title="Select Files", filetypes=[("Excel Files", "*.xlsx *.xlsb *.xlsm *.csv")])
        for f in files:
            abs_f = os.path.abspath(f)
            if abs_f not in self.file_paths:
                self.file_paths.append(abs_f)
                self.file_listbox.insert(tk.END, os.path.basename(abs_f))

    def remove_files(self):
        selected = list(self.file_listbox.curselection())
        for index in reversed(selected):
            self.file_listbox.delete(index)
            self.file_paths.pop(index)

    def save_process(self):
        if not self.is_running: return
        
        files_skipped = False
        
        for path in self.file_paths:
            success = False
            # 5 attempts with 2s delay = 10s of patience per file
            for attempt in range(5):
                try:
                    wb = xw.Book(path)
                    wb.save()
                    success = True
                    break 
                except Exception as e:
                    err_str = str(e)
                    # ALL 3 ERROR CODES FROM SCREENSHOTS: 
                    # 0x800ac472 (Busy), -2147467259 (Locked/Menu), -2147352567 (App Exception)
                    if any(code in err_str for code in ['0x800ac472', '-2147467259', '-2147352567']):
                        time.sleep(2) 
                        continue
                    else:
                        break # Critical/Unknown error
            
            if not success: files_skipped = True

        if files_skipped:
            self.status_label.config(text="Warning: Files skipped (Excel was busy/locked)", fg="orange")
        else:
            self.status_label.config(text=f"Last saved: {datetime.now().strftime('%H:%M:%S')}", fg="green")

        try:
            total_seconds = (float(self.min_entry.get() or 0) * 60) + float(self.sec_entry.get() or 0)
            if total_seconds <= 0: total_seconds = 1
            self.after_id = self.root.after(int(total_seconds * 1000), self.save_process)
        except:
            self.status_label.config(text="Status: Stopped (Invalid Time Input)", fg="red")
            self.stop_saving()

    def start_saving(self):
        if not self.file_paths:
            messagebox.showwarning("No Files", "Please add at least one file.")
            return
        self.is_running = True
        self.set_ui_state("running")
        self.save_process()

    def stop_saving(self):
        self.is_running = False
        if self.after_id: self.root.after_cancel(self.after_id)
        self.set_ui_state("stopped")

    def set_ui_state(self, state):
        if state == "running":
            self.start_btn.config(state="disabled")
            self.stop_btn.config(state="normal", bg="#f44336")
            for entry in [self.add_btn, self.remove_btn, self.min_entry, self.sec_entry]: entry.config(state="disabled")
        else:
            self.status_label.config(text="Status: STOPPED", fg="red")
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled", bg="#e0e0e0")
            for entry in [self.add_btn, self.remove_btn, self.min_entry, self.sec_entry]: entry.config(state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    app = UltimateAutoSaver(root)
    root.mainloop()
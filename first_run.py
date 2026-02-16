import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from config import save_data_root, ensure_data_root_structure


class FirstRunWindow(tk.Toplevel):
    def __init__(self, master, on_complete):
        super().__init__(master)
        self.on_complete = on_complete
        self.selected_path = None

        self.title("Welcome to Worship Slides")
        self.geometry("500x250")
        self.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        label = tk.Label(
            self,
            text=(
                "This app needs a folder to store:\n"
                "• Songs\n• Templates\n• Generated slides\n\n"
                "Choose or create a folder to continue."
            ),
            justify="left",
        )
        label.pack(pady=10)

        self.path_var = tk.StringVar()

        path_entry = tk.Entry(self, textvariable=self.path_var, width=50)
        path_entry.pack(pady=5)

        choose_btn = tk.Button(self, text="Choose Folder", command=self.choose_folder)
        choose_btn.pack(pady=5)

        continue_btn = tk.Button(self, text="Continue", command=self.continue_clicked)
        continue_btn.pack(pady=10)

    def choose_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_path = Path(folder)
            self.path_var.set(str(self.selected_path))

    def continue_clicked(self):
        if not self.selected_path:
            messagebox.showerror("Error", "Please choose a folder.")
            return

        try:
            ensure_data_root_structure(str(self.selected_path))

            save_data_root(self.selected_path)
            self.destroy()
            self.on_complete()

        except Exception as e:
            messagebox.showerror("Error", f"Could not set up folder:\n{e}")
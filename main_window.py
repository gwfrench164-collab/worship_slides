import tkinter as tk
from tkinter import messagebox
from config import load_data_root
from song_builder import SongBuilder
from build_window import BuildWindow
from pathlib import Path
from tkinter import filedialog
from pdf_importer_ocr import import_song_from_pdf

class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Worship Slides")
        self.geometry("400x300")

        self._build_menu()
        self._build_main_buttons()

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Song", command=self.open_song_builder)
        file_menu.add_command(label="Open Song", command=self.open_existing_song)
        file_menu.add_command(label="Import Song from PDF", command=self.import_song_from_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        file_menu.add_command(label="Build Slides", command=self.open_build_window)

        menubar.add_cascade(label="File", menu=file_menu)

        library_menu = tk.Menu(menubar, tearoff=0)
        library_menu.add_command(label="View Songs", command=self.not_implemented)
        library_menu.add_command(label="Manage Templates", command=self.not_implemented)
        menubar.add_cascade(label="Library", menu=library_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def _build_main_buttons(self):
        frame = tk.Frame(self)
        frame.pack(expand=True)

        tk.Button(frame, text="Create New Song", width=25, command=self.open_song_builder).pack(pady=5)
        tk.Button(frame, text="Open Existing Song", width=25, command=self.open_existing_song).pack(pady=5)
        tk.Button(frame, text="Import Song from PDF", width=25, command=self.import_song_from_pdf).pack(pady=5)
        tk.Button(frame, text="Build Slides", width=25, command=self.open_build_window).pack(pady=5)
        tk.Button(frame, text="Manage Library", width=25, command=self.not_implemented).pack(pady=5)

    def not_implemented(self):
        messagebox.showinfo("Not implemented", "This feature is not implemented yet.")

    def show_about(self):
        messagebox.showinfo(
            "About",
            "Worship Slides\nVersion 0.1\n\nCreate worship song slides quickly."
        )

    def open_song_builder(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        SongBuilder(self, songs_folder)

    def open_existing_song(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        song_path = filedialog.askopenfilename(
            title="Open Song JSON",
            initialdir=songs_folder,
            filetypes=[("Song JSON", "*.json")]
        )

        if not song_path:
            return

        SongBuilder(self, songs_folder, open_song=Path(song_path))

    def open_build_window(self):
        BuildWindow(self)

    def import_song_from_pdf(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        pdf_path = filedialog.askopenfilename(
            title="Select Song PDF",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not pdf_path:
            return

        songs_folder = Path(data_root) / "songs"

        try:
            song_json = import_song_from_pdf(
                Path(pdf_path),
                songs_folder
            )
        except Exception as e:
            messagebox.showerror("Import failed", str(e))
            return

        messagebox.showinfo(
            "Import complete",
            f"Song imported:\n{song_json.name}"
        )

        # Open the Song Builder with the new song
        SongBuilder(self, songs_folder, open_song=song_json)

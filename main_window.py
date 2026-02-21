import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from config import load_data_root
from song_builder import SongBuilder
from build_window import BuildWindow
from pathlib import Path
from pdf_importer_ocr import import_song_from_pdf
from library_window import LibraryWindow
from notes_reader import read_notes_text
from bible_extractor import extract_ordered_refs, load_bible_json, fetch_verse_text
from verse_slide_builder import build_verse_deck
from config import load_bible_json_path, save_bible_json_path, auto_find_kjv_json

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
        file_menu.add_separator()
        file_menu.add_command(label="Extract Verse List from Notes", command=self.extract_verse_list_from_notes)
        file_menu.add_command(label="Build Verse Slides from Notes", command=self.build_verse_slides_from_notes)

        menubar.add_cascade(label="File", menu=file_menu)

        library_menu = tk.Menu(menubar, tearoff=0)
        library_menu.add_command(label="Manage Songs", command=self.open_library_window)
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
        tk.Button(frame, text="Manage Library", width=25, command=self.open_library_window).pack(pady=5)

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

    def open_library_window(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        LibraryWindow(self, songs_folder)

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

    def _get_bible(self):
        # 1) from config
        p = load_bible_json_path()
        if p and Path(p).exists():
            return load_bible_json(p)

        # 2) auto-find (since you put kjv.json in repo)
        auto = auto_find_kjv_json()
        if auto and Path(auto).exists():
            save_bible_json_path(auto)
            return load_bible_json(auto)

        # 3) ask user once
        pick = filedialog.askopenfilename(
            title="Select kjv.json",
            filetypes=[("JSON files", "*.json")]
        )
        if not pick:
            raise RuntimeError("kjv.json not selected.")
        save_bible_json_path(pick)
        return load_bible_json(pick)

    def extract_verse_list_from_notes(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        notes_file = filedialog.askopenfilename(
            title="Select Notes File",
            filetypes=[("Pages", "*.pages"), ("Word", "*.docx"), ("Text", "*.txt"), ("All files", "*.*")]
        )
        if not notes_file:
            return

        try:
            text = read_notes_text(Path(notes_file))
            refs = extract_ordered_refs(text)
        except Exception as e:
            messagebox.showerror("Failed", str(e))
            return

        if not refs:
            messagebox.showinfo("No verses found", "No verse references were detected.")
            return

        refs_folder = Path(data_root) / "notes_refs"
        refs_folder.mkdir(parents=True, exist_ok=True)

        out_path = refs_folder / f"{Path(notes_file).stem}.refs.json"
        payload = {
            "schema_version": "1.0",
            "source": {"notes_file": Path(notes_file).name},
            "verses": refs
        }

        import json
        out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

        messagebox.showinfo("Saved", f"Found {len(refs)} references.\nSaved:\n{out_path.name}")

    def build_verse_slides_from_notes(self):
        data_root = load_data_root()
        if not data_root:
            messagebox.showerror("Error", "Data folder not set.")
            return

        notes_file = filedialog.askopenfilename(
            title="Select Notes File",
            filetypes=[("Pages", "*.pages"), ("Word", "*.docx"), ("Text", "*.txt"), ("All files", "*.*")]
        )
        if not notes_file:
            return

        try:
            text = read_notes_text(Path(notes_file))
            refs = extract_ordered_refs(text)
            if not refs:
                messagebox.showinfo("No verses found", "No verse references were detected.")
                return

            bible = self._get_bible()

            # choose template
            templates_folder = Path(data_root) / "templates"
            template_file = filedialog.askopenfilename(
                title="Select PPTX Template",
                initialdir=templates_folder,
                filetypes=[("PowerPoint", "*.pptx")]
            )
            if not template_file:
                return

            output_file = filedialog.asksaveasfilename(
                title="Save Verse Slides As",
                defaultextension=".pptx",
                initialfile=f"{Path(notes_file).stem}_verses.pptx",
                filetypes=[("PowerPoint", "*.pptx")]
            )
            if not output_file:
                return

            # simple control: ask max lines per slide
            max_lines = simpledialog.askinteger(
                "Max lines",
                "Max verse lines per slide? (Reference is added above)",
                initialvalue=4, minvalue=1, maxvalue=10
            )
            if not max_lines:
                return

            refs_and_texts = [(r, fetch_verse_text(r, bible)) for r in refs]

            build_verse_deck(
                Path(template_file),
                refs_and_texts,
                Path(output_file),
                fit_preset="loose",  # or "normal" / "loose"
            )

            messagebox.showinfo("Done", f"Created:\n{Path(output_file).name}")

        except Exception as e:
            messagebox.showerror("Build failed", str(e))

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

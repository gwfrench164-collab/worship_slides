import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from bible_extractor import extract_ordered_refs, fetch_verse_text, load_bible_json
from build_window import BuildWindow
from config import auto_find_kjv_json, load_bible_json_path, load_data_root, save_bible_json_path, load_ccli_number, save_ccli_number
from library_window import LibraryWindow
from notes_reader import read_notes_text
from pdf_importer_ocr import import_song_from_pdf
from pptx_utils import merge_presentations
from song_builder import SongBuilder

from verse_slide_builder import build_verse_deck


class VerseSlidesWindow(tk.Toplevel):
    def __init__(self, parent, data_root):
        super().__init__(parent)
        self.parent = parent
        self.data_root = Path(data_root)

        self.title("Create Verse Slides from Sermon Notes")
        self.geometry("640x420")
        self.minsize(640, 420)
        self.resizable(True, True)

        self.notes_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Step 1: Select sermon notes.")

        self._build_ui()

    def _build_ui(self):
        outer = ttk.Frame(self, padding=16)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Create Verse Slides from Sermon Notes", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text=(
                "This tool scans sermon notes for Bible references, then creates scripture slides.\n"
                "Steps: 1) Select notes  2) Select a verse template  3) Choose where to save  4) Build slides"
            ),
            style="Subheader.TLabel",
            justify="left",
        ).pack(anchor="w", pady=(4, 14))

        form = ttk.LabelFrame(outer, text="Verse Slide Inputs", style="Card.TLabelframe")
        form.pack(fill="x", pady=(0, 12))
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Sermon notes").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(form, textvariable=self.notes_var).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(form, text="Browse...", command=self._browse_notes).grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=6)

        ttk.Label(form, text="Verse template").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(form, textvariable=self.template_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(form, text="Browse...", command=self._browse_template).grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=6)

        ttk.Label(form, text="Save verse slides as").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=6)
        ttk.Entry(form, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(form, text="Browse...", command=self._browse_output).grid(row=2, column=2, sticky="ew", padx=(8, 0), pady=6)

        tips = ttk.LabelFrame(outer, text="What this does", style="Card.TLabelframe")
        tips.pack(fill="both", expand=True)
        ttk.Label(
            tips,
            text=(
                "• Finds Bible references like John 3:16 or Romans 8:28 in your notes\n"
                "• Looks up the verse text from your configured Bible JSON\n"
                "• Builds a PowerPoint deck using your verse slide template"
            ),
            justify="left",
        ).pack(anchor="w")

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(12, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Build Verse Slides", command=self._build).pack(side="right", padx=(0, 8))

        ttk.Label(self, textvariable=self.status_var, style="Status.TLabel", anchor="w", relief="sunken").pack(fill="x", side="bottom")

    def _browse_notes(self):
        path = filedialog.askopenfilename(
            title="Select Sermon Notes",
            filetypes=[("Pages", "*.pages"), ("Word", "*.docx"), ("Text", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return
        self.notes_var.set(path)
        if not self.output_var.get().strip():
            suggested = Path(path).with_name(f"{Path(path).stem}_verses.pptx")
            self.output_var.set(str(suggested))
        self.status_var.set("Step 2: Select a verse slide template.")

    def _browse_template(self):
        templates_folder = self.data_root / "templates"
        path = filedialog.askopenfilename(
            title="Select Verse Slide Template",
            initialdir=templates_folder,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not path:
            return
        self.template_var.set(path)
        self.status_var.set("Step 3: Choose where to save the finished PowerPoint.")

    def _browse_output(self):
        initial_name = self.output_var.get().strip() or "sermon_notes_verses.pptx"
        path = filedialog.asksaveasfilename(
            title="Save Verse Slides As",
            defaultextension=".pptx",
            initialfile=Path(initial_name).name,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not path:
            return
        self.output_var.set(path)
        self.status_var.set("Ready to build verse slides.")

    def _build(self):
        notes_file = self.notes_var.get().strip()
        template_file = self.template_var.get().strip()
        output_file = self.output_var.get().strip()

        if not notes_file:
            messagebox.showerror("Missing notes", "Please select sermon notes first.")
            self.status_var.set("Step 1: Select sermon notes.")
            return
        if not template_file:
            messagebox.showerror("Missing template", "Please select a verse slide template.")
            self.status_var.set("Step 2: Select a verse slide template.")
            return
        if not output_file:
            messagebox.showerror("Missing output file", "Please choose where to save the verse slides.")
            self.status_var.set("Step 3: Choose where to save the finished PowerPoint.")
            return

        self.status_var.set("Building verse slides...")
        self.parent.set_status("Building verse slides from sermon notes...")
        self.update_idletasks()

        try:
            text = read_notes_text(Path(notes_file))
            refs = extract_ordered_refs(text)
            if not refs:
                self.status_var.set("No Bible references were found in the selected notes.")
                messagebox.showinfo("No verses found", "No verse references were detected in the selected notes.")
                return

            bible = self.parent._get_bible()
            refs_and_texts = [(r, fetch_verse_text(r, bible)) for r in refs]
            build_verse_deck(Path(template_file), refs_and_texts, Path(output_file), fit_preset="normal")

            self.status_var.set(f"Done: created {Path(output_file).name}")
            self.parent.set_status(f"Verse slides created: {Path(output_file).name}")
            messagebox.showinfo("Done", f"Created:\n{Path(output_file).name}")
        except Exception as e:
            self.status_var.set("Verse slide build failed.")
            self.parent.set_status("Verse slide build failed.")
            messagebox.showerror("Build failed", str(e))


class MainWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Worship Slides")
        self.geometry("560x430")
        self.minsize(560, 430)

        self.status_var = tk.StringVar(value="Ready")

        self._configure_style()
        self._build_menu()
        self._build_main_buttons()

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Header.TLabel", font=("TkDefaultFont", 16, "bold"))
        style.configure("Subheader.TLabel", font=("TkDefaultFont", 10))
        style.configure("TButton", padding=(12, 8))
        style.configure("Card.TLabelframe", padding=10)
        style.configure("Status.TLabel", padding=(10, 6))

    def set_status(self, text: str):
        self.status_var.set(text)

    def _build_menu(self):
        menubar = tk.Menu(self)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New Song", command=self.open_song_builder)
        file_menu.add_command(label="Open Song", command=self.open_existing_song)
        file_menu.add_command(label="Import Song from PDF", command=self.import_song_from_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Build Slides", command=self.open_build_window)
        file_menu.add_separator()
        file_menu.add_command(label="Create Verse Slides from Sermon Notes", command=self.build_verse_slides_from_notes)
        file_menu.add_command(label="Extract Verse List from Notes", command=self.extract_verse_list_from_notes)
        file_menu.add_command(label="Merge Song + Verse Decks", command=self.merge_song_and_verse_decks)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)

        library_menu = tk.Menu(menubar, tearoff=0)
        library_menu.add_command(label="Manage Songs", command=self.open_library_window)
        library_menu.add_command(label="Manage Templates", command=self.not_implemented)
        menubar.add_cascade(label="Library", menu=library_menu)

        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="Set Church CCLI Number...", command=self.set_ccli_number)
        menubar.add_cascade(label="Settings", menu=settings_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.config(menu=menubar)

    def _build_main_buttons(self):
        outer = ttk.Frame(self, padding=18)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="Worship Slides", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Create song slides, verse slides, and merged service decks.",
            style="Subheader.TLabel",
        ).pack(anchor="w", pady=(4, 16))

        actions = ttk.LabelFrame(outer, text="Main Actions", style="Card.TLabelframe")
        actions.pack(fill="both", expand=True)

        button_specs = [
            ("Create New Song", self.open_song_builder),
            ("Open Existing Song", self.open_existing_song),
            ("Import Song from PDF", self.import_song_from_pdf),
            ("Manage Library", self.open_library_window),
            ("Build Slides", self.open_build_window),
            # ("Extract Verse List from Notes", self.extract_verse_list_from_notes),  # Removed as per instructions
            ("Create Verse Slides from Sermon Notes", self.build_verse_slides_from_notes),
            ("Merge Song + Verse Decks", self.merge_song_and_verse_decks),
        ]

        for col in range(2):
            actions.columnconfigure(col, weight=1)

        for i, (label, command) in enumerate(button_specs):
            row = i // 2
            col = i % 2
            ttk.Button(actions, text=label, command=command).grid(
                row=row,
                column=col,
                sticky="ew",
                padx=6,
                pady=6,
            )

        ttk.Label(self, textvariable=self.status_var, style="Status.TLabel", anchor="w", relief="sunken").pack(
            fill="x", side="bottom"
        )

    def not_implemented(self):
        self.set_status("This feature is not implemented yet.")
        messagebox.showinfo("Not implemented", "This feature is not implemented yet.")

    def show_about(self):
        self.set_status("Viewed About window.")
        messagebox.showinfo(
            "About",
            "Worship Slides\nVersion 0.1\n\nCreate worship song slides quickly.",
        )

    def set_ccli_number(self):
        """Prompt user to set the church CCLI number."""
        current = ""
        try:
            current = load_ccli_number() or ""
        except Exception:
            pass

        dialog = tk.Toplevel(self)
        dialog.title("Set Church CCLI Number")
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Enter your church CCLI number:").pack(anchor="w")

        entry_var = tk.StringVar(value=current)
        entry = ttk.Entry(frame, textvariable=entry_var, width=30)
        entry.pack(fill="x", pady=(6, 12))
        entry.focus_set()

        def save():
            value = entry_var.get().strip()
            try:
                save_ccli_number(value)
                self.set_status("CCLI number saved.")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")

        ttk.Button(buttons, text="Cancel", command=dialog.destroy).pack(side="right", padx=4)
        ttk.Button(buttons, text="Save", command=save).pack(side="right")

    def open_song_builder(self):
        self.set_status("Opening Song Builder...")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        SongBuilder(self, songs_folder)
        self.set_status("Song Builder opened.")

    def open_library_window(self):
        self.set_status("Opening Song Library...")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        LibraryWindow(self, songs_folder)
        self.set_status("Song Library opened.")

    def open_existing_song(self):
        self.set_status("Select a song to open.")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        songs_folder = Path(data_root) / "songs"
        song_path = filedialog.askopenfilename(
            title="Open Song JSON",
            initialdir=songs_folder,
            filetypes=[("Song JSON", "*.json")],
        )

        if not song_path:
            self.set_status("Open song cancelled.")
            return

        SongBuilder(self, songs_folder, open_song=Path(song_path))
        self.set_status(f"Opened song: {Path(song_path).name}")

    def open_build_window(self):
        self.set_status("Opening Build Slides window...")
        BuildWindow(self)
        self.set_status("Build Slides window opened.")

    def merge_song_and_verse_decks(self):
        self.set_status("Preparing deck merge...")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        output_folder = Path(data_root) / "output"
        output_folder.mkdir(parents=True, exist_ok=True)

        song_deck = filedialog.askopenfilename(
            title="Select Song Slide Deck",
            initialdir=output_folder,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not song_deck:
            self.set_status("Deck merge cancelled.")
            return

        verse_deck = filedialog.askopenfilename(
            title="Select Verse Slide Deck",
            initialdir=output_folder,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not verse_deck:
            self.set_status("Deck merge cancelled.")
            return

        default_name = f"{Path(song_deck).stem}_with_verses.pptx"
        output_path = filedialog.asksaveasfilename(
            title="Save Merged Deck As",
            initialdir=output_folder,
            defaultextension=".pptx",
            initialfile=default_name,
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if not output_path:
            self.set_status("Deck merge cancelled.")
            return

        try:
            merge_presentations(song_deck, verse_deck, output_path)
        except Exception as e:
            self.set_status("Deck merge failed.")
            messagebox.showerror("Merge failed", str(e))
            return

        self.set_status(f"Merged deck saved: {Path(output_path).name}")
        messagebox.showinfo("Merge complete", f"Created:\n{Path(output_path).name}")

    def _get_bible(self):
        p = load_bible_json_path()
        if p and Path(p).exists():
            return load_bible_json(p)

        auto = auto_find_kjv_json()
        if auto and Path(auto).exists():
            save_bible_json_path(auto)
            return load_bible_json(auto)

        pick = filedialog.askopenfilename(
            title="Select kjv.json",
            filetypes=[("JSON files", "*.json")],
        )
        if not pick:
            raise RuntimeError("kjv.json not selected.")
        save_bible_json_path(pick)
        return load_bible_json(pick)

    def extract_verse_list_from_notes(self):
        self.set_status("Select notes to extract verses.")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        notes_file = filedialog.askopenfilename(
            title="Select Notes File",
            filetypes=[("Pages", "*.pages"), ("Word", "*.docx"), ("Text", "*.txt"), ("All files", "*.*")],
        )
        if not notes_file:
            self.set_status("Verse extraction cancelled.")
            return

        try:
            text = read_notes_text(Path(notes_file))
            refs = extract_ordered_refs(text)
        except Exception as e:
            self.set_status("Verse extraction failed.")
            messagebox.showerror("Failed", str(e))
            return

        if not refs:
            self.set_status("No verses found in notes.")
            messagebox.showinfo("No verses found", "No verse references were detected.")
            return

        refs_folder = Path(data_root) / "notes_refs"
        refs_folder.mkdir(parents=True, exist_ok=True)

        out_path = refs_folder / f"{Path(notes_file).stem}.refs.json"
        payload = {
            "schema_version": "1.0",
            "source": {"notes_file": Path(notes_file).name},
            "verses": refs,
        }
        out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

        self.set_status(f"Saved verse list: {out_path.name}")
        messagebox.showinfo("Saved", f"Found {len(refs)} references.\nSaved:\n{out_path.name}")

    def build_verse_slides_from_notes(self):
        self.set_status("Opening Create Verse Slides from Sermon Notes...")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        VerseSlidesWindow(self, data_root)
        self.set_status("Verse slides window opened.")

    def import_song_from_pdf(self):
        self.set_status("Select a PDF to import.")
        data_root = load_data_root()
        if not data_root:
            self.set_status("Data folder not set.")
            messagebox.showerror("Error", "Data folder not set.")
            return

        pdf_path = filedialog.askopenfilename(
            title="Select Song PDF",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not pdf_path:
            self.set_status("PDF import cancelled.")
            return

        songs_folder = Path(data_root) / "songs"
        try:
            song_data = import_song_from_pdf(Path(pdf_path))
        except Exception as e:
            self.set_status("PDF import failed.")
            messagebox.showerror("Import failed", str(e))
            return

        messagebox.showinfo(
            "Import complete",
            "OCR import complete. Review the draft and click Save Song when ready.",
        )

        SongBuilder(self, songs_folder, draft_song=song_data)
        self.set_status("PDF imported into draft song.")

import json
import tkinter as tk
from tkinter import messagebox, simpledialog
from pathlib import Path
import re


SECTION_TYPES = ["Title", "Verse", "Chorus", "Bridge", "Outro", "Other"]


class SongBuilder(tk.Toplevel):
    def __init__(self, parent, songs_folder, open_song=None):
        super().__init__(parent)
        self.songs_folder = songs_folder
        self.open_song = open_song

        self.title("Song Builder")
        self.geometry("700x450")

        self.sections = []
        self.current_section_index = None

        self._build_ui()
        if self.open_song:
            self.load_song(self.open_song)

    def load_song(self, song_path):
        with open(song_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        # ---- Song metadata ----
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, data["song"].get("title", ""))

        self.author_entry.delete(0, tk.END)
        self.author_entry.insert(0, data["song"].get("author", ""))

        # ---- Sections ----
        self.sections = data["structure"]["sections"]

        self.section_listbox.delete(0, tk.END)
        for section in self.sections:
            self.section_listbox.insert(tk.END, section.get("label", "Section"))

        # ---- Auto-select first section ----
        if self.sections:
            self.section_listbox.selection_set(0)
            self.current_section_index = 0
            self._load_section_into_editor(0)

    def _load_section_into_editor(self, index):
        section = self.sections[index]

        self.section_label.config(text=section.get("label", ""))

        self.lyrics_text.delete("1.0", tk.END)

        # New format: section["lines"] is the source of truth.
        # Backward compatible: if the section still has "slides", we flatten them.
        lines = self._get_section_lines(section)
        self.lyrics_text.insert(tk.END, "\n".join(lines))

    def _get_section_lines(self, section: dict) -> list[str]:
        """Return raw lyric lines for a section (new or legacy format)."""
        if isinstance(section.get("lines"), list):
            return [str(x) for x in section.get("lines", [])]

        out: list[str] = []
        for slide in section.get("slides", []):
            for line in slide.get("lines", []):
                out.append(str(line))
        return out

    # ---------------- UI ----------------

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=5)

        tk.Label(top, text="Title *").grid(row=0, column=0, sticky="w")
        self.title_entry = tk.Entry(top, width=40)
        self.title_entry.grid(row=0, column=1, sticky="w")

        tk.Label(top, text="Author").grid(row=1, column=0, sticky="w")
        self.author_entry = tk.Entry(top, width=40)
        self.author_entry.grid(row=1, column=1, sticky="w")

        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=5)

        # Sections list
        left = tk.Frame(main)
        left.pack(side="left", fill="y")

        tk.Label(left, text="Sections").pack()

        self.section_listbox = tk.Listbox(left, width=20)
        self.section_listbox.pack(fill="y", expand=True)
        self.section_listbox.bind("<<ListboxSelect>>", self.on_section_select)

        tk.Button(left, text="+ Add", command=self.add_section).pack(fill="x", pady=2)
        tk.Button(left, text="− Remove", command=self.remove_section).pack(fill="x", pady=2)

        # Lyrics editor
        right = tk.Frame(main)
        right.pack(side="left", fill="both", expand=True, padx=10)

        self.section_label = tk.Label(right, text="No section selected")
        self.section_label.pack(anchor="w")

        tk.Label(right, text="Lyrics (one line per line):").pack(anchor="w")

        self.lyrics_text = tk.Text(right, height=10)
        self.lyrics_text.pack(fill="both", expand=True)

        # Bottom buttons
        bottom = tk.Frame(self)
        bottom.pack(pady=10)

        tk.Button(bottom, text="Cancel", command=self.destroy).pack(side="left", padx=5)
        tk.Button(bottom, text="Save Song", command=self.save_song).pack(side="left", padx=5)

    # ---------------- Sections ----------------

    def add_section(self):
        # Save current section lyrics BEFORE switching
        self._save_current_lyrics()

        section_type = simpledialog.askstring(
            "Add Section",
            "Section type (Title, Verse, Chorus, Bridge, Outro, Other):"
        )
        if not section_type:
            return

        section_type = section_type.strip().title()
        if section_type not in SECTION_TYPES:
            messagebox.showerror("Error", "Invalid section type.")
            return

        label = self._generate_label(section_type)

        section = {
            "id": self._make_id(label),
            "label": label,
            "type": section_type.lower(),
            # New format: store raw lines only. SlideBuilder decides chunking.
            "lines": []
        }

        self.sections.append(section)
        self.section_listbox.insert(tk.END, label)

        # Auto-select the new section
        index = len(self.sections) - 1
        self.section_listbox.select_clear(0, tk.END)
        self.section_listbox.select_set(index)
        self.section_listbox.event_generate("<<ListboxSelect>>")

    def remove_section(self):
        index = self.section_listbox.curselection()
        if not index:
            return

        i = index[0]
        del self.sections[i]
        self.section_listbox.delete(i)
        self.lyrics_text.delete("1.0", tk.END)
        self.section_label.config(text="No section selected")
        self.current_section_index = None

    def on_section_select(self, event):
        # Save current section before switching
        self._save_current_lyrics()

        selection = self.section_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        self.current_section_index = index
        self._load_section_into_editor(index)

    # ---------------- Helpers ----------------

    def _save_current_lyrics(self):
        if self.current_section_index is None:
            return

        raw_lines = self.lyrics_text.get("1.0", tk.END).splitlines()
        # Preserve line breaks; remove totally blank lines.
        lines = [ln.rstrip() for ln in raw_lines if ln.strip()]

        section = self.sections[self.current_section_index]
        section["lines"] = lines
        # If this section came from the old schema, remove legacy "slides".
        section.pop("slides", None)

    def _generate_label(self, section_type):
        if section_type in ("Verse",):
            count = sum(1 for s in self.sections if s["type"] == "verse") + 1
            return f"Verse {count}"
        return section_type

    def _make_id(self, label):
        return re.sub(r"[^a-z0-9]", "", label.lower())

    # ---------------- Save ----------------

    def save_song(self):
        self._save_current_lyrics()

        title = self.title_entry.get().strip()
        if not title:
            messagebox.showerror("Error", "Song title is required.")
            return

        if not self.sections:
            messagebox.showerror("Error", "Add at least one section.")
            return

        song_data = {
            "schema_version": "1.0",
            "song": {
                "title": title,
                "author": self.author_entry.get().strip(),
                "copyright": "",
                "ccli_number": "",
                "notes": ""
            },
            "structure": {
                "sections": self.sections
            },
            "chords": {
                "enabled": False,
                "sections": {}
            }
        }

        filename = title.lower().replace(" ", "_") + ".json"
        path = self.songs_folder / filename

        if path.exists():
            if not messagebox.askyesno(
                "Overwrite?",
                f"{filename} already exists. Replace it?"
            ):
                return

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(song_data, f, indent=2, ensure_ascii=False)

            messagebox.showinfo("Saved", f"Song saved:\n{filename}")
            self.destroy()

        except Exception as e:
            messagebox.showerror("Error", str(e))
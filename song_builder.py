import json
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path
import re


SECTION_TYPES = ["Title", "Verse", "Chorus", "Bridge", "Outro", "Other"]


class SectionTypeDialog(tk.Toplevel):
    def __init__(self, parent, section_types=None):
        super().__init__(parent)
        self.title("Add Section")
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self.result = None
        self.section_types = section_types or SECTION_TYPES
        self.section_var = tk.StringVar(value=self.section_types[0])

        container = ttk.Frame(self, padding=14)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Select section type", style="SB.SectionHeader.TLabel").pack(anchor="w")
        ttk.Label(
            container,
            text="Choose the type of section you want to add.",
            style="SB.Subtle.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        self.combo = ttk.Combobox(
            container,
            textvariable=self.section_var,
            values=self.section_types,
            state="readonly",
            width=24,
        )
        self.combo.pack(fill="x")
        self.combo.focus_set()

        buttons = ttk.Frame(container)
        buttons.pack(fill="x", pady=(14, 0))
        ttk.Button(buttons, text="Cancel", command=self._cancel).pack(side="right")
        ttk.Button(buttons, text="Add Section", command=self._ok, style="Accent.TButton").pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda _e: self._ok())
        self.bind("<Escape>", lambda _e: self._cancel())

        self.update_idletasks()
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()
        x = parent_x + max(20, (parent_w - self.winfo_width()) // 2)
        y = parent_y + max(20, (parent_h - self.winfo_height()) // 2)
        self.geometry(f"+{x}+{y}")

    def _ok(self):
        self.result = self.section_var.get()
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class SongBuilder(tk.Toplevel):
    def __init__(self, parent, songs_folder, open_song=None, draft_song=None):
        super().__init__(parent)
        self.songs_folder = songs_folder
        self.open_song = open_song
        self.draft_song = draft_song

        self.title("Song Builder")
        self.geometry("860x560")
        self.minsize(760, 500)

        self.sections = []
        self.current_section_index = None

        self._configure_styles()
        self._build_ui()

        if self.open_song:
            self.load_song(self.open_song)
            self.set_status("Loaded saved song.")
        elif self.draft_song:
            self.load_song_data(self.draft_song)
            self.set_status("Loaded imported draft. Review and save when ready.")
        else:
            self.set_status("Ready to create a new song.")

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        base_font = ("TkDefaultFont", 10)
        heading_font = ("TkDefaultFont", 12, "bold")
        subheading_font = ("TkDefaultFont", 10, "bold")

        style.configure("SB.Title.TLabel", font=heading_font)
        style.configure("SB.SectionHeader.TLabel", font=subheading_font)
        style.configure("SB.Subtle.TLabel", foreground="#555555")
        style.configure("SB.Status.TLabel", padding=(8, 4))
        style.configure("SB.Toolbar.TFrame", padding=(10, 8))
        style.configure("SB.Content.TFrame", padding=(10, 10))
        style.configure("SB.Side.TLabelframe", padding=(8, 8))
        style.configure("SB.Editor.TLabelframe", padding=(10, 10))
        style.configure("Accent.TButton", padding=(12, 6))
        style.configure("TButton", padding=(10, 6))
        style.configure("TLabel", font=base_font)
        style.configure("TEntry", padding=4)

    def load_song(self, song_path):
        with open(song_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.load_song_data(data)

    def load_song_data(self, data):
        self.title_entry.delete(0, tk.END)
        self.title_entry.insert(0, data.get("song", {}).get("title", ""))

        self.author_entry.delete(0, tk.END)
        self.author_entry.insert(0, data.get("song", {}).get("author", ""))

        self.sections = list(data.get("structure", {}).get("sections", []))

        self.section_listbox.delete(0, tk.END)
        for section in self.sections:
            self.section_listbox.insert(tk.END, section.get("label", "Section"))

        if self.sections:
            self.section_listbox.selection_set(0)
            self.current_section_index = 0
            self._load_section_into_editor(0)
        else:
            self.current_section_index = None
            self.section_label_var.set("No section selected")
            self.lyrics_text.delete("1.0", tk.END)

    def _get_section_lines(self, section: dict) -> list[str]:
        if isinstance(section.get("lines"), list):
            return [str(x) for x in section.get("lines", [])]

        out: list[str] = []
        for slide in section.get("slides", []):
            for line in slide.get("lines", []):
                out.append(str(line))
        return out

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, style="SB.Toolbar.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Song Builder", style="SB.Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Create or edit song sections, then save the song JSON to your library.",
            style="SB.Subtle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        content = ttk.Frame(self, style="SB.Content.TFrame")
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(1, weight=1)
        content.rowconfigure(1, weight=1)

        meta = ttk.Frame(content)
        meta.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        meta.columnconfigure(1, weight=1)
        meta.columnconfigure(3, weight=1)

        ttk.Label(meta, text="Title *").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.title_entry = ttk.Entry(meta)
        self.title_entry.grid(row=0, column=1, sticky="ew", padx=(0, 16))

        ttk.Label(meta, text="Author").grid(row=0, column=2, sticky="w", padx=(0, 8))
        self.author_entry = ttk.Entry(meta)
        self.author_entry.grid(row=0, column=3, sticky="ew")

        side = ttk.LabelFrame(content, text="Sections", style="SB.Side.TLabelframe")
        side.grid(row=1, column=0, sticky="nsw", padx=(0, 12))
        side.columnconfigure(0, weight=1)
        side.rowconfigure(1, weight=1)

        ttk.Label(side, text="Select a section to edit.", style="SB.Subtle.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))

        self.section_listbox = tk.Listbox(side, width=22, height=16, exportselection=False)
        self.section_listbox.grid(row=1, column=0, sticky="nsew")
        self.section_listbox.bind("<<ListboxSelect>>", self.on_section_select)

        side_buttons = ttk.Frame(side)
        side_buttons.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        side_buttons.columnconfigure(0, weight=1)
        side_buttons.columnconfigure(1, weight=1)
        ttk.Button(side_buttons, text="Add Section", command=self.add_section, style="Accent.TButton").grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(side_buttons, text="Remove", command=self.remove_section).grid(row=0, column=1, sticky="ew", padx=(4, 0))

        editor = ttk.LabelFrame(content, text="Lyrics Editor", style="SB.Editor.TLabelframe")
        editor.grid(row=1, column=1, sticky="nsew")
        editor.columnconfigure(0, weight=1)
        editor.rowconfigure(2, weight=1)

        self.section_label_var = tk.StringVar(value="No section selected")
        ttk.Label(editor, textvariable=self.section_label_var, style="SB.SectionHeader.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            editor,
            text="Enter one lyric line per line. Blank lines are removed when the section is saved.",
            style="SB.Subtle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 8))

        text_frame = ttk.Frame(editor)
        text_frame.grid(row=2, column=0, sticky="nsew")
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)

        self.lyrics_text = tk.Text(text_frame, height=16, wrap="word", undo=True)
        self.lyrics_text.grid(row=0, column=0, sticky="nsew")
        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.lyrics_text.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.lyrics_text.configure(yscrollcommand=yscroll.set)

        bottom = ttk.Frame(self, padding=(10, 0, 10, 8))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        actions = ttk.Frame(bottom)
        actions.grid(row=0, column=1, sticky="e")
        ttk.Button(actions, text="Cancel", command=self.destroy).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Save Song", command=self.save_song, style="Accent.TButton").pack(side="left")

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken", style="SB.Status.TLabel")
        status.grid(row=3, column=0, sticky="ew")

    def set_status(self, message: str):
        self.status_var.set(message)
        self.update_idletasks()

    def _load_section_into_editor(self, index):
        section = self.sections[index]
        self.section_label_var.set(section.get("label", ""))
        self.lyrics_text.delete("1.0", tk.END)
        lines = self._get_section_lines(section)
        self.lyrics_text.insert(tk.END, "\n".join(lines))
        self.set_status(f"Editing {section.get('label', 'section')}.")

    def prompt_section_type(self):
        dialog = SectionTypeDialog(self)
        self.wait_window(dialog)
        return dialog.result

    def add_section(self):
        self._save_current_lyrics()

        section_type = self.prompt_section_type()
        if not section_type:
            self.set_status("Add section cancelled.")
            return

        label = self._generate_label(section_type)
        section = {
            "id": self._make_id(label),
            "label": label,
            "type": section_type.lower(),
            "lines": [],
        }

        self.sections.append(section)
        self.section_listbox.insert(tk.END, label)

        index = len(self.sections) - 1
        self.section_listbox.select_clear(0, tk.END)
        self.section_listbox.select_set(index)
        self.section_listbox.event_generate("<<ListboxSelect>>")
        self.set_status(f"Added {label}.")

    def remove_section(self):
        index = self.section_listbox.curselection()
        if not index:
            self.set_status("Select a section to remove.")
            return

        i = index[0]
        removed_label = self.sections[i].get("label", "Section")
        del self.sections[i]
        self.section_listbox.delete(i)
        self.lyrics_text.delete("1.0", tk.END)
        self.section_label_var.set("No section selected")
        self.current_section_index = None
        self.set_status(f"Removed {removed_label}.")

    def on_section_select(self, _event):
        self._save_current_lyrics()
        selection = self.section_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        self.current_section_index = index
        self._load_section_into_editor(index)

    def _save_current_lyrics(self):
        if self.current_section_index is None:
            return

        raw_lines = self.lyrics_text.get("1.0", tk.END).splitlines()
        lines = [ln.rstrip() for ln in raw_lines if ln.strip()]

        section = self.sections[self.current_section_index]
        section["lines"] = lines
        section.pop("slides", None)

    def _generate_label(self, section_type):
        if section_type in ("Verse",):
            count = sum(1 for s in self.sections if s["type"] == "verse") + 1
            return f"Verse {count}"
        return section_type

    def _make_id(self, label):
        return re.sub(r"[^a-z0-9]", "", label.lower())

    def save_song(self):
        self._save_current_lyrics()

        title = self.title_entry.get().strip()
        if not title:
            messagebox.showerror("Error", "Song title is required.")
            self.set_status("Save failed: title is required.")
            return

        if not self.sections:
            messagebox.showerror("Error", "Add at least one section.")
            self.set_status("Save failed: add at least one section.")
            return

        song_data = {
            "schema_version": "1.0",
            "song": {
                "title": title,
                "author": self.author_entry.get().strip(),
                "copyright": "",
                "ccli_number": "",
                "notes": "",
            },
            "structure": {"sections": self.sections},
            "chords": {"enabled": False, "sections": {}},
        }

        filename = title.lower().replace(" ", "_") + ".json"
        path = self.songs_folder / filename

        if path.exists():
            if not messagebox.askyesno("Overwrite?", f"{filename} already exists. Replace it?"):
                self.set_status("Save cancelled.")
                return

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(song_data, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Saved", f"Song saved:\n{filename}")
            self.set_status(f"Saved {filename}.")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.set_status("Save failed.")

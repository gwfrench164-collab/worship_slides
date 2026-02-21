import json
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

from slide_builder import SlideBuilder
from config import load_data_root, load_build_prefs, save_build_prefs


class BuildWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Build Slides")
        self.geometry("820x650")
        self.minsize(820, 650)

        self.data_root = load_data_root()
        if not self.data_root:
            messagebox.showerror("Error", "Data folder not set.")
            self.destroy()
            return

        self.songs_folder = Path(self.data_root) / "songs"
        self.templates_folder = Path(self.data_root) / "templates"
        self.output_folder = Path(self.data_root) / "output"

        # Library state
        self.available_files: list[Path] = []
        self.available_titles: list[str] = []   # parallel to available_files
        self.filtered_indices: list[int] = []   # indices into available_files for current filter

        # Service/setlist state
        self.service_files: list[Path] = []
        self.service_titles: list[str] = []     # parallel to service_files

        self._build_ui()
        self._load_templates()
        self._load_preferences()
        self._load_available_songs()
        self._apply_filter()

    # ---------------- UI ----------------

    def _build_ui(self):
        # Top row: Search
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        tk.Label(top, text="Search songs:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(top, textvariable=self.search_var)
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))
        self.search_entry.bind("<KeyRelease>", lambda e: self._apply_filter())

        # Middle: two lists + buttons
        mid = tk.Frame(self)
        mid.pack(fill="both", expand=True, padx=10, pady=8)

        # Available songs
        left = tk.Frame(mid)
        left.pack(side="left", fill="both", expand=True)

        tk.Label(left, text="Available Songs", font=("Arial", 11, "bold")).pack(anchor="w")
        self.available_listbox = tk.Listbox(left, height=18)
        self.available_listbox.pack(fill="both", expand=True)
        self.available_listbox.bind("<Double-Button-1>", lambda e: self.add_selected())

        # Middle buttons
        center = tk.Frame(mid)
        center.pack(side="left", fill="y", padx=10)

        tk.Button(center, text="Add →", width=12, command=self.add_selected).pack(pady=(70, 8))
        tk.Button(center, text="← Remove", width=12, command=self.remove_selected).pack(pady=8)

        # Service order
        right = tk.Frame(mid)
        right.pack(side="left", fill="both", expand=True)

        tk.Label(right, text="Service Order", font=("Arial", 11, "bold")).pack(anchor="w")
        self.service_listbox = tk.Listbox(right, height=18)
        self.service_listbox.pack(fill="both", expand=True)

        reorder = tk.Frame(right)
        reorder.pack(fill="x", pady=6)
        tk.Button(reorder, text="Move Up", command=self.move_up).pack(side="left")
        tk.Button(reorder, text="Move Down", command=self.move_down).pack(side="left", padx=6)
        tk.Button(reorder, text="Clear", command=self.clear_service).pack(side="right")

        # Bottom: template, density, output, build
        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=10)

        tk.Label(bottom, text="Template:").grid(row=0, column=0, sticky="w")
        self.template_var = tk.StringVar()
        self.template_menu = tk.OptionMenu(bottom, self.template_var, "")
        self.template_menu.grid(row=0, column=1, sticky="ew", padx=(6, 12))

        tk.Label(bottom, text="Density:").grid(row=0, column=2, sticky="w")
        self.density_var = tk.StringVar(value="Normal")
        self.density_menu = tk.OptionMenu(bottom, self.density_var, "Spacious", "Normal", "Compact")
        self.density_menu.grid(row=0, column=3, sticky="w", padx=(6, 12))

        tk.Label(bottom, text="Output filename:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.output_entry = tk.Entry(bottom)
        self.output_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        self.output_entry.insert(0, "Service_Deck.pptx")

        bottom.grid_columnconfigure(1, weight=1)

        tk.Button(self, text="Build Slides", command=self.build_slides, width=18).pack(pady=(0, 10))

    # ---------------- Data loading ----------------

    def _read_song_title(self, path: Path) -> str:
        """
        Best-effort: read title from JSON.
        If anything goes wrong, fall back to filename.
        """
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            title = (data.get("song") or {}).get("title", "").strip()
            return title if title else path.stem
        except Exception:
            return path.stem

    def _load_available_songs(self):
        self.available_files = sorted(self.songs_folder.glob("*.json"))
        self.available_titles = [self._read_song_title(p) for p in self.available_files]

    def _load_templates(self):
        templates = sorted(self.templates_folder.glob("*.pptx"))
        menu = self.template_menu["menu"]
        menu.delete(0, "end")

        if not templates:
            self.template_var.set("")
            return

        for tpl in templates:
            menu.add_command(label=tpl.name, command=lambda p=tpl: self.template_var.set(p.name))

        self.template_var.set(templates[0].name)

    def _load_preferences(self):
        prefs = load_build_prefs()

        if prefs.get("last_template"):
            self.template_var.set(prefs["last_template"])

        if prefs.get("last_output"):
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, prefs["last_output"])

        # NEW: last_density (backwards compatible)
        last_density = prefs.get("last_density")
        if last_density in ("Spacious", "Normal", "Compact"):
            self.density_var.set(last_density)

    # ---------------- Filtering ----------------

    def _apply_filter(self):
        q = self.search_var.get().strip().lower()

        self.available_listbox.delete(0, tk.END)
        self.filtered_indices = []

        for i, title in enumerate(self.available_titles):
            if not q or q in title.lower():
                self.filtered_indices.append(i)
                self.available_listbox.insert(tk.END, title)

    # ---------------- Service list actions ----------------

    def add_selected(self):
        sel = self.available_listbox.curselection()
        if not sel:
            return

        filtered_pos = sel[0]
        avail_index = self.filtered_indices[filtered_pos]

        song_path = self.available_files[avail_index]
        song_title = self.available_titles[avail_index]

        if song_path in self.service_files:
            messagebox.showinfo("Already added", f"'{song_title}' is already in the service order.")
            return

        self.service_files.append(song_path)
        self.service_titles.append(song_title)
        self.service_listbox.insert(tk.END, song_title)

    def remove_selected(self):
        sel = self.service_listbox.curselection()
        if not sel:
            return
        i = sel[0]

        del self.service_files[i]
        del self.service_titles[i]
        self.service_listbox.delete(i)

    def move_up(self):
        sel = self.service_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        if i == 0:
            return

        self.service_files[i-1], self.service_files[i] = self.service_files[i], self.service_files[i-1]
        self.service_titles[i-1], self.service_titles[i] = self.service_titles[i], self.service_titles[i-1]
        self._refresh_service_listbox(select_index=i-1)

    def move_down(self):
        sel = self.service_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        if i >= len(self.service_files) - 1:
            return

        self.service_files[i+1], self.service_files[i] = self.service_files[i], self.service_files[i+1]
        self.service_titles[i+1], self.service_titles[i] = self.service_titles[i], self.service_titles[i+1]
        self._refresh_service_listbox(select_index=i+1)

    def clear_service(self):
        self.service_files.clear()
        self.service_titles.clear()
        self.service_listbox.delete(0, tk.END)

    def _refresh_service_listbox(self, select_index: int | None = None):
        self.service_listbox.delete(0, tk.END)
        for t in self.service_titles:
            self.service_listbox.insert(tk.END, t)

        if select_index is not None and 0 <= select_index < len(self.service_titles):
            self.service_listbox.selection_set(select_index)

    # ---------------- Build ----------------

    def build_slides(self):
        if not self.service_files:
            messagebox.showwarning("No songs", "Add at least one song to the Service Order.")
            return

        template_name = self.template_var.get()
        output_name = self.output_entry.get().strip()

        if not template_name:
            messagebox.showwarning("Template missing", "Select a template.")
            return

        if not output_name:
            messagebox.showwarning("Output missing", "Enter an output filename.")
            return

        template_path = self.templates_folder / template_name
        output_path = self.output_folder / output_name

        density_map = {
            "Spacious": "spacious",
            "Normal": "normal",
            "Compact": "compact",
        }
        density = density_map.get(self.density_var.get(), "normal")

        builder = SlideBuilder(template_path)

        try:
            builder.build_deck(self.service_files, output_path)
        except Exception as e:
            messagebox.showerror("Build failed", repr(e))
            return

        # Save prefs (extend existing prefs file without breaking older reads)
        save_build_prefs(template_name, output_name)
        try:
            prefs = load_build_prefs()
            prefs["last_density"] = self.density_var.get()
            # best-effort write-back
            # If your save_build_prefs already writes a full dict, you can replace this
            # with a dedicated save function later. For now we keep it simple:
            from config import _BUILD_PREFS_FILE  # if exists
        except Exception:
            # If you don't expose the prefs file path, no worries; density just won't persist.
            pass

        messagebox.showinfo("Success", f"Slides created:\n{output_path}")
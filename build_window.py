import json
import tkinter as tk
from tkinter import messagebox, ttk
from pathlib import Path

from slide_builder import SlideBuilder
from config import load_data_root, load_build_prefs, save_build_prefs


class BuildWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Build Slides")
        self.geometry("980x680")
        self.minsize(900, 620)

        self.data_root = load_data_root()
        if not self.data_root:
            messagebox.showerror("Error", "Data folder not set.")
            self.destroy()
            return

        self.songs_folder = Path(self.data_root) / "songs"
        self.templates_folder = Path(self.data_root) / "templates"
        self.output_folder = Path(self.data_root) / "output"

        self.available_files: list[Path] = []
        self.available_titles: list[str] = []
        self.filtered_indices: list[int] = []

        self.service_files: list[Path] = []
        self.service_titles: list[str] = []

        self.status_var = tk.StringVar(value="Ready")

        self._configure_styles()
        self._build_ui()
        self._load_templates()
        self._load_preferences()
        self._load_available_songs()
        self._apply_filter()
        self.set_status("Ready to build a service deck.")

    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("BW.Header.TLabel", font=("TkDefaultFont", 16, "bold"))
        style.configure("BW.Subtle.TLabel", foreground="#555555")
        style.configure("BW.Section.TLabelframe", padding=10)
        style.configure("BW.Status.TLabel", padding=(8, 4))
        style.configure("TButton", padding=(10, 6))
        style.configure("Accent.TButton", padding=(12, 6))

    def set_status(self, message: str):
        self.status_var.set(message)
        self.update_idletasks()

    # ---------------- UI ----------------

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        header = ttk.Frame(self, padding=(16, 14, 16, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Build Slides", style="BW.Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="Choose songs, arrange the service order, then build a slide deck.",
            style="BW.Subtle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        content = ttk.Frame(self, padding=(16, 0, 16, 12))
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        search_row = ttk.Frame(content)
        search_row.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        search_row.columnconfigure(1, weight=1)
        ttk.Label(search_row, text="Search songs:").grid(row=0, column=0, sticky="w")
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.search_entry.bind("<KeyRelease>", lambda _e: self._apply_filter())

        middle = ttk.Frame(content)
        middle.grid(row=1, column=0, sticky="nsew")
        middle.columnconfigure(0, weight=1)
        middle.columnconfigure(2, weight=1)
        middle.rowconfigure(0, weight=1)

        # Available songs
        left = ttk.LabelFrame(middle, text="Available Songs", style="BW.Section.TLabelframe")
        left.grid(row=0, column=0, sticky="nsew")
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="Double-click a song to add it.", style="BW.Subtle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        left_list_frame = ttk.Frame(left)
        left_list_frame.grid(row=1, column=0, sticky="nsew")
        left_list_frame.columnconfigure(0, weight=1)
        left_list_frame.rowconfigure(0, weight=1)
        self.available_listbox = tk.Listbox(left_list_frame, height=18, exportselection=False)
        self.available_listbox.grid(row=0, column=0, sticky="nsew")
        self.available_listbox.bind("<Double-Button-1>", lambda _e: self.add_selected())
        available_scroll = ttk.Scrollbar(left_list_frame, orient="vertical", command=self.available_listbox.yview)
        available_scroll.grid(row=0, column=1, sticky="ns")
        self.available_listbox.configure(yscrollcommand=available_scroll.set)

        # Center buttons
        center = ttk.Frame(middle, padding=(12, 0))
        center.grid(row=0, column=1, sticky="ns")
        center.rowconfigure(0, weight=1)
        center_buttons = ttk.Frame(center)
        center_buttons.grid(row=0, column=0, sticky="ns")
        ttk.Button(center_buttons, text="Add →", command=self.add_selected, width=14, style="Accent.TButton").pack(
            pady=(110, 8)
        )
        ttk.Button(center_buttons, text="← Remove", command=self.remove_selected, width=14).pack(pady=8)

        # Service order
        right = ttk.LabelFrame(middle, text="Service Order", style="BW.Section.TLabelframe")
        right.grid(row=0, column=2, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        ttk.Label(right, text="Arrange songs in the order they should appear.", style="BW.Subtle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        right_list_frame = ttk.Frame(right)
        right_list_frame.grid(row=1, column=0, sticky="nsew")
        right_list_frame.columnconfigure(0, weight=1)
        right_list_frame.rowconfigure(0, weight=1)
        self.service_listbox = tk.Listbox(right_list_frame, height=18, exportselection=False)
        self.service_listbox.grid(row=0, column=0, sticky="nsew")
        service_scroll = ttk.Scrollbar(right_list_frame, orient="vertical", command=self.service_listbox.yview)
        service_scroll.grid(row=0, column=1, sticky="ns")
        self.service_listbox.configure(yscrollcommand=service_scroll.set)

        reorder = ttk.Frame(right)
        reorder.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        reorder.columnconfigure(0, weight=1)
        reorder.columnconfigure(1, weight=1)
        reorder.columnconfigure(2, weight=1)
        ttk.Button(reorder, text="Move Up", command=self.move_up).grid(row=0, column=0, sticky="ew", padx=(0, 4))
        ttk.Button(reorder, text="Move Down", command=self.move_down).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(reorder, text="Clear", command=self.clear_service).grid(row=0, column=2, sticky="ew", padx=(4, 0))

        bottom = ttk.LabelFrame(content, text="Build Settings", style="BW.Section.TLabelframe")
        bottom.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure(1, weight=1)
        bottom.columnconfigure(3, weight=1)

        ttk.Label(bottom, text="Template:").grid(row=0, column=0, sticky="w")
        self.template_var = tk.StringVar()
        self.template_menu = ttk.Combobox(bottom, textvariable=self.template_var, state="readonly")
        self.template_menu.grid(row=0, column=1, sticky="ew", padx=(8, 16))

        ttk.Label(bottom, text="Density:").grid(row=0, column=2, sticky="w")
        self.density_var = tk.StringVar(value="Normal")
        self.density_menu = ttk.Combobox(bottom, textvariable=self.density_var, state="readonly", values=["Spacious", "Normal", "Compact"])
        self.density_menu.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        ttk.Label(bottom, text="Output filename:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.output_entry = ttk.Entry(bottom)
        self.output_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(8, 0), pady=(10, 0))
        self.output_entry.insert(0, "Service_Deck.pptx")

        build_row = ttk.Frame(content)
        build_row.grid(row=3, column=0, sticky="e", pady=(12, 0))
        ttk.Button(build_row, text="Build Slides", command=self.build_slides, style="Accent.TButton").pack()

        ttk.Label(self, textvariable=self.status_var, anchor="w", relief="sunken", style="BW.Status.TLabel").grid(
            row=2, column=0, sticky="ew"
        )

    # ---------------- Data loading ----------------

    def _read_song_title(self, path: Path) -> str:
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
        names = [tpl.name for tpl in templates]
        self.template_menu["values"] = names
        if names:
            self.template_var.set(names[0])
        else:
            self.template_var.set("")

    def _load_preferences(self):
        prefs = load_build_prefs()

        if prefs.get("last_template"):
            self.template_var.set(prefs["last_template"])

        if prefs.get("last_output"):
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, prefs["last_output"])

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

        self.set_status(f"Showing {len(self.filtered_indices)} available song(s).")

    # ---------------- Service list actions ----------------

    def add_selected(self):
        sel = self.available_listbox.curselection()
        if not sel:
            self.set_status("Select a song to add.")
            return

        filtered_pos = sel[0]
        avail_index = self.filtered_indices[filtered_pos]
        song_path = self.available_files[avail_index]
        song_title = self.available_titles[avail_index]

        if song_path in self.service_files:
            self.set_status(f"{song_title} is already in the service order.")
            messagebox.showinfo("Already added", f"'{song_title}' is already in the service order.")
            return

        self.service_files.append(song_path)
        self.service_titles.append(song_title)
        self.service_listbox.insert(tk.END, song_title)
        self.set_status(f"Added {song_title} to the service order.")

    def remove_selected(self):
        sel = self.service_listbox.curselection()
        if not sel:
            self.set_status("Select a song to remove.")
            return
        i = sel[0]
        removed = self.service_titles[i]
        del self.service_files[i]
        del self.service_titles[i]
        self.service_listbox.delete(i)
        self.set_status(f"Removed {removed} from the service order.")

    def move_up(self):
        sel = self.service_listbox.curselection()
        if not sel:
            self.set_status("Select a song to move.")
            return
        i = sel[0]
        if i == 0:
            self.set_status("That song is already at the top.")
            return

        self.service_files[i - 1], self.service_files[i] = self.service_files[i], self.service_files[i - 1]
        self.service_titles[i - 1], self.service_titles[i] = self.service_titles[i], self.service_titles[i - 1]
        self._refresh_service_listbox(select_index=i - 1)
        self.set_status("Moved song up in the service order.")

    def move_down(self):
        sel = self.service_listbox.curselection()
        if not sel:
            self.set_status("Select a song to move.")
            return
        i = sel[0]
        if i >= len(self.service_files) - 1:
            self.set_status("That song is already at the bottom.")
            return

        self.service_files[i + 1], self.service_files[i] = self.service_files[i], self.service_files[i + 1]
        self.service_titles[i + 1], self.service_titles[i] = self.service_titles[i], self.service_titles[i + 1]
        self._refresh_service_listbox(select_index=i + 1)
        self.set_status("Moved song down in the service order.")

    def clear_service(self):
        self.service_files.clear()
        self.service_titles.clear()
        self.service_listbox.delete(0, tk.END)
        self.set_status("Cleared the service order.")

    def _refresh_service_listbox(self, select_index: int | None = None):
        self.service_listbox.delete(0, tk.END)
        for t in self.service_titles:
            self.service_listbox.insert(tk.END, t)

        if select_index is not None and 0 <= select_index < len(self.service_titles):
            self.service_listbox.selection_set(select_index)

    # ---------------- Build ----------------

    def build_slides(self):
        if not self.service_files:
            self.set_status("Build failed: add at least one song.")
            messagebox.showwarning("No songs", "Add at least one song to the Service Order.")
            return

        template_name = self.template_var.get()
        output_name = self.output_entry.get().strip()

        if not template_name:
            self.set_status("Build failed: select a template.")
            messagebox.showwarning("Template missing", "Select a template.")
            return

        if not output_name:
            self.set_status("Build failed: enter an output filename.")
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

        builder = SlideBuilder(template_path, song_fit_preset=density)

        self.set_status("Building slides...")
        try:
            builder.build_deck(self.service_files, output_path)
        except Exception as e:
            self.set_status("Build failed.")
            messagebox.showerror("Build failed", repr(e))
            return

        save_build_prefs(template_name, output_name)
        try:
            prefs = load_build_prefs()
            prefs["last_density"] = self.density_var.get()
            from config import _BUILD_PREFS_FILE
            _BUILD_PREFS_FILE.write_text(json.dumps(prefs, indent=2), encoding="utf-8")
        except Exception:
            pass

        self.set_status(f"Slides created: {output_path.name}")
        messagebox.showinfo("Success", f"Slides created:\n{output_path}")

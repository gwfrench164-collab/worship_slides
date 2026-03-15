import json
import shutil
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
from pathlib import Path
import re

from song_builder import SongBuilder


def _slugify_title(title: str) -> str:
    """
    Convert a title into a safe filename base, e.g.:
    "The Lily of the Valley" -> "the_lily_of_the_valley"
    """
    s = title.strip().lower()
    s = re.sub(r"[^\w\s-]", "", s)
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"_+", "_", s)
    return s.strip("_") or "song"


def _read_song_title(path: Path) -> str:
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return (data.get("song", {}) or {}).get("title", "") or path.stem
    except Exception:
        return path.stem


class LibraryWindow(tk.Toplevel):
    def __init__(self, parent, songs_folder: Path):
        super().__init__(parent)
        self.title("Song Library")
        self.geometry("900x620")
        self.minsize(860, 560)

        self.parent = parent
        self.songs_folder = Path(songs_folder)
        self.all_song_paths: list[Path] = []
        self.filtered_song_paths: list[Path] = []

        self.status_var = tk.StringVar(value="Ready")
        self.search_var = tk.StringVar()
        self.song_count_var = tk.StringVar(value="0 songs")

        self._build_ui()
        self._load_songs()
        self._refresh_list()

    # ---------------- UI ----------------

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=(14, 14, 14, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Song Library", font=("TkDefaultFont", 15, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text="Search, edit, duplicate, rename, or delete songs from your library.",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        toolbar = ttk.Frame(self, padding=(14, 0, 14, 10))
        toolbar.grid(row=1, column=0, sticky="ew")
        toolbar.columnconfigure(1, weight=1)

        ttk.Label(toolbar, text="Search:").grid(row=0, column=0, sticky="w")
        self.search_entry = ttk.Entry(toolbar, textvariable=self.search_var, width=42)
        self.search_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8))
        self.search_entry.bind("<KeyRelease>", lambda e: self._refresh_list())

        ttk.Button(toolbar, text="Refresh", command=self._refresh_all).grid(row=0, column=2, padx=(0, 8))
        ttk.Label(toolbar, textvariable=self.song_count_var).grid(row=0, column=3, sticky="e")

        content = ttk.Frame(self, padding=(14, 0, 14, 10))
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        list_group = ttk.LabelFrame(content, text="Songs", padding=10)
        list_group.grid(row=0, column=0, sticky="nsew")
        list_group.columnconfigure(0, weight=1)
        list_group.rowconfigure(0, weight=1)

        list_frame = ttk.Frame(list_group)
        list_frame.grid(row=0, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, activestyle="dotbox", exportselection=False)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        self.listbox.bind("<Double-Button-1>", lambda e: self.open_selected())
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._update_status_for_selection())

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.config(yscrollcommand=scrollbar.set)

        actions = ttk.LabelFrame(content, text="Actions", padding=10)
        actions.grid(row=0, column=1, sticky="ns", padx=(12, 0))

        for i in range(5):
            actions.rowconfigure(i, weight=0)
        actions.columnconfigure(0, weight=1)

        ttk.Button(actions, text="Open / Edit", width=20, command=self.open_selected).grid(row=0, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Duplicate", width=20, command=self.duplicate_selected).grid(row=1, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Rename Title", width=20, command=self.rename_title_selected).grid(row=2, column=0, sticky="ew", pady=4)
        ttk.Button(actions, text="Delete", width=20, command=self.delete_selected).grid(row=3, column=0, sticky="ew", pady=4)
        ttk.Separator(actions, orient="horizontal").grid(row=4, column=0, sticky="ew", pady=10)
        ttk.Button(actions, text="Close", width=20, command=self.destroy).grid(row=5, column=0, sticky="ew", pady=4)

        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=(8, 4))
        status.grid(row=3, column=0, sticky="ew")

    # ---------------- Status helpers ----------------

    def set_status(self, message: str):
        self.status_var.set(message)

    def _update_song_count(self):
        count = len(self.filtered_song_paths)
        self.song_count_var.set(f"{count} song{'s' if count != 1 else ''}")

    def _update_status_for_selection(self):
        path = self._get_selected_path(show_warning=False)
        if not path:
            self.set_status("Ready")
            return
        title = _read_song_title(path)
        self.set_status(f"Selected: {title}")

    # ---------------- Data ----------------

    def _load_songs(self):
        self.all_song_paths = sorted(self.songs_folder.glob("*.json"))

    def _refresh_all(self):
        self._load_songs()
        self._refresh_list()
        self.set_status("Library refreshed.")

    def _refresh_list(self):
        query = self.search_var.get().strip().lower()

        items = []
        self.filtered_song_paths = []

        for path in self.all_song_paths:
            title = _read_song_title(path)
            display = f"{title}  —  {path.name}"
            haystack = (title + " " + path.name).lower()

            if query and query not in haystack:
                continue

            items.append(display)
            self.filtered_song_paths.append(path)

        self.listbox.delete(0, tk.END)
        for item in items:
            self.listbox.insert(tk.END, item)

        self._update_song_count()
        if query:
            self.set_status(f"Showing {len(items)} matching song(s).")
        else:
            self.set_status(f"Loaded {len(items)} song(s).")

    def _get_selected_path(self, show_warning: bool = True) -> Path | None:
        sel = self.listbox.curselection()
        if not sel:
            if show_warning:
                messagebox.showwarning("No selection", "Select a song first.")
                self.set_status("No song selected.")
            return None
        idx = sel[0]
        if idx < 0 or idx >= len(self.filtered_song_paths):
            return None
        return self.filtered_song_paths[idx]

    # ---------------- Actions ----------------

    def open_selected(self):
        path = self._get_selected_path()
        if not path:
            return

        title = _read_song_title(path)
        self.set_status(f"Opening {title}...")

        # Open SongBuilder and refresh the library when it closes
        win = SongBuilder(self.parent, self.songs_folder, open_song=path)
        win.grab_set()

        def _on_close(_evt=None):
            self._refresh_all()
            self.set_status(f"Finished editing {title}.")

        win.bind("<Destroy>", _on_close)

    def duplicate_selected(self):
        src = self._get_selected_path()
        if not src:
            return

        old_title = _read_song_title(src)
        new_title = simpledialog.askstring(
            "Duplicate Song",
            "New title:",
            initialvalue=f"{old_title} (Copy)",
            parent=self,
        )
        if not new_title:
            self.set_status("Duplicate cancelled.")
            return

        base = _slugify_title(new_title)
        dest = self.songs_folder / f"{base}.json"

        # Avoid overwriting
        counter = 2
        while dest.exists():
            dest = self.songs_folder / f"{base}_{counter}.json"
            counter += 1

        try:
            shutil.copy2(src, dest)

            # Update title inside the duplicated JSON
            with open(dest, "r", encoding="utf-8") as f:
                data = json.load(f)
            data.setdefault("song", {})
            data["song"]["title"] = new_title

            with open(dest, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            messagebox.showinfo("Duplicated", f"Created:\n{dest.name}")
            self._refresh_all()
            self.set_status(f"Duplicated {old_title}.")

        except Exception as e:
            messagebox.showerror("Duplicate failed", str(e))
            self.set_status("Duplicate failed.")

    def rename_title_selected(self):
        path = self._get_selected_path()
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read JSON:\n{e}")
            self.set_status("Could not read selected song.")
            return

        old_title = (data.get("song", {}) or {}).get("title", "") or path.stem
        new_title = simpledialog.askstring("Rename Title", "New title:", initialvalue=old_title, parent=self)
        if not new_title:
            self.set_status("Rename cancelled.")
            return

        data.setdefault("song", {})
        data["song"]["title"] = new_title

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save JSON:\n{e}")
            self.set_status("Could not save title change.")
            return

        # Optional filename rename
        if messagebox.askyesno("Rename file too?", "Rename the filename to match the new title?"):
            base = _slugify_title(new_title)
            new_path = self.songs_folder / f"{base}.json"
            counter = 2
            while new_path.exists() and new_path != path:
                new_path = self.songs_folder / f"{base}_{counter}.json"
                counter += 1

            try:
                path.rename(new_path)
            except Exception as e:
                messagebox.showwarning("Filename not changed", f"Title updated, but rename failed:\n{e}")

        self._refresh_all()
        self.set_status(f"Renamed song to {new_title}.")

    def delete_selected(self):
        path = self._get_selected_path()
        if not path:
            return

        title = _read_song_title(path)
        if not messagebox.askyesno("Delete song?", f"Delete:\n{title}\n\nFile: {path.name}"):
            self.set_status("Delete cancelled.")
            return

        try:
            path.unlink()
            self._refresh_all()
            self.set_status(f"Deleted {title}.")
        except Exception as e:
            messagebox.showerror("Delete failed", str(e))
            self.set_status("Delete failed.")

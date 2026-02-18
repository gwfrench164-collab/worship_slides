import json
import shutil
import tkinter as tk
from tkinter import messagebox, simpledialog
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
        self.geometry("820x600")
        self.minsize(820, 600)

        self.parent = parent
        self.songs_folder = Path(songs_folder)
        self.all_song_paths: list[Path] = []
        self.filtered_song_paths: list[Path] = []

        self._build_ui()
        self._load_songs()
        self._refresh_list()

    # ---------------- UI ----------------

    def _build_ui(self):
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        tk.Label(top, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(top, textvariable=self.search_var, width=40)
        self.search_entry.pack(side="left", padx=8)
        self.search_entry.bind("<KeyRelease>", lambda e: self._refresh_list())

        tk.Button(top, text="Refresh", command=self._refresh_all).pack(side="left", padx=8)

        main = tk.Frame(self)
        main.pack(fill="both", expand=True, padx=10, pady=5)

        # List + scrollbar
        list_frame = tk.Frame(main)
        list_frame.pack(side="left", fill="both", expand=True)

        self.listbox = tk.Listbox(list_frame)
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<Double-Button-1>", lambda e: self.open_selected())

        scrollbar = tk.Scrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Buttons
        buttons = tk.Frame(main)
        buttons.pack(side="left", fill="y", padx=(10, 0))

        tk.Button(buttons, text="Open / Edit", width=18, command=self.open_selected).pack(pady=4)
        tk.Button(buttons, text="Duplicate", width=18, command=self.duplicate_selected).pack(pady=4)
        tk.Button(buttons, text="Rename Title", width=18, command=self.rename_title_selected).pack(pady=4)
        tk.Button(buttons, text="Delete", width=18, command=self.delete_selected).pack(pady=4)

        tk.Label(buttons, text="").pack(pady=10)
        tk.Button(buttons, text="Close", width=18, command=self.destroy).pack(pady=4)

    # ---------------- Data ----------------

    def _load_songs(self):
        self.all_song_paths = sorted(self.songs_folder.glob("*.json"))

    def _refresh_all(self):
        self._load_songs()
        self._refresh_list()

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

    def _get_selected_path(self) -> Path | None:
        sel = self.listbox.curselection()
        if not sel:
            messagebox.showwarning("No selection", "Select a song first.")
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

        # Open SongBuilder and refresh the library when it closes
        win = SongBuilder(self.parent, self.songs_folder, open_song=path)
        win.grab_set()

        def _on_close(_evt=None):
            self._refresh_all()

        win.bind("<Destroy>", _on_close)

    def duplicate_selected(self):
        src = self._get_selected_path()
        if not src:
            return

        old_title = _read_song_title(src)
        new_title = simpledialog.askstring(
            "Duplicate Song",
            "New title:",
            initialvalue=f"{old_title} (Copy)"
        )
        if not new_title:
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

        except Exception as e:
            messagebox.showerror("Duplicate failed", str(e))

    def rename_title_selected(self):
        path = self._get_selected_path()
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read JSON:\n{e}")
            return

        old_title = (data.get("song", {}) or {}).get("title", "") or path.stem
        new_title = simpledialog.askstring("Rename Title", "New title:", initialvalue=old_title)
        if not new_title:
            return

        data.setdefault("song", {})
        data["song"]["title"] = new_title

        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save JSON:\n{e}")
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

    def delete_selected(self):
        path = self._get_selected_path()
        if not path:
            return

        title = _read_song_title(path)
        if not messagebox.askyesno("Delete song?", f"Delete:\n{title}\n\nFile: {path.name}"):
            return

        try:
            path.unlink()
            self._refresh_all()
        except Exception as e:
            messagebox.showerror("Delete failed", str(e))
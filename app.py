from pathlib import Path
from config import load_data_root, ensure_data_root_structure
from first_run import FirstRunWindow
from main_window import MainWindow


def main():
    data_root = load_data_root()
    ensure_data_root_structure(data_root)

    app = MainWindow()
    app.withdraw()

    def start_main_app():
        app.deiconify()

    if not data_root or not Path(data_root).exists():
        FirstRunWindow(app, on_complete=start_main_app)
    else:
        start_main_app()

    app.mainloop()


if __name__ == "__main__":
    main()
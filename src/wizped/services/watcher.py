import time
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from .sync import sync_sqlite_to_excel
from ..core.db import DB_NAME


class DBEventHandler(FileSystemEventHandler):
    """Handles file system events for the SQLite database."""

    def __init__(self, db_path, excel_path=None):
        self.db_path = str(db_path)
        self.excel_path = excel_path
        self.last_sync = 0
        self.debounce_seconds = 1.0

    def on_modified(self, event):
        if event.is_directory:
            return

        if str(event.src_path).endswith(DB_NAME):
            current_time = time.time()
            if current_time - self.last_sync > self.debounce_seconds:
                print(f"\nDetectada alteração: {event.src_path}")
                try:
                    sync_sqlite_to_excel(self.excel_path)
                    print("Sync OK.")
                    self.last_sync = current_time
                except Exception as e:
                    print(f"Erro no Sync: {e}")


def start_watcher(excel_path=None):
    """Starts the file watcher."""
    # DB is in root now, relative to CWD
    path = Path.cwd()
    db_path = path / DB_NAME

    event_handler = DBEventHandler(db_path, excel_path)
    observer = Observer()
    observer.schedule(event_handler, str(path), recursive=False)
    observer.start()

    print(f"Monitorando: {db_path}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

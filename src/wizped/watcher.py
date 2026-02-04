import time

from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from .mirror import sync_sqlite_to_excel
from .db import DB_NAME


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

        # Check if the modified file is our database
        # Watchdog might return different paths depending on OS/setup,
        # so checking if it ends with our DB name is a safe bet for this directory watch.
        if str(event.src_path).endswith(DB_NAME):
            current_time = time.time()
            if current_time - self.last_sync > self.debounce_seconds:
                print(f"\nDetectada alteração no banco de dados: {event.src_path}")
                print("Iniciando sincronização automática...")
                try:
                    sync_sqlite_to_excel(self.excel_path)
                    print("Sincronização concluída.")
                    self.last_sync = current_time
                except Exception as e:
                    print(f"Erro na sincronização automática: {e}")


def start_watcher(excel_path=None):
    """Starts the file watcher monitoring the current directory for database changes."""
    path = Path.cwd()
    db_path = path / DB_NAME

    event_handler = DBEventHandler(db_path, excel_path)
    observer = Observer()
    observer.schedule(event_handler, str(path), recursive=False)
    observer.start()

    print(f"Monitorando alterações em: {db_path}")
    print("Pressione Ctrl+C para parar.")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()

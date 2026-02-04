import xlwings as xw
import pandas as pd
from .db import get_db


def sync_sqlite_to_excel(excel_path=None):
    """Reads tables from SQLite and pushes them to Excel as properly formatted Tables."""
    db = get_db()

    # Connect to Excel
    if excel_path:
        import os

        # Ensure absolute path for xlwings if needed, though usually handles relative fine
        # But best to be safe if running from different CWD
        abs_path = os.path.abspath(excel_path)
        print(f"Abrindo arquivo Excel: {abs_path}")
        wb = xw.Book(abs_path)
    else:
        # Try to connect to active book, otherwise create a new one
        try:
            wb = xw.books.active
        except (
            FileNotFoundError
        ):  # No active book (unlikely if Excel is not running, but acts as check)
            wb = xw.Book()

        if not wb:
            wb = xw.Book()

    print(f"Conectado ao Excel: {wb.name}")

    # List of tables to sync
    tables_to_sync = ["clientes", "produtos"]

    for table_name in tables_to_sync:
        if table_name not in db.table_names():
            print(f"Aviso: Tabela '{table_name}' não encontrada no banco de dados.")
            continue

        # Read data into DataFrame
        df = pd.DataFrame(list(db[table_name].rows))

        # Select or Create Sheet
        if table_name in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets[table_name]
            sheet.clear_contents()  # Clear old data logic can be improved to keep formulas
        else:
            sheet = wb.sheets.add(table_name)

        # Write generic plain list first
        # Ideally we want a real ListObject (Smart Table)
        # xlwings makes this easy

        # Clear existing tables if any on the sheet to avoid collision
        # (Naive approach suitable for 'Mirroring' where DB is source of truth)
        sheet.range("A1").expand().clear()

        # Dump Data
        # index=False because we don't need the pandas index, but we might want the DB PK
        sheet.range("A1").options(pd.DataFrame, index=False, header=True).value = df

        # Create/Format as Table (ListObject)
        tbl_range = sheet.range("A1").expand()

        # name string for the table in Excel
        xl_table_name = f"tbl_{table_name}"

        # Check if table object already exists and resize/update or creating new one
        # xlwings generic way to add table:
        found_table = None
        for tbl in sheet.tables:
            if tbl.name == xl_table_name:
                found_table = tbl
                break

        if found_table:
            found_table.resize(tbl_range)
        else:
            sheet.tables.add(source=tbl_range, name=xl_table_name)

        print(
            f"Sincronizado: {table_name} -> Aba '{table_name}' (Tabela: {xl_table_name})"
        )

    print("Sincronização concluída com sucesso.")

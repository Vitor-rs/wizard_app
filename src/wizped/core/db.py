import sqlite_utils
from pathlib import Path

DB_NAME = "wizped_data.db"


def get_db_path() -> Path:
    """Returns the path to the SQLite database in the current working directory."""
    # Adjusted to look in CWD since we run from root usually
    return Path.cwd() / DB_NAME


def get_db() -> sqlite_utils.Database:
    """Connects to the SQLite database."""
    return sqlite_utils.Database(get_db_path())


def init_dummy_data():
    """Initializes the database with some dummy data."""
    db = get_db()

    # Clients
    clients = db["clientes"]
    if not clients.exists():
        clients.create({"id": int, "nome": str, "email": str, "cidade": str}, pk="id")
        clients.insert_all(
            [
                {
                    "id": 1,
                    "nome": "Empresa A",
                    "email": "a@empresa.com",
                    "cidade": "SP",
                },
                {
                    "id": 2,
                    "nome": "Empresa B",
                    "email": "b@empresa.com",
                    "cidade": "RJ",
                },
            ]
        )
        print("Tabela 'clientes' inicializada.")

    # Products
    produtos = db["produtos"]
    if not produtos.exists():
        produtos.create(
            {"sku": str, "nome": str, "preco": float, "estoque": int}, pk="sku"
        )
        print("Tabela 'produtos' inicializada.")


def upsert_product(sku, nome, preco, estoque):
    db = get_db()
    db["produtos"].upsert(
        {"sku": sku, "nome": nome, "preco": float(preco), "estoque": int(estoque)},
        pk="sku",
    )
    print(f"Produto {sku} salvo.")


def delete_product(sku):
    db = get_db()
    db["produtos"].delete(sku)
    print(f"Produto {sku} removido.")

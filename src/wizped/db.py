import sqlite_utils
from pathlib import Path

DB_NAME = "wizped_data.db"


def get_db_path() -> Path:
    """Returns the path to the SQLite database in the current working directory."""
    return Path.cwd() / DB_NAME


def get_db() -> sqlite_utils.Database:
    """Connects to the SQLite database."""
    return sqlite_utils.Database(get_db_path())


def init_dummy_data():
    """Initializes the database with some dummy data for customers and products."""
    db = get_db()

    # Create Clients Table
    clients = db["clientes"]
    if not clients.exists():
        clients.create({"id": int, "nome": str, "email": str, "cidade": str}, pk="id")

        clients.insert_all(
            [
                {
                    "id": 1,
                    "nome": "Empresa A",
                    "email": "contato@empresaa.com",
                    "cidade": "São Paulo",
                },
                {
                    "id": 2,
                    "nome": "Mercado B",
                    "email": "sac@mercadob.com",
                    "cidade": "Rio de Janeiro",
                },
                {
                    "id": 3,
                    "nome": "Loja C",
                    "email": "gerencia@lojac.com",
                    "cidade": "Belo Horizonte",
                },
            ]
        )
        print("Tabela 'clientes' criada e populada.")
    else:
        print("Tabela 'clientes' já existe.")

    # Create Products Table
    produtos = db["produtos"]
    if not produtos.exists():
        produtos.create(
            {"sku": str, "nome": str, "preco": float, "estoque": int}, pk="sku"
        )

        produtos.insert_all(
            [
                {
                    "sku": "PROD-001",
                    "nome": "Cadeira de Escritório",
                    "preco": 450.00,
                    "estoque": 15,
                },
                {
                    "sku": "PROD-002",
                    "nome": "Mesa Gamer",
                    "preco": 890.90,
                    "estoque": 8,
                },
                {
                    "sku": "PROD-003",
                    "nome": "Mouse Sem Fio",
                    "preco": 120.50,
                    "estoque": 50,
                },
            ]
        )
        print("Tabela 'produtos' criada e populada.")
    else:
        print("Tabela 'produtos' já existe.")


def upsert_product(sku, nome, preco, estoque):
    """Inserts or updates a product."""
    db = get_db()
    db["produtos"].upsert(
        {"sku": sku, "nome": nome, "preco": float(preco), "estoque": int(estoque)},
        pk="sku",
    )
    print(f"Produto {sku} salvo/atualizado com sucesso.")


def delete_product(sku):
    """Deletes a product by SKU."""
    db = get_db()
    db["produtos"].delete(sku)
    print(f"Produto {sku} removido com sucesso.")

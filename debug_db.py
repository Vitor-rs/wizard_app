from src.wizped.db import get_db

db = get_db()
print("--- Últimos 5 Produtos no Banco de Dados ---")
for row in db.query("SELECT * FROM produtos ORDER BY rowid DESC LIMIT 5"):
    print(row)
print("--------------------------------------------")

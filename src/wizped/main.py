import argparse
from .db import init_dummy_data
from .mirror import sync_sqlite_to_excel
from .watcher import start_watcher


def main():
    parser = argparse.ArgumentParser(description="Wizped - Excel & SQLite Integration")
    subparsers = parser.add_subparsers(dest="command", help="Comandos disponíveis")

    # Command: init-db
    subparsers.add_parser(
        "init-db", help="Inicializa o banco de dados com dados de teste"
    )

    # Command: sync
    sync_parser = subparsers.add_parser(
        "sync", help="Sincroniza do SQLite para o Excel"
    )
    sync_parser.add_argument(
        "--file",
        help="Caminho do arquivo Excel (opcional). Se omitido, usa a planilha ativa.",
    )

    # Command: watch
    watch_parser = subparsers.add_parser(
        "watch", help="Monitora o banco de dados e sincroniza em tempo real"
    )
    watch_parser.add_argument("--file", help="Caminho do arquivo Excel (opcional).")

    # Command: save (Create/Update)
    save_parser = subparsers.add_parser("save", help="Salva ou atualiza um produto")
    save_parser.add_argument("--sku", required=True)
    save_parser.add_argument("--nome", required=True)
    save_parser.add_argument("--preco", required=True, type=float)
    save_parser.add_argument("--estoque", required=True, type=int)

    # Command: delete
    delete_parser = subparsers.add_parser("delete", help="Remove um produto")
    delete_parser.add_argument("--sku", required=True)

    args = parser.parse_args()

    if args.command == "init-db":
        print("Inicializando banco de dados...")
        init_dummy_data()
    elif args.command == "sync":
        print("Iniciando sincronização com Excel...")
        try:
            sync_sqlite_to_excel(excel_path=args.file)
        except Exception as e:
            print(f"Erro durante a sincronização: {e}")
    elif args.command == "watch":
        print("Iniciando modo Watch...")
        try:
            start_watcher(excel_path=args.file)
        except Exception as e:
            print(f"Erro no watcher: {e}")
    elif args.command == "save":
        from .db import upsert_product

        upsert_product(args.sku, args.nome, args.preco, args.estoque)
    elif args.command == "delete":
        from .db import delete_product

        delete_product(args.sku)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

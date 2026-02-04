import argparse
from .core.db import init_dummy_data, upsert_product, delete_product
from .services.sync import sync_sqlite_to_excel
from .services.watcher import start_watcher


def main():
    parser = argparse.ArgumentParser(description="Wizped CLI")
    subparsers = parser.add_subparsers(dest="command")

    # init-db
    subparsers.add_parser("init-db")

    # sync
    sync_p = subparsers.add_parser("sync")
    sync_p.add_argument("--file")

    # watch
    watch_p = subparsers.add_parser("watch")
    watch_p.add_argument("--file")

    # save
    save_p = subparsers.add_parser("save")
    save_p.add_argument("--sku", required=True)
    save_p.add_argument("--nome", required=True)
    save_p.add_argument("--preco", required=True, type=float)
    save_p.add_argument("--estoque", required=True, type=int)

    # delete
    del_p = subparsers.add_parser("delete")
    del_p.add_argument("--sku", required=True)

    args = parser.parse_args()

    if args.command == "init-db":
        init_dummy_data()
    elif args.command == "sync":
        sync_sqlite_to_excel(args.file)
    elif args.command == "watch":
        start_watcher(args.file)
    elif args.command == "save":
        upsert_product(args.sku, args.nome, args.preco, args.estoque)
        sync_sqlite_to_excel()
    elif args.command == "delete":
        delete_product(args.sku)
        sync_sqlite_to_excel()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()

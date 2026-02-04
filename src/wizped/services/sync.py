import xlwings as xw
import pandas as pd
import os
import sys
from datetime import datetime
from ..core.db import get_db

LOG_FILE = "sync_log.txt"


def log(msg):
    timestamp = datetime.now().strftime("%H:%M:%S")
    full_msg = f"[{timestamp}] {msg}"
    print(full_msg)
    try:
        with open(LOG_FILE, "a") as f:
            f.write(full_msg + "\n")
    except:
        pass


def sync_sqlite_to_excel(excel_path=None):
    """Refreshes Excel tables from SQLite data."""
    log("--- INICIANDO SYNC ---")
    db = get_db()

    wb = None

    # 1. Tentar pegar o Workbook Ativo (Melhor caso)
    try:
        if xw.apps:
            wb = xw.books.active
            log(f"Workbook Ativo Detectado: {wb.name}")
    except Exception as e:
        log(f"Erro ao detectar ativo: {e}")

    # 2. Fallback para path se fornecido
    if not wb and excel_path:
        log(f"Tentando abrir por caminho: {excel_path}")
        try:
            wb = xw.Book(excel_path)
            log(f"Aberto com sucesso: {wb.name}")
        except Exception as e:
            log(f"Erro ao abrir arquivo: {e}")

    if not wb:
        log("CRITICO: Nao foi possivel conectar ao Excel.")
        return

    tables = ["clientes", "produtos"]

    for tbl_name in tables:
        if tbl_name not in db.table_names():
            log(f"Tabela DB '{tbl_name}' nao existe. Pulando.")
            continue

        log(f"Processando tabela: {tbl_name}")
        df = pd.DataFrame(list(db[tbl_name].rows))
        log(f"Linhas carregadas do Banco: {len(df)}")

        # Selecionar/Criar Aba
        try:
            if tbl_name in [s.name for s in wb.sheets]:
                sheet = wb.sheets[tbl_name]
            else:
                sheet = wb.sheets.add(tbl_name)
                log(f"Aba '{tbl_name}' criada.")

            # --- VISUAL FEEDBACK: ATIVAR ABA ---
            # sheet.activate()
            # (Comentado para não ficar pulando na cara do usuario se ele estiver em outra aba,
            # mas util para debug se necessario. O usuario pediu 'tempo real', ver a aba piscar ajuda).

            # Limpar Header + Dados
            sheet.range("A1").expand().clear_contents()

            if not df.empty:
                # Escrever Dados
                sheet.range("A1").options(pd.DataFrame, index=False).value = df

                # Ajustar ListObject (Tabela Excel)
                rng = sheet.range("A1").expand()
                xl_tbl_name = f"tbl_{tbl_name}"

                found = False
                for t in sheet.tables:
                    if t.name == xl_tbl_name:
                        t.resize(rng)
                        found = True
                        break

                if not found:
                    sheet.tables.add(source=rng, name=xl_tbl_name)
                    log(f"Tabela Excel '{xl_tbl_name}' criada.")
                else:
                    log(f"Tabela Excel '{xl_tbl_name}' redimensionada.")

            else:
                log("DataFrame vazio. Tabela limpa.")

        except Exception as e:
            log(f"ERRO processando aba '{tbl_name}': {e}")

    log("--- SYNC CONCLUIDO ---")

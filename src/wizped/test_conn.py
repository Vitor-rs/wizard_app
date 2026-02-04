import xlwings as xw
import sys


def test_connection():
    print("--- DIAGNOSTICO XLWINGS ---")
    try:
        # Tenta listar apps
        print(f"Apps encontrados: {len(xw.apps)}")
        for app in xw.apps:
            print(f"App PID: {app.pid}")
            for book in app.books:
                print(f"  - Livro Aberto: '{book.name}' (Caminho: {book.fullname})")

        # Tenta pegar o ativo
        wb = xw.books.active
        print(f"Livro Ativo detectado: {wb.name}")

        # Teste de escrita
        sheet = wb.sheets.active
        print(f"Escrevendo teste na celula Z1 da aba '{sheet.name}'...")
        sheet.range("Z1").value = "Conexao_OK"
        print("Escrita realizada com sucesso!")

    except Exception as e:
        print(f"ERRO CRITICO: {e}")
        print(
            "Detalhes: Verifique se o Excel esta realmente aberto e se nao ha janelas de dialogo bloqueando."
        )


if __name__ == "__main__":
    test_connection()

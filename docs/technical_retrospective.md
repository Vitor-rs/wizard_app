# Resumo Técnico: Evolução do Projeto Wizped

Este documento compila os desafios enfrentados e as soluções aplicadas durante o desenvolvimento da integração Excel-Python para o projeto Wizped.

## 1. Arquitetura e Estrutura
**Desafio:** O projeto iniciou com scripts dispersos (`main.py`, `mirror.py`) e módulos VBA fragmentados (`modMain`, `ui_builder`), dificultando a manutenção e a portabilidade.
**Solução:**
- Refatoração para estrutura modular: `src/wizped/core`, `src/wizped/services`.
- Criação de um ponto de entrada único CLI: `src/wizped/cli.py`.
- Consolidação do VBA em um único gerador: `Installer.bas`.

## 2. A Ponte VBA <-> Python (Execução de Comandos)
**Desafio (Crítico):** O Excel falhava silenciosamente ao tentar executar os scripts Python.
- *Sintoma:* "O arquivo de debug não foi gerado", "Nada acontece".
- *Causa:* O uso de caminhos relativos (`..\..`) no comando `Shell` do VBA dependia do diretório atual de trabalho (`CWD`), que o Windows/Excel nem sempre define como a pasta do arquivo aberto. Além disso, problemas de "escaping" de aspas no `cmd.exe` quebravam os argumentos.

**Solução Definitiva:**
- **Resolução de Caminho Absoluto:** Utilização da biblioteca `Scripting.FileSystemObject` no VBA para calcular matematicamente o caminho absoluto do script `debug_launcher.bat` relativo à pasta do `ThisWorkbook`.
- **Execução Direta:** Abandono do `cmd /c "..."` complexo em favor da execução direta do arquivo `.bat` via `WScript.Shell`, eliminando erros de sintaxe de aspas.

## 3. Sincronização de Dados (CRUD)
**Desafio 1:** O usuário não via as atualizações na hora (feedback lento ou inexistente).
**Solução:**
- Alteração no `cli.py` para forçar a sincronização (`sync_sqlite_to_excel`) imediatamente após qualquer operação de `save` ou `delete`.
- Alteração no VBA para recarregar a `ListBox` (`UserForm_Initialize`) assim que o comando Python retorna.

**Desafio 2:** "Perda de dados" (Excel limpava dados antigos).
**Solução:**
- Esclarecimento da lógica **Database-First**: O SQLite é a fonte da verdade. O Excel é apenas um espelho. Se o banco começa vazio, ele limpa o Excel. A persistência ocorre no arquivo `.db`.

## 4. O Instalador Descartável
**Desafio:** Atualizar o código VBA no arquivo do usuário sem quebrar formulários existentes ou exigir passos manuais complexos.
**Solução:**
- Criação do conceito **`Installer.bas`**:
    - Um módulo temporário que o usuário importa.
    - A macro `InstalarCompleto` apaga/recria o formulário com o código mais recente embutido.
    - Gera um lançador (`modWizped`) separado.
    - Pode ser excluído após o uso, deixando o projeto limpo.

## 5. Sintaxe e Detalhes
**Desafio:** Erros de sintaxe no VBA (`RunPython "..."`) ao passar argumentos com muitas aspas.
**Solução:**
- Uso da função `Chr(34)` para concatenar aspas de forma limpa e legível no código injetado, evitando o "inferno de aspas" (`""""`).

---
**Estado Atual:**
O sistema opera com um fluxo robusto:
1. **INPUT:** Usuário digita no VBA.
2. **PROCESS:** VBA chama `debug_launcher.bat` (caminho absoluto) -> Python grava no SQLite.
3. **SYNC:** Python atualiza a planilha do Excel.
4. **FEEDBACK:** VBA relê a planilha e atualiza a interface visualmente.

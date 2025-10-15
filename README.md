# Google Sheets Cashback App

Este repositório contém o código completo do projeto Google Apps Script (arquivo `Código.gs`) e o frontend HTML (`index.html`) utilizado pela planilha de cashback.

## Como copiar os arquivos para o Apps Script

1. Abra sua planilha vinculada e escolha **Extensões → Apps Script**.
2. Na IDE do Apps Script:
   - Substitua o conteúdo do arquivo padrão `Code.gs` pelo conteúdo do arquivo [`Código.gs`](./Código.gs) deste repositório.
   - Crie (ou abra) um arquivo HTML chamado `index` e substitua o conteúdo pelo arquivo [`index.html`](./index.html) deste repositório.
3. Clique em **Salvar** e implante a Web App normalmente.

### Como copiar os arquivos daqui

Você pode copiar e colar diretamente pelo navegador:

1. Abra o arquivo desejado (`Código.gs` ou `index.html`) no GitHub.
2. Clique em **Raw** para ver o conteúdo puro.
3. Pressione `Ctrl+A` (ou `Cmd+A` no macOS) para selecionar tudo, seguido de `Ctrl+C`/`Cmd+C` para copiar.
4. Volte à IDE do Apps Script e use `Ctrl+V`/`Cmd+V` para substituir o conteúdo do arquivo correspondente.

Se preferir baixar o projeto, utilize **Code → Download ZIP** no GitHub, descompacte o arquivo e abra os arquivos localmente para copiar o conteúdo.

Os arquivos do repositório já estão atualizados com todas as correções discutidas. Basta copiá-los na íntegra para o Apps Script.

### Dica: se o GitHub apontar conflitos

Caso o GitHub mostre avisos de conflito ao tentar fazer merge, basta abrir os arquivos `Código.gs` e `README.md` deste repositório e substituir **todo** o conteúdo correspondente na IDE do Apps Script. Os arquivos publicados aqui já estão sem marcadores de conflito (`<<<<<<<`, `=======`, `>>>>>>>`), portanto a cópia direta garante o estado correto do código.

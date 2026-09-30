# Ferramentas PDF (Juntar, Dividir e Comprimir)

Interface gráfica (Tkinter) multifuncional que oferece uma suíte de utilitários para trabalhar com arquivos PDF. O programa permite converter arquivos do Office (Word, Excel e PowerPoint) em massa, extrair páginas específicas de documentos e comprimir PDFs.

Requer **Windows com Microsoft Office instalado** para as funcionalidades de conversão (que utilizam a automação COM do Office).

## Funcionalidades

O aplicativo é dividido em três abas principais:

### 1. Converter e Juntar PDF
- Seleção de uma pasta completa para processamento.
- **Converte automaticamente** arquivos Word (`.doc`, `.docx`), Excel (`.xls`, `.xlsx`) e PowerPoint (`.ppt`, `.pptx`) para PDF.
- Une todos os PDFs convertidos e os já existentes na pasta em um único arquivo consolidado.
- Opção de **forçar reconversão** de arquivos antigos.
- Nome do arquivo de saída configurável (padrão: `std.pdf`).

### 2. Dividir PDF (Extração de Páginas)
- Permite extrair páginas de um PDF original usando o formato padrão de impressão.
- Aceita intervalos com hífen e páginas isoladas por vírgula (ex: `1-5, 8, 11-13`).
- Remove metadados (como o Título Oculto do PDF original) para não causar confusão ao abrir o arquivo novo.

### 3. Comprimir PDF
- Reduz o tamanho do arquivo PDF aplicando compressão nos fluxos de dados internos.

### Funcionalidades Gerais
- Barra de progresso e status em tempo real.
- Log de execução exibido diretamente na interface.
- Operações assíncronas (Thread dedicada), permitindo cancelar o processamento a qualquer momento sem travar a interface.
- Botão **Abrir Resultado** que abre o Windows Explorer com o arquivo recém-gerado já selecionado (destacado).

---

## Como rodar localmente

Clone o repositório ou baixe os arquivos, e no seu terminal, instale as dependências:

```bash
pip install pywin32 pypdf
```

Para iniciar a interface:
```bash
python app_gui.py
```

---

## Como gerar o Executável (.exe)

Para distribuir o programa para pessoas que não têm Python instalado na máquina, você pode gerar um executável `standalone`. Os usuários finais só precisarão ter o Microsoft Office instalado.

Instale o PyInstaller:
```bash
pip install pyinstaller
```

Gere o executável com o seguinte comando (se você tiver um arquivo de ícone chamado `icon.ico` na mesma pasta), ou rode o arquivo "setup.py":
```bash
pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." --name FerramentasPDF app_gui.py
```
*Se não tiver ícone, basta omitir o `--icon` e o `--add-data`.*

### Dicas para o Executável:
- O arquivo final será gerado dentro da pasta `dist\`.
- A flag `--windowed` evita que a janela preta do terminal (console) apareça atrás da interface gráfica.
- O primeiro *build* (compilação) pode demorar um pouco e o `.exe` pode ficar relativamente grande, pois o PyInstaller embute o Python e todas as bibliotecas dentro dele.
- Alguns antivírus corporativos podem alertar sobre executáveis não assinados gerados pelo PyInstaller na primeira execução. Isso é um falso-positivo comum.

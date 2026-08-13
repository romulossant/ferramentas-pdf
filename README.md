# Consolidador de PDFs

Interface gráfica (Tkinter) para converter arquivos Word/Excel de uma pasta em PDF
e unir tudo (PDFs existentes + convertidos) em um único arquivo consolidado.

Requer **Windows com Microsoft Office instalado** (a conversão usa automação COM
do Word/Excel).

## Funções

- Interface gráfica: seleção de pasta por diálogo.
- Nome do PDF final configurável na tela (padrão `MANUAL_CONSOLIDADO.pdf`).
- Opção de **reconverter** arquivos mesmo se já existir um PDF de mesmo nome.
- Barra de progresso e status em tempo real.
- Log exibido na tela (além de continuar sendo salvo em `process_log.txt`
  dentro da pasta processada).
- Botão para **cancelar** o processamento em andamento.
- Botões para abrir a pasta processada e o PDF final ao terminar.
- Aviso automático se `pywin32` ou `pypdf` não estiverem instalados.

## Como rodar localmente (para testar antes de empacotar)

```bash
pip install -r requirements.txt
python app_gui.py
```

## Como gerar o .exe para distribuir

Em uma máquina Windows com Python instalado:

```bash
pip install -r requirements.txt
pyinstaller --onefile --windowed --name ConsolidadorPDF app_gui.py
```

O executável final aparece em `dist\ConsolidadorPDF.exe`. Esse é o único
arquivo que você precisa enviar para as outras pessoas — elas só precisam ter
o Microsoft Office instalado na máquina (Word e/ou Excel, conforme os tipos
de arquivo que forem converter).

Dicas:
- `--windowed` evita que uma janela de terminal (console) apareça atrás da
  interface gráfica.
- Se quiser um ícone customizado, adicione `--icon=caminho\para\icone.ico`
  ao comando.
- O primeiro build pode demorar um pouco e o `.exe` pode ficar relativamente
  grande (é normal, o PyInstaller empacota o Python inteiro).
- Alguns antivírus corporativos podem bloquear executáveis não assinados
  gerados por PyInstaller na primeira execução.

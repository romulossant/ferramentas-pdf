"""
Consolidador de PDFs - Interface Gráfica
==========================================
Percorre uma pasta escolhida pelo usuário, converte arquivos Word/Excel
para PDF (usando automação COM do Office) e une todos os PDFs
(existentes + convertidos) em um único arquivo consolidado.

Requisitos (Windows, com Microsoft Office instalado):
    pip install pywin32 pypdf

Para gerar o executável (.exe) com PyInstaller:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name ConsolidadorPDF app_gui.py

O executável final ficará em dist\\ConsolidadorPDF.exe
"""

import os
import sys
import queue
import threading
import datetime
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# --- Dependências específicas (Windows + Office) ---
# Importadas de forma protegida para que a interface ainda abra e explique
# o problema, caso o usuário rode em uma máquina sem os requisitos corretos.
try:
    import win32com.client
    import pythoncom
    COM_DISPONIVEL = True
except ImportError:
    COM_DISPONIVEL = False

try:
    from pypdf import PdfWriter
    PYPDF_DISPONIVEL = True
except ImportError:
    PYPDF_DISPONIVEL = False


# Constantes para os formatos de conversão do Office
WD_FORMAT_PDF = 17
XL_FORMAT_PDF = 0

NOME_SAIDA_PADRAO = "MANUAL_CONSOLIDADO.pdf"


def resource_path(nome_relativo):
    """
    Resolve o caminho de um recurso (ex: ícone) tanto rodando o script
    diretamente quanto rodando dentro do .exe gerado pelo PyInstaller
    (que extrai os arquivos embutidos para uma pasta temporária em
    sys._MEIPASS quando usado com --onefile).
    """
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, nome_relativo)


# ============================================================
#                LÓGICA DE PROCESSAMENTO (BACKEND)
# ============================================================

class Processador:
    """
    Encapsula toda a lógica de conversão/união, praticamente idêntica ao
    script original, mas reportando progresso e mensagens através de uma
    fila (queue), para que a thread de processamento nunca mexa
    diretamente nos widgets do Tkinter.
    """

    def __init__(self, diretorio_raiz, nome_arquivo_saida, reprocessar,
                 log_queue, cancel_event):
        self.diretorio_raiz = diretorio_raiz
        self.nome_arquivo_saida = nome_arquivo_saida or NOME_SAIDA_PADRAO
        self.reprocessar = reprocessar
        self.log_queue = log_queue
        self.cancel_event = cancel_event
        self.log_file_path = os.path.join(self.diretorio_raiz, "process_log.txt")

        self.total_convertidos = 0
        self.total_erros = 0
        self.caminho_pdf_final = None

    # -------------------- utilidades de log/progresso --------------------

    def log(self, message, log_to_file=True):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}"
        if log_to_file:
            try:
                with open(self.log_file_path, 'a', encoding='utf-8') as f:
                    f.write(entry + '\n')
            except Exception as e:
                entry += f"  (ERRO ao gravar log em arquivo: {e})"
        self.log_queue.put(("log", entry))

    def progresso(self, atual, total):
        self.log_queue.put(("progress", (atual, total)))

    def status(self, texto):
        self.log_queue.put(("status", texto))

    # -------------------- descoberta de arquivos --------------------

    @staticmethod
    def listar_arquivos_em_dicionario(diretorio_raiz):
        arquivos_por_pasta = {}
        for caminho_atual, _subdiretorios, arquivos in os.walk(diretorio_raiz):
            if arquivos:
                arquivos_por_pasta[caminho_atual] = arquivos
        return arquivos_por_pasta

    # -------------------- conversão via COM --------------------

    def converter_para_pdf(self, caminho_completo_entrada, caminho_completo_saida):
        """Converte um arquivo Word/Excel para PDF com gerenciamento COM robusto."""
        # Normaliza os caminhos (barras consistentes, sem duplicidade) antes de
        # entregar ao Word/Excel via COM. Caminhos com barras "/" misturadas com
        # "\" fazem o Office interpretar o caminho como URL e falhar ao abrir
        # o arquivo, mesmo que ele exista.
        caminho_completo_entrada = os.path.normpath(os.path.abspath(caminho_completo_entrada))
        caminho_completo_saida = os.path.normpath(os.path.abspath(caminho_completo_saida))

        extensao = caminho_completo_entrada.split('.')[-1].lower()
        word = None
        doc = None
        excel = None
        wb = None

        pythoncom.CoInitialize()
        try:
            # ================= WORD =================
            if extensao in ['docx', 'doc']:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0  # wdAlertsNone

                word.Options.ConfirmConversions = False
                word.Options.SaveNormalPrompt = False
                word.Options.WarnBeforeSavingPrintingSendingMarkup = False

                doc = word.Documents.Open(
                    caminho_completo_entrada,
                    ReadOnly=True,
                    ConfirmConversions=False,
                    AddToRecentFiles=False,
                    NoEncodingDialog=True
                )
                doc.ExportAsFixedFormat(
                    OutputFileName=caminho_completo_saida,
                    ExportFormat=WD_FORMAT_PDF,
                    OpenAfterExport=False
                )
                return True

            # ================= EXCEL =================
            elif extensao in ['xlsx', 'xls']:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                excel.AskToUpdateLinks = False
                excel.AutomationSecurity = 3  # bloqueia macros

                wb = excel.Workbooks.Open(
                    caminho_completo_entrada,
                    ReadOnly=True,
                    UpdateLinks=0,
                    IgnoreReadOnlyRecommended=True
                )
                wb.ExportAsFixedFormat(XL_FORMAT_PDF, caminho_completo_saida)
                return True

            else:
                self.log(f"[AVISO] Extensão não suportada: {caminho_completo_entrada}")
                return False

        except Exception as e:
            self.log(f"[ERRO GRAVE] Falha ao converter {caminho_completo_entrada}: {e}")
            return False

        finally:
            if doc:
                try:
                    doc.Close(False)
                except Exception:
                    pass
                del doc
            if word:
                try:
                    word.Quit()
                except Exception:
                    pass
                del word
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
                del wb
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
                del excel
            pythoncom.CoUninitialize()

    # -------------------- união dos PDFs --------------------

    def unir_pdfs(self, lista_pdfs):
        merger = PdfWriter()
        self.caminho_pdf_final = os.path.join(self.diretorio_raiz, self.nome_arquivo_saida)

        self.log("\nIniciando união de PDFs...")
        self.log(f"Total de PDFs a serem unidos: {len(lista_pdfs)}")

        if not lista_pdfs:
            self.log("Nenhum arquivo PDF encontrado para unir.")
            return

        try:
            for pdf in sorted(lista_pdfs):
                try:
                    merger.append(pdf)
                except Exception as e:
                    self.log(f"[AVISO] Não foi possível adicionar o PDF '{pdf}' à união. Motivo: {e}")

            with open(self.caminho_pdf_final, "wb") as saida:
                merger.write(saida)

            self.log(f"\n[SUCESSO] Todos os PDFs foram unidos em: {self.caminho_pdf_final}")

        except Exception as e:
            self.log(f"[ERRO] Falha ao unir PDFs: {e}")
            self.caminho_pdf_final = None
        finally:
            merger.close()

    # -------------------- orquestração principal --------------------

    def executar(self):
        try:
            with open(self.log_file_path, 'w', encoding='utf-8') as f:
                f.write(f"--- INÍCIO DO PROCESSO EM "
                        f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        except Exception as e:
            self.log_queue.put(("log", f"Erro ao inicializar o arquivo de log: {e}"))

        self.log(f"Iniciando processo no diretório: {self.diretorio_raiz}")
        self.status("Procurando arquivos...")

        arquivos_por_pasta = self.listar_arquivos_em_dicionario(self.diretorio_raiz)

        # Conta total de arquivos "relevantes" para a barra de progresso
        arquivos_relevantes = []
        for pasta, arquivos in arquivos_por_pasta.items():
            for nome_arquivo in arquivos:
                nome_lower = nome_arquivo.lower()
                if nome_lower.endswith('.pdf') and not nome_lower.startswith('manual_consolidado'):
                    arquivos_relevantes.append((pasta, nome_arquivo))
                elif nome_lower.endswith(('.docx', '.doc', '.xlsx', '.xls')):
                    arquivos_relevantes.append((pasta, nome_arquivo))

        total = len(arquivos_relevantes)
        self.progresso(0, max(total, 1))

        lista_de_pdfs_para_unir = []

        for indice, (pasta, nome_arquivo) in enumerate(arquivos_relevantes, start=1):
            if self.cancel_event.is_set():
                self.log("\n[CANCELADO] Processo interrompido pelo usuário.")
                self.status("Cancelado.")
                self.progresso(indice, total)
                return

            caminho_entrada = os.path.join(pasta, nome_arquivo)
            nome_lower = nome_arquivo.lower()
            self.status(f"Processando ({indice}/{total}): {nome_arquivo}")

            if nome_lower.endswith('.pdf'):
                lista_de_pdfs_para_unir.append(caminho_entrada)
                self.log(f"  [EXISTENTE] PDF encontrado: {nome_arquivo}")
                self.progresso(indice, total)
                continue

            nome_base = os.path.splitext(nome_arquivo)[0]
            nome_saida = nome_base + ".pdf"
            caminho_saida = os.path.join(pasta, nome_saida)

            if os.path.exists(caminho_saida) and not self.reprocessar:
                self.log(f"  [PULAR] PDF de destino já existe para {nome_arquivo}. "
                          f"Adicionando à lista de união.")
                lista_de_pdfs_para_unir.append(caminho_saida)
                self.progresso(indice, total)
                continue

            if self.converter_para_pdf(caminho_entrada, caminho_saida):
                self.log(f"  [SUCESSO] Convertido {nome_arquivo} para {nome_saida}")
                lista_de_pdfs_para_unir.append(caminho_saida)
                self.total_convertidos += 1
            else:
                self.total_erros += 1

            self.progresso(indice, total)

        self.log("\n--- Processo de Conversão Finalizado ---")
        self.log(f"Total de arquivos convertidos com sucesso: {self.total_convertidos}")
        self.log(f"Total de falhas na conversão: {self.total_erros}")
        self.log(f"Total de PDFs coletados para união: {len(lista_de_pdfs_para_unir)}")
        self.log("---------------------------------------")

        self.status("Unindo PDFs...")
        self.unir_pdfs(lista_de_pdfs_para_unir)

        self.log(f"--- FIM DO PROCESSO EM "
                  f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        self.status("Concluído.")


# ============================================================
#                        INTERFACE GRÁFICA
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CONVERSÃO E CONSOLIDAÇÃO DE PDFs")
        self.geometry("720x560")
        self.minsize(640, 480)
        self._aplicar_icone()

        self.log_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.caminho_pdf_final = None

        self._montar_widgets()
        self._checar_dependencias()
        self.after(100, self._processar_fila)

    def _aplicar_icone(self):
        """Define o ícone da janela (título/barra de tarefas), tanto em modo
        script quanto empacotado no .exe."""
        try:
            caminho_icone = resource_path(os.path.join("src", "icon.ico"))
            if not os.path.exists(caminho_icone):
                # fallback: caso o .ico esteja na raiz junto do script/exe
                caminho_icone = resource_path("icon.ico")
            if os.path.exists(caminho_icone):
                self.iconbitmap(caminho_icone)
        except Exception:
            # Se falhar (ex: rodando fora do Windows), a janela simplesmente
            # segue sem ícone customizado, sem travar o programa.
            pass

    # -------------------- construção da interface --------------------

    def _montar_widgets(self):
        padding = {"padx": 10, "pady": 6}

        # --- Seleção de pasta ---
        frame_pasta = ttk.LabelFrame(self, text="1. Pasta a processar")
        frame_pasta.pack(fill="x", **padding)

        self.var_pasta = tk.StringVar()
        entry_pasta = ttk.Entry(frame_pasta, textvariable=self.var_pasta, state="readonly")
        entry_pasta.pack(side="left", fill="x", expand=True, padx=(10, 6), pady=8)

        btn_pasta = ttk.Button(frame_pasta, text="Selecionar pasta...",
                                command=self._selecionar_pasta)
        btn_pasta.pack(side="left", padx=(0, 10), pady=8)

        # --- Opções ---
        frame_opcoes = ttk.LabelFrame(self, text="2. Opções")
        frame_opcoes.pack(fill="x", **padding)

        linha1 = ttk.Frame(frame_opcoes)
        linha1.pack(fill="x", padx=10, pady=(8, 4))
        ttk.Label(linha1, text="Nome do PDF final:").pack(side="left")
        self.var_nome_saida = tk.StringVar(value=NOME_SAIDA_PADRAO)
        ttk.Entry(linha1, textvariable=self.var_nome_saida, width=35).pack(
            side="left", padx=(8, 0))

        linha2 = ttk.Frame(frame_opcoes)
        linha2.pack(fill="x", padx=10, pady=(0, 8))
        self.var_reprocessar = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            linha2, text="Reconverter arquivos mesmo se já existir um PDF com o mesmo nome",
            variable=self.var_reprocessar
        ).pack(side="left")

        # --- Ações ---
        frame_acoes = ttk.Frame(self)
        frame_acoes.pack(fill="x", **padding)

        self.btn_iniciar = ttk.Button(frame_acoes, text="Iniciar processamento",
                                       command=self._iniciar_processamento)
        self.btn_iniciar.pack(side="left")

        self.btn_cancelar = ttk.Button(frame_acoes, text="Cancelar",
                                        command=self._cancelar_processamento,
                                        state="disabled")
        self.btn_cancelar.pack(side="left", padx=(8, 0))

        self.btn_abrir_pasta = ttk.Button(frame_acoes, text="Abrir pasta do resultado",
                                           command=self._abrir_pasta_resultado,
                                           state="disabled")
        self.btn_abrir_pasta.pack(side="left", padx=(8, 0))

        self.btn_abrir_pdf = ttk.Button(frame_acoes, text="Abrir PDF gerado",
                                         command=self._abrir_pdf_resultado,
                                         state="disabled")
        self.btn_abrir_pdf.pack(side="left", padx=(8, 0))

        # --- Progresso ---
        frame_progresso = ttk.Frame(self)
        frame_progresso.pack(fill="x", **padding)

        self.var_status = tk.StringVar(value="Aguardando início.")
        ttk.Label(frame_progresso, textvariable=self.var_status).pack(anchor="w")

        self.progressbar = ttk.Progressbar(frame_progresso, mode="determinate")
        self.progressbar.pack(fill="x", pady=(4, 0))

        # --- Log ---
        frame_log = ttk.LabelFrame(self, text="Log de execução")
        frame_log.pack(fill="both", expand=True, **padding)

        self.text_log = scrolledtext.ScrolledText(frame_log, state="disabled", wrap="word")
        self.text_log.pack(fill="both", expand=True, padx=8, pady=8)

    def _checar_dependencias(self):
        faltando = []
        if not COM_DISPONIVEL:
            faltando.append("pywin32 (win32com / pythoncom) — necessário para converter "
                             "Word/Excel. Instale com: pip install pywin32")
        if not PYPDF_DISPONIVEL:
            faltando.append("pypdf — necessário para unir os PDFs. "
                             "Instale com: pip install pypdf")
        if faltando:
            mensagem = ("Algumas dependências não foram encontradas:\n\n- "
                         + "\n- ".join(faltando) +
                         "\n\nO programa pode não funcionar corretamente até que "
                         "sejam instaladas.")
            messagebox.showwarning("Dependências ausentes", mensagem)

    # -------------------- ações da interface --------------------

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta a ser processada")
        if pasta:
            # O diálogo devolve o caminho com barras "/" (estilo Unix); normalizamos
            # para o padrão do Windows ("\") para evitar caminhos mistos como
            # "D:/Pasta\Subpasta", que fazem o Word/Excel interpretarem o caminho
            # como URL e falharem ao abrir o arquivo.
            pasta = os.path.normpath(pasta)
            self.var_pasta.set(pasta)

    def _iniciar_processamento(self):
        pasta = self.var_pasta.get().strip()
        if not pasta:
            messagebox.showerror("Erro", "Selecione uma pasta antes de iniciar.")
            return
        if not os.path.isdir(pasta):
            messagebox.showerror("Erro", "A pasta selecionada não existe mais.")
            return
        if not COM_DISPONIVEL or not PYPDF_DISPONIVEL:
            messagebox.showerror(
                "Dependências ausentes",
                "Não é possível iniciar: dependências obrigatórias não instaladas.\n"
                "Veja o aviso exibido ao abrir o programa."
            )
            return

        nome_saida = self.var_nome_saida.get().strip() or NOME_SAIDA_PADRAO
        if not nome_saida.lower().endswith(".pdf"):
            nome_saida += ".pdf"

        # Limpa log visual e reseta estado
        self.text_log.configure(state="normal")
        self.text_log.delete("1.0", tk.END)
        self.text_log.configure(state="disabled")
        self.progressbar["value"] = 0
        self.caminho_pdf_final = None
        self.btn_abrir_pasta.configure(state="disabled")
        self.btn_abrir_pdf.configure(state="disabled")
        self.cancel_event.clear()

        self.btn_iniciar.configure(state="disabled")
        self.btn_cancelar.configure(state="normal")
        self.var_status.set("Iniciando...")

        processador = Processador(
            diretorio_raiz=pasta,
            nome_arquivo_saida=nome_saida,
            reprocessar=self.var_reprocessar.get(),
            log_queue=self.log_queue,
            cancel_event=self.cancel_event,
        )
        self._processador_atual = processador

        self.worker_thread = threading.Thread(target=self._executar_com_seguranca,
                                               args=(processador,), daemon=True)
        self.worker_thread.start()

    def _executar_com_seguranca(self, processador):
        try:
            processador.executar()
        except Exception as e:
            self.log_queue.put(("log", f"[ERRO INESPERADO] {e}"))
            self.log_queue.put(("status", "Erro."))
        finally:
            self.log_queue.put(("done", processador.caminho_pdf_final))

    def _cancelar_processamento(self):
        self.cancel_event.set()
        self.btn_cancelar.configure(state="disabled")
        self.var_status.set("Cancelando... aguarde a etapa atual terminar.")

    def _abrir_pasta_resultado(self):
        pasta = self.var_pasta.get().strip()
        if pasta and os.path.isdir(pasta):
            self._abrir_no_explorador(pasta)

    def _abrir_pdf_resultado(self):
        if self.caminho_pdf_final and os.path.exists(self.caminho_pdf_final):
            self._abrir_no_explorador(self.caminho_pdf_final)

    @staticmethod
    def _abrir_no_explorador(caminho):
        try:
            if sys.platform.startswith("win"):
                os.startfile(caminho)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", caminho], check=False)
            else:
                subprocess.run(["xdg-open", caminho], check=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir: {e}")

    # -------------------- fila de comunicação com a thread --------------------

    def _processar_fila(self):
        try:
            while True:
                tipo, valor = self.log_queue.get_nowait()

                if tipo == "log":
                    self.text_log.configure(state="normal")
                    self.text_log.insert(tk.END, valor + "\n")
                    self.text_log.see(tk.END)
                    self.text_log.configure(state="disabled")

                elif tipo == "status":
                    self.var_status.set(valor)

                elif tipo == "progress":
                    atual, total = valor
                    self.progressbar["maximum"] = total
                    self.progressbar["value"] = atual

                elif tipo == "done":
                    self.caminho_pdf_final = valor
                    self.btn_iniciar.configure(state="normal")
                    self.btn_cancelar.configure(state="disabled")
                    if valor and os.path.exists(valor):
                        self.btn_abrir_pasta.configure(state="normal")
                        self.btn_abrir_pdf.configure(state="normal")
                        messagebox.showinfo("Concluído",
                                             f"Processo finalizado.\nPDF gerado em:\n{valor}")
                    else:
                        self.btn_abrir_pasta.configure(state="normal")
                        messagebox.showwarning(
                            "Concluído com pendências",
                            "O processo terminou, mas o PDF final não foi gerado. "
                            "Verifique o log para detalhes."
                        )
        except queue.Empty:
            pass
        finally:
            self.after(100, self._processar_fila)


if __name__ == "__main__":
    app = App()
    app.mainloop()

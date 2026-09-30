import os
import sys
import queue
import threading
import datetime
import subprocess

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

try:
    import win32com.client
    import pythoncom
    COM_DISPONIVEL = True
except ImportError:
    COM_DISPONIVEL = False

try:
    from pypdf import PdfWriter, PdfReader
    PYPDF_DISPONIVEL = True
except ImportError:
    PYPDF_DISPONIVEL = False

# Constantes para os formatos de conversão do Office
WD_FORMAT_PDF = 17
XL_FORMAT_PDF = 0
PP_FORMAT_PDF = 32  # Constante ppSaveAsPDF do PowerPoint

NOME_SAIDA_PADRAO = "std.pdf"


def resource_path(nome_relativo):
    """Resolve o caminho de um recurso para suportar o empacotamento do PyInstaller."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, nome_relativo)


# ============================================================
#                LÓGICA DE PROCESSAMENTO (BACKEND)
# ============================================================

class ProcessadorBase:
    """Classe base para padronizar a comunicação com a interface gráfica."""
    def __init__(self, log_queue, cancel_event):
        self.log_queue = log_queue
        self.cancel_event = cancel_event
        self.caminho_resultado = None

    def log(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] {message}"
        self.log_queue.put(("log", entry))

    def progresso(self, atual, total):
        self.log_queue.put(("progress", (atual, total)))

    def status(self, texto):
        self.log_queue.put(("status", texto))

    def executar(self):
        raise NotImplementedError("As subclasses devem implementar o método executar().")


class ProcessadorConsolidacao(ProcessadorBase):
    def __init__(self, diretorio_raiz, nome_arquivo_saida, reprocessar, log_queue, cancel_event):
        super().__init__(log_queue, cancel_event)
        self.diretorio_raiz = diretorio_raiz
        self.nome_arquivo_saida = nome_arquivo_saida or NOME_SAIDA_PADRAO
        self.reprocessar = reprocessar
        
        self.total_convertidos = 0
        self.total_erros = 0

    @staticmethod
    def listar_arquivos_em_dicionario(diretorio_raiz):
        arquivos_por_pasta = {}
        for caminho_atual, _subdiretorios, arquivos in os.walk(diretorio_raiz):
            if arquivos:
                arquivos_por_pasta[caminho_atual] = arquivos
        return arquivos_por_pasta

    def converter_para_pdf(self, caminho_completo_entrada, caminho_completo_saida):
        caminho_completo_entrada = os.path.normpath(os.path.abspath(caminho_completo_entrada))
        caminho_completo_saida = os.path.normpath(os.path.abspath(caminho_completo_saida))
        extensao = caminho_completo_entrada.split('.')[-1].lower()
        
        word, doc, excel, wb, ppt, presentation = None, None, None, None, None, None

        pythoncom.CoInitialize()
        try:
            # --- WORD ---
            if extensao in ['docx', 'doc']:
                word = win32com.client.DispatchEx("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0
                doc = word.Documents.Open(caminho_completo_entrada, ReadOnly=True, WithWindow=False)
                doc.ExportAsFixedFormat(OutputFileName=caminho_completo_saida, ExportFormat=WD_FORMAT_PDF)
                return True

            # --- EXCEL ---
            elif extensao in ['xlsx', 'xls']:
                excel = win32com.client.DispatchEx("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                wb = excel.Workbooks.Open(caminho_completo_entrada, ReadOnly=True)
                wb.ExportAsFixedFormat(XL_FORMAT_PDF, caminho_completo_saida)
                return True

            # --- POWERPOINT ---
            elif extensao in ['pptx', 'ppt']:
                ppt = win32com.client.DispatchEx("PowerPoint.Application")
                # PowerPoint precisa de flags específicas para rodar oculto em background
                presentation = ppt.Presentations.Open(
                    caminho_completo_entrada, 
                    ReadOnly=1,     # msoTrue
                    Untitled=0,     # msoFalse
                    WithWindow=0    # msoFalse
                )
                presentation.SaveAs(caminho_completo_saida, PP_FORMAT_PDF)
                return True

            else:
                self.log(f"[AVISO] Extensão não suportada: {caminho_completo_entrada}")
                return False

        except Exception as e:
            self.log(f"[ERRO] Falha ao converter {caminho_completo_entrada}: {e}")
            return False

        finally:
            # Limpeza cuidadosa dos processos COM para não deixar instâncias zumbis
            if doc: 
                try: doc.Close(False)
                except: pass
            if word: 
                try: word.Quit()
                except: pass
            if wb:
                try: wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
            if presentation:
                try: presentation.Close()
                except: pass
            if ppt:
                try: ppt.Quit()
                except: pass
            pythoncom.CoUninitialize()

    def unir_pdfs(self, lista_pdfs):
        merger = PdfWriter()
        self.caminho_resultado = os.path.join(self.diretorio_raiz, self.nome_arquivo_saida)

        self.log(f"\nIniciando união de {len(lista_pdfs)} PDFs...")
        if not lista_pdfs:
            self.log("Nenhum arquivo PDF encontrado para unir.")
            return

        try:
            for pdf in sorted(lista_pdfs):
                try:
                    merger.append(pdf)
                except Exception as e:
                    self.log(f"[AVISO] Erro ao adicionar '{pdf}': {e}")

            with open(self.caminho_resultado, "wb") as saida:
                merger.write(saida)
            self.log(f"\n[SUCESSO] PDFs unidos em: {self.caminho_resultado}")
        except Exception as e:
            self.log(f"[ERRO] Falha ao unir PDFs: {e}")
            self.caminho_resultado = None
        finally:
            merger.close()

    def executar(self):
        self.log(f"--- INÍCIO DA CONSOLIDAÇÃO ---")
        self.log(f"Diretório raiz: {self.diretorio_raiz}")
        self.status("Procurando arquivos...")

        arquivos_por_pasta = self.listar_arquivos_em_dicionario(self.diretorio_raiz)
        arquivos_relevantes = []
        
        for pasta, arquivos in arquivos_por_pasta.items():
            for nome_arquivo in arquivos:
                nome_lower = nome_arquivo.lower()
                if nome_lower.endswith('.pdf') and not nome_lower.startswith(self.nome_arquivo_saida.lower().replace('.pdf', '')):
                    arquivos_relevantes.append((pasta, nome_arquivo))
                elif nome_lower.endswith(('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')):
                    arquivos_relevantes.append((pasta, nome_arquivo))

        total = len(arquivos_relevantes)
        self.progresso(0, max(total, 1))
        lista_de_pdfs_para_unir = []

        for indice, (pasta, nome_arquivo) in enumerate(arquivos_relevantes, start=1):
            if self.cancel_event.is_set():
                self.log("\n[CANCELADO] Processo interrompido.")
                self.status("Cancelado.")
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
                self.log(f"  [PULAR] Destino já existe. Adicionando à união: {nome_saida}")
                lista_de_pdfs_para_unir.append(caminho_saida)
            else:
                if self.converter_para_pdf(caminho_entrada, caminho_saida):
                    self.log(f"  [SUCESSO] Convertido: {nome_arquivo}")
                    lista_de_pdfs_para_unir.append(caminho_saida)
                    self.total_convertidos += 1
                else:
                    self.total_erros += 1

            self.progresso(indice, total)

        self.log(f"\n--- Processo de Conversão Finalizado ---")
        self.log(f"Convertidos: {self.total_convertidos} | Falhas: {self.total_erros}")
        self.status("Unindo PDFs...")
        self.unir_pdfs(lista_de_pdfs_para_unir)
        self.status("Concluído.")


class ProcessadorDivisao(ProcessadorBase):
    def __init__(self, pdf_entrada, arquivo_saida, string_paginas, log_queue, cancel_event):
        super().__init__(log_queue, cancel_event)
        self.pdf_entrada = pdf_entrada
        self.arquivo_saida = arquivo_saida
        self.string_paginas = string_paginas

    def analisar_paginas(self, string_paginas, total_paginas):
        paginas_para_extrair = []
        partes = string_paginas.split(',')
        for parte in partes:
            parte = parte.strip()
            if not parte: continue
            if '-' in parte:
                subpartes = parte.split('-')
                if len(subpartes) == 2:
                    try:
                        inicio = int(subpartes[0].strip())
                        fim = int(subpartes[1].strip())
                        # Converte de base-1 (humana) para base-0 (pypdf)
                        for p in range(inicio, fim + 1):
                            if 1 <= p <= total_paginas:
                                paginas_para_extrair.append(p - 1)
                    except ValueError:
                        pass
            else:
                try:
                    p = int(parte)
                    if 1 <= p <= total_paginas:
                        paginas_para_extrair.append(p - 1)
                except ValueError:
                    pass
        
        # Remover duplicatas mantendo a ordem escolhida pelo usuário
        vistos = set()
        resultado = []
        for p in paginas_para_extrair:
            if p not in vistos:
                vistos.add(p)
                resultado.append(p)
        return resultado

    def executar(self):
        self.log("--- INÍCIO DA EXTRAÇÃO DE PÁGINAS ---")
        self.log(f"Arquivo origem: {self.pdf_entrada}")
        self.log(f"Intervalo solicitado: {self.string_paginas}")
        self.status("Lendo arquivo PDF...")

        try:
            reader = PdfReader(self.pdf_entrada)
            total_paginas = len(reader.pages)
            self.log(f"O arquivo tem {total_paginas} página(s).")
            
            indices_paginas = self.analisar_paginas(self.string_paginas, total_paginas)
            
            if not indices_paginas:
                self.log("[ERRO] Nenhuma página válida foi selecionada.")
                self.status("Erro: Páginas inválidas.")
                return

            self.log(f"Total de páginas a extrair: {len(indices_paginas)}")
            self.progresso(0, len(indices_paginas))

            writer = PdfWriter()
            
            for idx, num_pagina in enumerate(indices_paginas):
                if self.cancel_event.is_set():
                    self.log("[CANCELADO] Extração interrompida.")
                    self.status("Cancelado.")
                    return

                writer.add_page(reader.pages[num_pagina])
                self.status(f"Extraindo página {num_pagina + 1}...")
                self.progresso(idx + 1, len(indices_paginas))
            
            self.status("Salvando novo arquivo...")
            
            # Tenta limpar o metadado de título do PDF original para não confundir o leitor de PDF
            try:
                writer.add_metadata({"/Title": os.path.basename(self.arquivo_saida)})
            except Exception:
                pass
                
            with open(self.arquivo_saida, "wb") as f_out:
                writer.write(f_out)
            
            self.caminho_resultado = self.arquivo_saida
            self.log(f"Arquivo gerado com sucesso: {self.arquivo_saida}")
            self.log("--- EXTRAÇÃO CONCLUÍDA ---")
            self.status("Concluído.")

        except Exception as e:
            self.log(f"[ERRO] Falha ao extrair PDF: {e}")
            self.status("Erro na extração.")


class ProcessadorCompressao(ProcessadorBase):
    def __init__(self, pdf_entrada, arquivo_saida, log_queue, cancel_event):
        super().__init__(log_queue, cancel_event)
        self.pdf_entrada = pdf_entrada
        self.arquivo_saida = arquivo_saida

    def executar(self):
        self.log("--- INÍCIO DA COMPRESSÃO DE PDF ---")
        self.log(f"Origem: {self.pdf_entrada}")
        self.status("Lendo arquivo...")

        try:
            reader = PdfReader(self.pdf_entrada)
            writer = PdfWriter()
            total_paginas = len(reader.pages)
            self.progresso(0, total_paginas)

            for i, page in enumerate(reader.pages):
                if self.cancel_event.is_set():
                    self.log("[CANCELADO] Compressão interrompida.")
                    self.status("Cancelado.")
                    return
                writer.add_page(page)
                self.progresso(i + 1, total_paginas)

            self.status("Comprimindo fluxos de dados...")
            self.log("Aplicando compressão nos objetos internos do PDF...")
            for page in writer.pages:
                page.compress_content_streams()

            self.status("Salvando arquivo...")
            with open(self.arquivo_saida, "wb") as f_out:
                writer.write(f_out)

            # Comparar tamanhos
            tamanho_original = os.path.getsize(self.pdf_entrada) / (1024 * 1024)
            tamanho_novo = os.path.getsize(self.arquivo_saida) / (1024 * 1024)
            
            self.log(f"Tamanho Original: {tamanho_original:.2f} MB")
            self.log(f"Tamanho Comprimido: {tamanho_novo:.2f} MB")
            
            self.caminho_resultado = self.arquivo_saida
            self.log("--- COMPRESSÃO CONCLUÍDA ---")
            self.status("Concluído.")

        except Exception as e:
            self.log(f"[ERRO] Falha ao comprimir PDF: {e}")
            self.status("Erro na compressão.")


# ============================================================
#                        INTERFACE GRÁFICA
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDFs")
        self.geometry("780x620")
        self.minsize(700, 500)
        self._aplicar_icone()

        self.log_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.caminho_resultado_atual = None

        self._montar_widgets()
        self._checar_dependencias()
        self.after(100, self._processar_fila)

    def _aplicar_icone(self):
        try:
            caminho_icone = resource_path(os.path.join("src", "icon.ico"))
            if not os.path.exists(caminho_icone):
                caminho_icone = resource_path("icon.ico")
            if os.path.exists(caminho_icone):
                self.iconbitmap(caminho_icone)
        except Exception:
            pass

    def _montar_widgets(self):
        padding = {"padx": 10, "pady": 6}

        # --- ABAS ---
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="x", padx=10, pady=10)

        # Abas
        self.tab_consolidar = ttk.Frame(self.notebook)
        self.tab_dividir = ttk.Frame(self.notebook)
        self.tab_comprimir = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_consolidar, text="Converter e Juntar PDF")
        self.notebook.add(self.tab_dividir, text="Dividir PDF")
        self.notebook.add(self.tab_comprimir, text="Comprimir PDF")

        self._montar_aba_consolidar()
        self._montar_aba_dividir()
        self._montar_aba_comprimir()

        # --- ÁREA COMUM (Ações, Progresso e Log) ---
        frame_acoes = ttk.Frame(self)
        frame_acoes.pack(fill="x", **padding)

        self.btn_cancelar = ttk.Button(frame_acoes, text="Cancelar Operação", command=self._cancelar_processamento, state="disabled")
        self.btn_cancelar.pack(side="left")

        self.btn_abrir_resultado = ttk.Button(frame_acoes, text="Abrir Resultado", command=self._abrir_resultado, state="disabled")
        self.btn_abrir_resultado.pack(side="left", padx=(8, 0))

        frame_progresso = ttk.Frame(self)
        frame_progresso.pack(fill="x", **padding)

        self.var_status = tk.StringVar(value="Aguardando ação...")
        ttk.Label(frame_progresso, textvariable=self.var_status).pack(anchor="w")

        self.progressbar = ttk.Progressbar(frame_progresso, mode="determinate")
        self.progressbar.pack(fill="x", pady=(4, 0))

        frame_log = ttk.LabelFrame(self, text="Log de Execução")
        frame_log.pack(fill="both", expand=True, **padding)

        self.text_log = scrolledtext.ScrolledText(frame_log, state="disabled", wrap="word", height=10)
        self.text_log.pack(fill="both", expand=True, padx=8, pady=8)

    def _montar_aba_consolidar(self):
        pad = {"padx": 10, "pady": 5}
        
        # Pasta de origem
        frm_pasta = ttk.Frame(self.tab_consolidar)
        frm_pasta.pack(fill="x", **pad)
        ttk.Label(frm_pasta, text="Pasta com arquivos (.docx/.xls/.ppt/.pdf):").pack(anchor="w")
        
        box_pasta = ttk.Frame(frm_pasta)
        box_pasta.pack(fill="x", pady=(2, 0))
        self.var_pasta_cons = tk.StringVar()
        ttk.Entry(box_pasta, textvariable=self.var_pasta_cons, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(box_pasta, text="Procurar...", command=lambda: self._selecionar_pasta(self.var_pasta_cons)).pack(side="left", padx=(5, 0))

        # Nome de Saída
        frm_opc = ttk.Frame(self.tab_consolidar)
        frm_opc.pack(fill="x", **pad)
        ttk.Label(frm_opc, text="Nome do PDF Final:").pack(side="left")
        self.var_nome_saida = tk.StringVar(value=NOME_SAIDA_PADRAO)
        ttk.Entry(frm_opc, textvariable=self.var_nome_saida, width=35).pack(side="left", padx=(5, 10))

        # Checkbox Reprocessar
        self.var_reprocessar = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_opc, text="Forçar reconversão (se o PDF já existir)", variable=self.var_reprocessar).pack(side="left")

        # Iniciar
        ttk.Button(self.tab_consolidar, text="▶ Iniciar Consolidação", command=self._iniciar_consolidacao).pack(pady=10)

    def _montar_aba_dividir(self):
        pad = {"padx": 10, "pady": 5}
        
        frm_arq = ttk.Frame(self.tab_dividir)
        frm_arq.pack(fill="x", **pad)
        ttk.Label(frm_arq, text="Arquivo PDF original:").pack(anchor="w")
        box_arq = ttk.Frame(frm_arq)
        box_arq.pack(fill="x")
        self.var_arq_div = tk.StringVar()
        ttk.Entry(box_arq, textvariable=self.var_arq_div, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(box_arq, text="Procurar...", command=lambda: self._selecionar_arquivo(self.var_arq_div)).pack(side="left", padx=(5, 0))

        frm_opc = ttk.Frame(self.tab_dividir)
        frm_opc.pack(fill="x", **pad)
        ttk.Label(frm_opc, text="Páginas a extrair (ex: 1-5, 8, 11-13):").pack(anchor="w")
        self.var_paginas_str = tk.StringVar(value="")
        ttk.Entry(frm_opc, textvariable=self.var_paginas_str, width=40).pack(side="left", pady=(2, 0))

        frm_arq_out = ttk.Frame(self.tab_dividir)
        frm_arq_out.pack(fill="x", **pad)
        ttk.Label(frm_arq_out, text="Salvar como (Novo PDF com as páginas):").pack(anchor="w")
        box_arq_out = ttk.Frame(frm_arq_out)
        box_arq_out.pack(fill="x")
        self.var_arq_div_out = tk.StringVar()
        ttk.Entry(box_arq_out, textvariable=self.var_arq_div_out, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(box_arq_out, text="Salvar em...", command=lambda: self._selecionar_salvar_como(self.var_arq_div_out)).pack(side="left", padx=(5, 0))

        ttk.Button(self.tab_dividir, text="▶ Iniciar Extração", command=self._iniciar_divisao).pack(pady=10)

    def _montar_aba_comprimir(self):
        pad = {"padx": 10, "pady": 5}
        
        frm_arq = ttk.Frame(self.tab_comprimir)
        frm_arq.pack(fill="x", **pad)
        ttk.Label(frm_arq, text="Arquivo PDF para comprimir:").pack(anchor="w")
        box_arq = ttk.Frame(frm_arq)
        box_arq.pack(fill="x")
        self.var_arq_comp = tk.StringVar()
        ttk.Entry(box_arq, textvariable=self.var_arq_comp, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(box_arq, text="Procurar...", command=lambda: self._selecionar_arquivo(self.var_arq_comp)).pack(side="left", padx=(5, 0))

        frm_arq_out = ttk.Frame(self.tab_comprimir)
        frm_arq_out.pack(fill="x", **pad)
        ttk.Label(frm_arq_out, text="Salvar como (PDF Comprimido):").pack(anchor="w")
        box_arq_out = ttk.Frame(frm_arq_out)
        box_arq_out.pack(fill="x")
        self.var_arq_comp_out = tk.StringVar()
        ttk.Entry(box_arq_out, textvariable=self.var_arq_comp_out, state="readonly").pack(side="left", fill="x", expand=True)
        ttk.Button(box_arq_out, text="Salvar em...", command=self._selecionar_salvar_como).pack(side="left", padx=(5, 0))

        ttk.Button(self.tab_comprimir, text="▶ Iniciar Compressão", command=self._iniciar_compressao).pack(pady=10)

    def _checar_dependencias(self):
        faltando = []
        if not COM_DISPONIVEL: faltando.append("pywin32 (necessário para conversões)")
        if not PYPDF_DISPONIVEL: faltando.append("pypdf (necessário para manipular PDFs)")
        if faltando:
            messagebox.showwarning("Aviso", "Dependências ausentes:\n- " + "\n- ".join(faltando))

    def _selecionar_pasta(self, variavel_string):
        pasta = filedialog.askdirectory(title="Selecione a pasta")
        if pasta:
            variavel_string.set(os.path.normpath(pasta))

    def _selecionar_arquivo(self, variavel_string):
        arquivo = filedialog.askopenfilename(title="Selecione um PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if arquivo:
            variavel_string.set(os.path.normpath(arquivo))

    def _selecionar_salvar_como(self, variavel_string=None):
        arquivo = filedialog.asksaveasfilename(
            title="Salvar PDF como...",
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if arquivo:
            # Se chamado por um botão sem usar lambda, o variavel_string pode vir como um objeto de evento
            if variavel_string is None or not isinstance(variavel_string, tk.StringVar):
                self.var_arq_comp_out.set(os.path.normpath(arquivo))
            else:
                variavel_string.set(os.path.normpath(arquivo))

    def _preparar_ui_para_execucao(self):
        if not COM_DISPONIVEL or not PYPDF_DISPONIVEL:
            messagebox.showerror("Erro", "As dependências 'pywin32' e 'pypdf' são obrigatórias.")
            return False

        self.text_log.configure(state="normal")
        self.text_log.delete("1.0", tk.END)
        self.text_log.configure(state="disabled")
        
        self.progressbar["value"] = 0
        self.caminho_resultado_atual = None
        self.btn_abrir_resultado.configure(state="disabled")
        self.cancel_event.clear()

        # Desabilita botões temporariamente via abas iterando pelos filhos
        self.btn_cancelar.configure(state="normal")
        return True

    def _iniciar_consolidacao(self):
        pasta = self.var_pasta_cons.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione uma pasta válida.")
            return
            
        if self._preparar_ui_para_execucao():
            nome_saida = self.var_nome_saida.get().strip() or NOME_SAIDA_PADRAO
            if not nome_saida.lower().endswith(".pdf"): nome_saida += ".pdf"

            processador = ProcessadorConsolidacao(pasta, nome_saida, self.var_reprocessar.get(), self.log_queue, self.cancel_event)
            self._disparar_thread(processador)

    def _iniciar_divisao(self):
        arquivo = self.var_arq_div.get().strip()
        arquivo_saida = self.var_arq_div_out.get().strip()
        string_paginas = self.var_paginas_str.get().strip()
        
        if not arquivo or not os.path.isfile(arquivo):
            messagebox.showerror("Erro", "Selecione um arquivo PDF válido para extrair as páginas.")
            return
        if not string_paginas:
            messagebox.showerror("Erro", "Defina o intervalo de páginas (ex: 1-5, 8).")
            return
        if not arquivo_saida:
            messagebox.showerror("Erro", "Selecione o local de salvamento do novo arquivo PDF.")
            return
        if os.path.isdir(arquivo_saida):
            messagebox.showerror("Atenção", "O local de salvamento não pode ser apenas uma pasta.\n\nClique em 'Salvar em...' e digite o NOME do arquivo desejado (ex: paginas_extraidas.pdf).")
            return

        if self._preparar_ui_para_execucao():
            processador = ProcessadorDivisao(arquivo, arquivo_saida, string_paginas, self.log_queue, self.cancel_event)
            self._disparar_thread(processador)

    def _iniciar_compressao(self):
        arquivo = self.var_arq_comp.get().strip()
        saida = self.var_arq_comp_out.get().strip()
        
        if not arquivo or not os.path.isfile(arquivo):
            messagebox.showerror("Erro", "Selecione um arquivo PDF válido para comprimir.")
            return
        if not saida:
            messagebox.showerror("Erro", "Selecione o local de salvamento do novo arquivo.")
            return
        if os.path.isdir(saida):
            messagebox.showerror("Atenção", "O local de salvamento não pode ser apenas uma pasta.\n\nClique em 'Salvar em...' e digite o NOME do arquivo desejado (ex: arquivo_comprimido.pdf).")
            return

        if self._preparar_ui_para_execucao():
            processador = ProcessadorCompressao(arquivo, saida, self.log_queue, self.cancel_event)
            self._disparar_thread(processador)

    def _disparar_thread(self, processador):
        self.worker_thread = threading.Thread(target=self._executar_com_seguranca, args=(processador,), daemon=True)
        self.worker_thread.start()

    def _executar_com_seguranca(self, processador):
        try:
            processador.executar()
        except Exception as e:
            self.log_queue.put(("log", f"[ERRO INESPERADO] {e}"))
            self.log_queue.put(("status", "Erro."))
        finally:
            self.log_queue.put(("done", processador.caminho_resultado))

    def _cancelar_processamento(self):
        self.cancel_event.set()
        self.btn_cancelar.configure(state="disabled")
        self.var_status.set("Cancelando... aguarde a conclusão da etapa atual.")

    def _abrir_resultado(self):
        caminho = self.caminho_resultado_atual
        if caminho and os.path.exists(caminho):
            try:
                caminho = os.path.normpath(caminho)
                if sys.platform.startswith("win"):
                    # Abre a pasta no Windows Explorer com o arquivo recém-gerado já selecionado
                    subprocess.run(f'explorer /select,"{caminho}"')
                elif sys.platform == "darwin":
                    subprocess.run(["open", "-R", caminho], check=False)
                else:
                    pasta = os.path.dirname(caminho)
                    subprocess.run(["xdg-open", pasta], check=False)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o item: {e}")

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
                    self.caminho_resultado_atual = valor
                    self.btn_cancelar.configure(state="disabled")
                    
                    if valor and os.path.exists(valor):
                        self.btn_abrir_resultado.configure(state="normal")
                        messagebox.showinfo("Concluído", "Operação finalizada com sucesso!")
                    else:
                        messagebox.showwarning("Aviso", "O processo terminou, mas o arquivo final pode não ter sido gerado. Verifique o log.")
        except queue.Empty:
            pass
        finally:
            self.after(100, self._processar_fila)


if __name__ == "__main__":
    app = App()
    app.mainloop()
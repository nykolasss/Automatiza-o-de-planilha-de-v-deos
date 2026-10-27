"""
Interface gráfica para o videos_para_sheets.py.

Como usar: dê dois cliques neste arquivo (ou rode `python interface_videos.pyw`).
Mantenha este arquivo na MESMA pasta do videos_para_sheets.py e do credenciais.json.

Fluxo:
    1. Verificar vídeos novos  -> mostra a prévia do que seria gravado (nada é gravado ainda)
    2. Gravar na planilha      -> grava de verdade, depois de você confirmar
"""

import json
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

try:
    import videos_para_sheets as core
except ImportError as erro:
    _raiz = tk.Tk()
    _raiz.withdraw()
    messagebox.showerror(
        "Falta instalar algo",
        "Não consegui abrir o videos_para_sheets.py.\n\n"
        "1) Confirme que ele está na mesma pasta deste arquivo.\n"
        "2) Abra o terminal e rode:  pip install gspread\n\n"
        f"Detalhe: {erro}",
    )
    raise SystemExit

ARQUIVO_CONFIG = Path(__file__).with_name("config_interface.json")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Vídeos para o Google Sheets")
        self.minsize(820, 640)

        self.fila = queue.Queue()   # mensagens vindas da thread de trabalho
        self.plano = None           # resultado da última verificação
        self.ocupado = False

        cfg = self.carregar_config()
        self.v_pasta = tk.StringVar(value=cfg["pasta"])
        self.v_cred = tk.StringVar(value=cfg["credenciais"])
        self.v_planilha = tk.StringVar(value=cfg["planilha"])
        self.v_aba = tk.StringVar(value=cfg["aba"])
        self.v_desde = tk.StringVar(value=cfg["desde"])
        self.v_status = tk.StringVar(value="Preencha os campos e clique em “Verificar vídeos novos”.")

        self.montar()
        self.after(100, self.processar_fila)

    # ---------- configurações salvas ----------
    def carregar_config(self) -> dict:
        cfg = {
            "pasta": core.PASTA_VIDEOS,
            "credenciais": core.ARQUIVO_CREDENCIAIS,
            "planilha": core.NOME_PLANILHA,
            "aba": core.NOME_ABA,
            "desde": core.DATA_INICIAL,
        }
        try:
            cfg.update(json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            pass
        return cfg

    def salvar_config(self):
        cfg = {
            "pasta": self.v_pasta.get().strip(),
            "credenciais": self.v_cred.get().strip(),
            "planilha": self.v_planilha.get().strip(),
            "aba": self.v_aba.get().strip(),
            "desde": self.v_desde.get().strip(),
        }
        try:
            ARQUIVO_CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass

    # ---------- montagem da janela ----------
    def montar(self):
        corpo = ttk.Frame(self, padding=14)
        corpo.pack(fill="both", expand=True)

        # Configurações
        cfg = ttk.LabelFrame(corpo, text=" Configurações ", padding=10)
        cfg.pack(fill="x")
        cfg.columnconfigure(1, weight=1)

        def linha(n, rotulo, var, botao=None, largura=None, dica=None):
            ttk.Label(cfg, text=rotulo).grid(row=n, column=0, sticky="w", padx=(0, 10), pady=4)
            entrada = ttk.Entry(cfg, textvariable=var, width=largura or 50)
            entrada.grid(row=n, column=1, sticky="ew" if not largura else "w", pady=4)
            if botao:
                ttk.Button(cfg, text="Escolher…", command=botao).grid(row=n, column=2, padx=(8, 0))
            if dica:
                ttk.Label(cfg, text=dica, foreground="#666").grid(row=n, column=2, sticky="w", padx=(8, 0))

        linha(0, "Pasta dos vídeos", self.v_pasta, self.escolher_pasta)
        linha(1, "Credenciais (.json)", self.v_cred, self.escolher_credenciais)
        linha(2, "Planilha", self.v_planilha)
        linha(3, "Aba do pacote", self.v_aba)
        linha(4, "A partir de", self.v_desde, largura=14,
              dica="dd/mm/aaaa  (vazio = todos os vídeos)")

        # Botões
        barra = ttk.Frame(corpo)
        barra.pack(fill="x", pady=(12, 6))
        self.btn_verificar = ttk.Button(barra, text="1. Verificar vídeos novos", command=self.verificar)
        self.btn_verificar.pack(side="left")
        self.btn_gravar = ttk.Button(barra, text="2. Gravar na planilha", command=self.gravar,
                                     state="disabled")
        self.btn_gravar.pack(side="left", padx=8)
        self.progresso = ttk.Progressbar(barra, mode="indeterminate", length=140)
        self.progresso.pack(side="right")

        ttk.Label(corpo, textvariable=self.v_status, wraplength=780).pack(fill="x", pady=(0, 6))

        # Prévia
        previa = ttk.LabelFrame(corpo, text=" Prévia: o que será gravado ", padding=6)
        previa.pack(fill="both", expand=True)
        colunas = ("linha", "nome", "data", "minutagem")
        self.tabela = ttk.Treeview(previa, columns=colunas, show="headings", height=9)
        for col, titulo, largura, ancora in (
            ("linha", "Linha", 60, "center"),
            ("nome", "Nome do vídeo", 420, "w"),
            ("data", "Data", 100, "center"),
            ("minutagem", "Minutagem", 100, "center"),
        ):
            self.tabela.heading(col, text=titulo)
            self.tabela.column(col, width=largura, anchor=ancora, stretch=(col == "nome"))
        rolagem = ttk.Scrollbar(previa, orient="vertical", command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=rolagem.set)
        self.tabela.pack(side="left", fill="both", expand=True)
        rolagem.pack(side="right", fill="y")

        # Registro
        registro = ttk.LabelFrame(corpo, text=" Registro ", padding=6)
        registro.pack(fill="both", pady=(10, 0))
        self.txt = ScrolledText(registro, height=7, state="disabled", font=("Consolas", 9))
        self.txt.pack(fill="both", expand=True)

    # ---------- escolhas de arquivo/pasta ----------
    def escolher_pasta(self):
        pasta = filedialog.askdirectory(initialdir=self.v_pasta.get() or None,
                                        title="Escolha a pasta central dos vídeos")
        if pasta:
            self.v_pasta.set(str(Path(pasta)))

    def escolher_credenciais(self):
        arquivo = filedialog.askopenfilename(title="Escolha o credenciais.json",
                                             filetypes=[("JSON", "*.json"), ("Todos", "*.*")])
        if arquivo:
            self.v_cred.set(str(Path(arquivo)))

    # ---------- utilidades da tela ----------
    def log(self, texto: str):
        self.txt.configure(state="normal")
        self.txt.insert("end", texto + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def definir_ocupado(self, ocupado: bool):
        self.ocupado = ocupado
        self.btn_verificar.configure(state="disabled" if ocupado else "normal")
        pode_gravar = (not ocupado) and self.plano is not None and bool(self.plano.itens)
        self.btn_gravar.configure(state="normal" if pode_gravar else "disabled")
        if ocupado:
            self.progresso.start(12)
        else:
            self.progresso.stop()

    def limpar_previa(self):
        self.tabela.delete(*self.tabela.get_children())

    # ---------- ações ----------
    def verificar(self):
        if self.ocupado:
            return
        self.salvar_config()
        self.plano = None
        self.limpar_previa()
        self.definir_ocupado(True)
        self.v_status.set("Verificando… (ler a duração dos vídeos pode levar alguns instantes)")
        self.log("— Nova verificação —")
        argumentos = (
            self.v_pasta.get().strip(),
            self.v_cred.get().strip(),
            self.v_planilha.get().strip(),
            self.v_aba.get().strip(),
            self.v_desde.get().strip() or None,
        )
        threading.Thread(target=self._trabalho_verificar, args=argumentos, daemon=True).start()

    def _trabalho_verificar(self, *argumentos):
        try:
            plano = core.preparar(*argumentos, log=lambda m: self.fila.put(("log", m)))
            self.fila.put(("plano", plano))
        except core.ErroApp as e:
            self.fila.put(("erro", str(e)))
        except Exception as e:  # qualquer outra falha também precisa aparecer na tela
            self.fila.put(("erro", f"Erro inesperado: {e}"))

    def gravar(self):
        if self.ocupado or not self.plano or not self.plano.itens:
            return
        n = len(self.plano.itens)
        if not messagebox.askyesno(
            "Confirmar gravação",
            f"Gravar {n} vídeo(s) na aba “{self.plano.nome_aba}”?",
        ):
            return
        self.definir_ocupado(True)
        self.v_status.set("Gravando na planilha…")
        threading.Thread(target=self._trabalho_gravar, args=(self.plano,), daemon=True).start()

    def _trabalho_gravar(self, plano):
        try:
            n = core.gravar_plano(plano)
            self.fila.put(("gravado", n))
        except core.ErroApp as e:
            self.fila.put(("erro", str(e)))
        except Exception as e:
            self.fila.put(("erro", f"Erro inesperado: {e}"))

    # ---------- recebe mensagens das threads ----------
    def processar_fila(self):
        try:
            while True:
                tipo, dado = self.fila.get_nowait()
                if tipo == "log":
                    self.log(dado)
                elif tipo == "plano":
                    self.mostrar_plano(dado)
                elif tipo == "gravado":
                    self.plano = None
                    self.limpar_previa()
                    self.definir_ocupado(False)
                    self.v_status.set(f"Pronto! {dado} vídeo(s) gravado(s) na planilha.")
                    self.log(f"{dado} vídeo(s) gravado(s).")
                    messagebox.showinfo("Gravado", f"{dado} vídeo(s) gravado(s) na planilha.")
                elif tipo == "erro":
                    self.definir_ocupado(False)
                    self.v_status.set("Algo deu errado. Veja a mensagem e tente de novo.")
                    self.log(f"ERRO: {dado}")
                    messagebox.showerror("Algo deu errado", dado)
        except queue.Empty:
            pass
        self.after(100, self.processar_fila)

    def mostrar_plano(self, plano):
        self.plano = plano
        for item in plano.itens:
            self.tabela.insert("", "end", values=(
                item["linha"], item["nome"], f"{item['data']:%d/%m/%Y}", item["minutagem"],
            ))
        texto = core.resumo(plano)
        if plano.itens:
            texto += "  →  confira a prévia e clique em “Gravar na planilha”."
        else:
            texto += "  →  nada novo para gravar."
        if plano.sem_linha:
            texto += "\n" + core.AVISO_SEM_LINHA
        self.v_status.set(texto)
        self.log(core.resumo(plano))
        self.definir_ocupado(False)


if __name__ == "__main__":
    try:  # texto mais nítido no Windows
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()

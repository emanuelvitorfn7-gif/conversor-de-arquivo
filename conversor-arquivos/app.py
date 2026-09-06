from pathlib import Path
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd


# ----------------------------------------------------------------------
# Dados / tema
# ----------------------------------------------------------------------

FORMATOS = {
    "excel": {"label": "Excel", "ext": ".xlsx", "desc": "Planilhas", "icon": "📊"},
    "csv": {"label": "CSV", "ext": ".csv", "desc": "Universal", "icon": "📄"},
    "json": {"label": "JSON", "ext": ".json", "desc": "APIs & devs", "icon": "🧩"},
    "tsv": {"label": "TSV", "ext": ".tsv", "desc": "Tabulado", "icon": "📑"},
}

CORES = {
    "bg": "#0A0E17",
    "panel": "#0F1524",
    "card": "#131B2E",
    "card_hover": "#182136",
    "drop": "#111A2E",
    "border": "#232F47",
    "border_light": "#334155",
    "accent": "#38BDF8",
    "accent_hover": "#7DD3FC",
    "accent_dark": "#0C4A6E",
    "text": "#F1F5F9",
    "muted": "#8B98AF",
    "muted_dark": "#5B6B84",
    "success": "#22C55E",
    "success_bg": "#052E1B",
    "error": "#F87171",
    "warning": "#FBBF24",
}

FONT_TITULO = ("Segoe UI", 22, "bold")
FONT_SUB = ("Segoe UI", 10)
FONT_STEP = ("Segoe UI", 8, "bold")
FONT_BODY = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)


def carregar_arquivo(caminho: str) -> pd.DataFrame:
    """Carrega um arquivo suportado e devolve seus dados."""
    extensao = Path(caminho).suffix.lower()

    if extensao == ".xlsx":
        return pd.read_excel(caminho, engine="openpyxl")
    if extensao == ".csv":
        return pd.read_csv(caminho, sep=None, engine="python", encoding="utf-8-sig")
    if extensao == ".json":
        return pd.read_json(caminho)
    if extensao == ".tsv":
        return pd.read_csv(caminho, sep="\t", encoding="utf-8-sig")

    raise ValueError("Formato não suportado. Use XLSX, CSV, JSON ou TSV.")


def salvar_arquivo(dados: pd.DataFrame, caminho: str) -> None:
    """Salva os dados conforme a extensão escolhida."""
    extensao = Path(caminho).suffix.lower()

    if extensao == ".xlsx":
        dados.to_excel(caminho, index=False, engine="openpyxl")
    elif extensao == ".csv":
        dados.to_csv(caminho, index=False, encoding="utf-8-sig")
    elif extensao == ".json":
        dados.to_json(caminho, orient="records", indent=4, force_ascii=False)
    elif extensao == ".tsv":
        dados.to_csv(caminho, sep="\t", index=False, encoding="utf-8-sig")
    else:
        raise ValueError("Formato de saída não suportado.")


def tamanho_legivel(n_bytes: int) -> str:
    if n_bytes < 1024:
        return f"{n_bytes} B"
    if n_bytes < 1024**2:
        return f"{n_bytes / 1024:.1f} KB"
    return f"{n_bytes / 1024**2:.1f} MB"


class ConversorApp:
    def __init__(self, janela: tk.Tk) -> None:
        self.janela = janela
        self.arquivo = ""
        self.df: pd.DataFrame | None = None
        self.formato_key = tk.StringVar(value="csv")
        self.status_texto = tk.StringVar(value="Pronto para converter")
        self.status_tipo = tk.StringVar(value="idle")  # idle | ok | erro | busy
        self.convertendo = False

        self._configurar_janela()
        self._configurar_estilos()
        self._criar_interface()
        self._atualizar_formato_visual()

    # -- janela / estilo -------------------------------------------------
    def _configurar_janela(self) -> None:
        self.janela.title("Conversor de Arquivos — rápido, local e elegante")
        self.janela.configure(bg=CORES["bg"])
        self.janela.minsize(920, 660)

        largura, altura = 1020, 720
        self.janela.geometry(f"{largura}x{altura}")
        self.janela.update_idletasks()
        x = (self.janela.winfo_screenwidth() - largura) // 2
        y = (self.janela.winfo_screenheight() - altura) // 2
        self.janela.geometry(f"{largura}x{altura}+{x}+{y}")

        try:
            self.janela.iconbitmap(default="")
        except Exception:
            pass

    def _configurar_estilos(self) -> None:
        estilo = ttk.Style()
        try:
            estilo.theme_use("clam")
        except Exception:
            pass

        estilo.configure(
            "Dark.Treeview",
            background=CORES["card"],
            fieldbackground=CORES["card"],
            foreground=CORES["text"],
            rowheight=28,
            borderwidth=0,
            font=("Segoe UI", 9),
        )
        estilo.configure(
            "Dark.Treeview.Heading",
            background="#1A2438",
            foreground="#9FB0C8",
            relief="flat",
            font=("Segoe UI", 8, "bold"),
            padding=(10, 8),
        )
        estilo.map("Dark.Treeview", background=[("selected", "#1D4ED8")])
        estilo.configure(
            "Dark.Vertical.TScrollbar",
            background=CORES["card"],
            troughcolor=CORES["panel"],
            borderwidth=0,
            arrowsize=0,
        )
        estilo.configure(
            "Dark.Horizontal.TScrollbar",
            background=CORES["card"],
            troughcolor=CORES["panel"],
            borderwidth=0,
            arrowsize=0,
        )
        estilo.configure(
            "Convert.Horizontal.TProgressbar",
            background=CORES["accent"],
            troughcolor="#1A2438",
            borderwidth=0,
            thickness=6,
        )

    # -- construção ------------------------------------------------------
    def _criar_interface(self) -> None:
        # Header
        header = tk.Frame(self.janela, bg=CORES["bg"])
        header.pack(fill="x", padx=28, pady=(22, 6))

        badge = tk.Label(
            header,
            text="⇄  CONVERSOR",
            bg="#0E2A3F",
            fg=CORES["accent"],
            font=("Segoe UI", 8, "bold"),
            padx=12,
            pady=6,
        )
        badge.pack(side="left", anchor="n", pady=(4, 0))

        pill_frame = tk.Frame(header, bg=CORES["bg"])
        pill_frame.pack(side="right", anchor="n", pady=(4, 0))
        for txt in ("100% local", "4 formatos"):
            tk.Label(
                pill_frame,
                text=txt,
                bg="#141D31",
                fg=CORES["muted"],
                font=("Segoe UI", 8, "bold"),
                padx=10,
                pady=5,
            ).pack(side="left", padx=(0, 8))

        titulos = tk.Frame(self.janela, bg=CORES["bg"])
        titulos.pack(fill="x", padx=28, pady=(0, 14))
        tk.Label(
            titulos,
            text="Seus dados, no formato certo.",
            bg=CORES["bg"],
            fg=CORES["text"],
            font=FONT_TITULO,
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            titulos,
            text="Arraste a ideia: escolha o arquivo, veja a pré-visualização e converta em segundos.",
            bg=CORES["bg"],
            fg=CORES["muted"],
            font=FONT_SUB,
            anchor="w",
        ).pack(anchor="w", pady=(4, 0))

        # Corpo: 2 colunas
        corpo = tk.Frame(self.janela, bg=CORES["bg"])
        corpo.pack(fill="both", expand=True, padx=28, pady=(0, 8))
        corpo.grid_columnconfigure(0, weight=0, minsize=350)
        corpo.grid_columnconfigure(1, weight=1)
        corpo.grid_rowconfigure(0, weight=1)

        # ---- coluna esquerda (controles) ----
        esquerda = tk.Frame(corpo, bg=CORES["bg"])
        esquerda.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        esquerda.grid_rowconfigure(2, weight=1)

        self._card_arquivo(esquerda)
        self._card_formatos(esquerda)

        # Botão converter + progresso
        acoes = tk.Frame(esquerda, bg=CORES["bg"])
        acoes.grid(row=2, column=0, sticky="new", pady=(14, 0))

        self.btn_converter = tk.Button(
            acoes,
            text="⚡  Converter e salvar",
            command=self.converter,
            bg=CORES["accent"],
            fg="#07111F",
            activebackground=CORES["accent_hover"],
            activeforeground="#07111F",
            relief="flat",
            cursor="hand2",
            font=("Segoe UI", 11, "bold"),
            padx=18,
            pady=13,
            borderwidth=0,
            highlightthickness=0,
            state="disabled",
            disabledforeground="#5B6B84",
        )
        self.btn_converter.pack(fill="x")
        self._hover(self.btn_converter, CORES["accent"], CORES["accent_hover"])

        self.progresso = ttk.Progressbar(
            acoes, style="Convert.Horizontal.TProgressbar", mode="indeterminate"
        )
        # fica oculto até converter (gerenciado por pack/forget)

        linha_limpar = tk.Frame(acoes, bg=CORES["bg"])
        linha_limpar.pack(fill="x", pady=(8, 0))
        tk.Button(
            linha_limpar,
            text="Limpar",
            command=self._limpar_tudo,
            bg=CORES["bg"],
            fg=CORES["muted_dark"],
            activebackground=CORES["bg"],
            activeforeground=CORES["muted"],
            relief="flat",
            cursor="hand2",
            font=("Segoe UI", 9, "underline"),
            borderwidth=0,
            highlightthickness=0,
        ).pack(side="right")
        tk.Label(
            linha_limpar,
            text="Ctrl+O para abrir arquivo",
            bg=CORES["bg"],
            fg="#3A4a63",
            font=("Segoe UI", 8),
        ).pack(side="left")
        self.janela.bind("<Control-o>", lambda _e: self.selecionar_arquivo())
        self.janela.bind("<Control-O>", lambda _e: self.selecionar_arquivo())

        # ---- coluna direita (preview) ----
        direita = tk.Frame(
            corpo,
            bg=CORES["card"],
            highlightbackground=CORES["border"],
            highlightthickness=1,
        )
        direita.grid(row=0, column=1, sticky="nsew")
        direita.grid_rowconfigure(1, weight=1)
        direita.grid_columnconfigure(0, weight=1)

        prev_header = tk.Frame(direita, bg=CORES["card"])
        prev_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 10))
        tk.Label(
            prev_header,
            text="PRÉ-VISUALIZAÇÃO",
            bg=CORES["card"],
            fg=CORES["muted_dark"],
            font=FONT_STEP,
        ).pack(side="left")
        self.badge_linhas = tk.Label(
            prev_header,
            text="nenhum arquivo",
            bg="#1A2438",
            fg="#9FB0C8",
            font=("Segoe UI", 8, "bold"),
            padx=10,
            pady=4,
        )
        self.badge_linhas.pack(side="right")

        tabela_wrap = tk.Frame(direita, bg=CORES["card"])
        tabela_wrap.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 8))
        tabela_wrap.grid_rowconfigure(0, weight=1)
        tabela_wrap.grid_columnconfigure(0, weight=1)

        self.tabela = ttk.Treeview(tabela_wrap, style="Dark.Treeview", show="headings")
        vsb = ttk.Scrollbar(
            tabela_wrap, orient="vertical", command=self.tabela.yview,
            style="Dark.Vertical.TScrollbar",
        )
        hsb = ttk.Scrollbar(
            tabela_wrap, orient="horizontal", command=self.tabela.xview,
            style="Dark.Horizontal.TScrollbar",
        )
        self.tabela.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tabela.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # estado vazio (sobrepõe a tabela)
        self.empty = tk.Frame(tabela_wrap, bg=CORES["card"])
        self.empty.grid(row=0, column=0, sticky="nsew")
        tk.Label(
            self.empty, text="👁️", bg=CORES["card"], fg=CORES["muted"],
            font=("Segoe UI", 28),
        ).pack(pady=(60, 6))
        tk.Label(
            self.empty, text="Nada para mostrar ainda",
            bg=CORES["card"], fg=CORES["text"], font=("Segoe UI", 11, "bold"),
        ).pack()
        tk.Label(
            self.empty,
            text="Escolha um arquivo Excel, CSV, JSON ou TSV\npara ver as primeiras linhas aqui.",
            bg=CORES["card"], fg=CORES["muted"], font=FONT_SMALL, justify="center",
        ).pack(pady=(6, 0))

        dica = tk.Frame(direita, bg="#101A2E", padx=16, pady=12)
        dica.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 16))
        tk.Label(
            dica,
            text="💡  Dica: arquivos grandes mostram só as 50 primeiras linhas, mas a conversão usa o arquivo completo.",
            bg="#101A2E",
            fg="#7E8DA6",
            font=("Segoe UI", 8),
            wraplength=480,
            justify="left",
        ).pack(anchor="w")

        # Status bar
        status = tk.Frame(
            self.janela, bg=CORES["panel"],
            highlightbackground=CORES["border"], highlightthickness=1,
        )
        status.pack(fill="x", padx=28, pady=(0, 18))
        self.dot = tk.Label(status, text="●", bg=CORES["panel"], fg=CORES["muted_dark"], font=("Segoe UI", 10))
        self.dot.pack(side="left", padx=(14, 6), pady=10)
        tk.Label(
            status, textvariable=self.status_texto, bg=CORES["panel"],
            fg=CORES["muted"], font=("Segoe UI", 9),
        ).pack(side="left", pady=10)
        self.status_ext = tk.Label(
            status, text="", bg=CORES["panel"], fg=CORES["muted_dark"], font=("Segoe UI", 8, "bold"),
        )
        self.status_ext.pack(side="right", padx=14, pady=10)

    # -- cards -----------------------------------------------------------
    def _base_card(self, pai) -> tk.Frame:
        return tk.Frame(
            pai, bg=CORES["card"],
            highlightbackground=CORES["border"], highlightthickness=1,
            padx=18, pady=16,
        )

    def _card_arquivo(self, pai) -> None:
        card = self._base_card(pai)
        card.grid(row=0, column=0, sticky="ew")
        tk.Label(
            card, text="01  —  ARQUIVO DE ORIGEM", bg=CORES["card"],
            fg=CORES["muted_dark"], font=FONT_STEP,
        ).pack(anchor="w", pady=(0, 10))

        # dropzone clicável
        self.dropzone = tk.Frame(
            card, bg=CORES["drop"],
            highlightbackground=CORES["border_light"], highlightthickness=1,
            padx=14, pady=16, cursor="hand2",
        )
        self.dropzone.pack(fill="x")
        self.dropzone.bind("<Button-1>", lambda _e: self.selecionar_arquivo())
        self.dropzone.bind("<Enter>", lambda _e: self.dropzone.configure(bg=CORES["card_hover"]))
        self.dropzone.bind("<Leave>", lambda _e: self.dropzone.configure(bg=CORES["drop"]))

        self.drop_icon = tk.Label(
            self.dropzone, text="📂", bg=CORES["drop"], fg=CORES["text"], font=("Segoe UI", 22),
        )
        self.drop_icon.pack()
        self.drop_icon.bind("<Button-1>", lambda _e: self.selecionar_arquivo())

        self.drop_titulo = tk.Label(
            self.dropzone, text="Clique para escolher", bg=CORES["drop"],
            fg=CORES["text"], font=("Segoe UI", 10, "bold"),
        )
        self.drop_titulo.pack(pady=(6, 2))
        self.drop_titulo.bind("<Button-1>", lambda _e: self.selecionar_arquivo())

        tk.Label(
            self.dropzone, text=".xlsx  •  .csv  •  .json  •  .tsv",
            bg=CORES["drop"], fg=CORES["muted_dark"], font=("Segoe UI", 8, "bold"),
        ).pack()

        # info do arquivo (oculta até selecionar)
        self.file_info = tk.Frame(card, bg=CORES["card"])
        self.file_icon = tk.Label(self.file_info, text="📄", bg=CORES["card"], font=("Segoe UI", 16))
        self.file_icon.pack(side="left", padx=(0, 10))
        textos = tk.Frame(self.file_info, bg=CORES["card"])
        textos.pack(side="left", fill="x", expand=True)
        self.arquivo_label = tk.Label(
            textos, text="", bg=CORES["card"], fg=CORES["text"],
            font=("Segoe UI", 10, "bold"), anchor="w",
        )
        self.arquivo_label.pack(anchor="w")
        self.arquivo_meta = tk.Label(
            textos, text="", bg=CORES["card"], fg=CORES["muted"],
            font=("Segoe UI", 8), anchor="w",
        )
        self.arquivo_meta.pack(anchor="w")
        tk.Button(
            self.file_info, text="✕", command=self._limpar_tudo,
            bg=CORES["card"], fg=CORES["muted_dark"],
            activebackground=CORES["card"], activeforeground=CORES["error"],
            relief="flat", cursor="hand2", font=("Segoe UI", 10, "bold"),
            borderwidth=0, highlightthickness=0,
        ).pack(side="right")

    def _card_formatos(self, pai) -> None:
        card = self._base_card(pai)
        card.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        topo = tk.Frame(card, bg=CORES["card"])
        topo.pack(fill="x", pady=(0, 10))
        tk.Label(
            topo, text="02  —  FORMATO DE DESTINO", bg=CORES["card"],
            fg=CORES["muted_dark"], font=FONT_STEP,
        ).pack(side="left")

        grade = tk.Frame(card, bg=CORES["card"])
        grade.pack(fill="x")
        grade.grid_columnconfigure(0, weight=1)
        grade.grid_columnconfigure(1, weight=1)

        self.format_btns: dict[str, tk.Button] = {}
        for i, (key, info) in enumerate(FORMATOS.items()):
            b = tk.Button(
                grade,
                text=f"{info['icon']}  {info['label']}\n{info['desc']}",
                command=lambda k=key: self._escolher_formato(k),
                relief="flat", cursor="hand2",
                font=("Segoe UI", 9, "bold"),
                justify="center",
                padx=10, pady=10,
                borderwidth=0, highlightthickness=1,
            )
            b.grid(row=i // 2, column=i % 2, sticky="ew", padx=4, pady=4)
            self.format_btns[key] = b

    # -- helpers visuais -------------------------------------------------
    @staticmethod
    def _hover(botao: tk.Button, normal: str, hover: str) -> None:
        def entrar(_e):
            if str(botao["state"]) == "normal":
                botao.configure(bg=hover)

        def sair(_e):
            if str(botao["state"]) == "normal":
                botao.configure(bg=normal)

        botao.bind("<Enter>", entrar)
        botao.bind("<Leave>", sair)

    def _escolher_formato(self, key: str) -> None:
        self.formato_key.set(key)
        self._atualizar_formato_visual()

    def _atualizar_formato_visual(self) -> None:
        sel = self.formato_key.get()
        for key, btn in self.format_btns.items():
            if key == sel:
                btn.configure(
                    bg="#12324A", fg=CORES["accent_hover"],
                    highlightbackground=CORES["accent"],
                    highlightcolor=CORES["accent"],
                )
            else:
                btn.configure(
                    bg="#1A2438", fg="#9FB0C8",
                    highlightbackground=CORES["border"],
                    highlightcolor=CORES["border"],
                )

    def _set_status(self, texto: str, tipo: str = "idle", ext: str | None = None) -> None:
        self.status_texto.set(texto)
        cores_dot = {
            "idle": CORES["muted_dark"], "busy": CORES["accent"],
            "ok": CORES["success"], "erro": CORES["error"],
        }
        self.dot.configure(fg=cores_dot.get(tipo, CORES["muted_dark"]))
        if ext is not None:
            self.status_ext.configure(text=ext)

    # -- ações -----------------------------------------------------------
    def selecionar_arquivo(self) -> None:
        if self.convertendo:
            return
        caminho = filedialog.askopenfilename(
            title="Selecionar arquivo",
            filetypes=[
                ("Arquivos suportados", "*.xlsx *.csv *.json *.tsv"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if caminho:
            self._carregar(caminho)

    def _carregar(self, caminho: str) -> None:
        try:
            self._set_status("Lendo arquivo…", "busy")
            self.janela.update_idletasks()
            df = carregar_arquivo(caminho)
        except (ValueError, OSError, UnicodeError, pd.errors.ParserError) as erro:
            self._set_status("Não foi possível ler o arquivo", "erro")
            messagebox.showerror("Erro", str(erro))
            return
        except Exception as erro:  # noqa: BLE001
            self._set_status("Erro inesperado ao ler", "erro")
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo.\n\n{erro}")
            return

        self.arquivo = caminho
        self.df = df
        nome = Path(caminho).name
        try:
            tam = tamanho_legivel(os.path.getsize(caminho))
        except OSError:
            tam = "—"
        linhas, colunas = df.shape

        # mostra info, esconde dropzone compacta? mantém ambos: dropzone vira "trocar"
        self.file_info.pack(fill="x", pady=(12, 0))
        self.arquivo_label.configure(text=nome if len(nome) <= 32 else nome[:29] + "…")
        self.arquivo_meta.configure(text=f"{tam}  •  {linhas} linhas  •  {colunas} colunas")
        self.drop_titulo.configure(text="Trocar arquivo")

        self.btn_converter.configure(state="normal")
        self.badge_linhas.configure(text=f"{linhas} linhas × {colunas} colunas")
        self.status_ext.configure(text=f"ORIGEM {Path(caminho).suffix.upper() or '—'}")
        self._set_status(f"Arquivo pronto — {nome}", "ok")
        self._mostrar_preview(df)

    def _mostrar_preview(self, df: pd.DataFrame) -> None:
        self.empty.grid_remove()
        cols = [str(c) for c in df.columns[:20]]
        self.tabela.configure(columns=cols)
        for c in self.tabela.get_children():
            self.tabela.delete(c)
        for c in cols:
            self.tabela.heading(c, text=c)
            self.tabela.column(c, width=130, minwidth=90, anchor="w")
        for _, row in df.head(50).iterrows():
            vals = [(str(v)[:60] if v == v else "") for v in list(row.values)[:20]]
            self.tabela.insert("", "end", values=vals)
        # zebra
        for i, iid in enumerate(self.tabela.get_children()):
            if i % 2 == 1:
                self.tabela.item(iid, tags=("odd",))
        try:
            self.tabela.tag_configure("odd", background="#162032")
        except Exception:
            pass

    def _limpar_tudo(self) -> None:
        if self.convertendo:
            return
        self.arquivo = ""
        self.df = None
        self.file_info.pack_forget()
        self.drop_titulo.configure(text="Clique para escolher")
        self.btn_converter.configure(state="disabled")
        self.badge_linhas.configure(text="nenhum arquivo")
        self.status_ext.configure(text="")
        self._set_status("Pronto para converter", "idle")
        for c in self.tabela.get_children():
            self.tabela.delete(c)
        self.tabela.configure(columns=())
        self.empty.grid()

    def converter(self) -> None:
        if not self.arquivo or self.df is None:
            messagebox.showwarning("Atenção", "Selecione um arquivo primeiro.")
            return
        if self.convertendo:
            return

        info = FORMATOS[self.formato_key.get()]
        extensao = info["ext"]
        nome_sugerido = f"{Path(self.arquivo).stem}_convertido{extensao}"
        destino = filedialog.asksaveasfilename(
            title="Salvar arquivo convertido",
            defaultextension=extensao,
            initialfile=nome_sugerido,
            filetypes=[(f'{info["label"]} ({extensao})', f"*{extensao}")],
        )
        if not destino:
            return

        self.convertendo = True
        self.btn_converter.configure(state="disabled", text="⏳  Convertendo…")
        self.progresso.pack(fill="x", pady=(10, 0))
        self.progresso.start(12)
        self._set_status(f"Convertendo para {info['label']}…", "busy")

        origem, df = self.arquivo, self.df

        def trabalho():
            try:
                # recarrega do disco para garantir arquivo completo (preview usa head)
                dados = carregar_arquivo(origem)
                salvar_arquivo(dados, destino)
                self.janela.after(0, lambda: self._fim_conversao(True, destino))
            except (ValueError, OSError, UnicodeError, pd.errors.ParserError) as erro:
                self.janela.after(0, lambda: self._fim_conversao(False, str(erro)))
            except Exception as erro:  # noqa: BLE001
                self.janela.after(0, lambda: self._fim_conversao(False, f"Erro inesperado.\n\n{erro}"))

        threading.Thread(target=trabalho, daemon=True).start()

    def _fim_conversao(self, ok: bool, detalhe: str) -> None:
        self.convertendo = False
        self.progresso.stop()
        self.progresso.pack_forget()
        self.btn_converter.configure(state="normal", text="⚡  Converter e salvar")
        if ok:
            self._set_status("Conversão concluída ✓", "ok")
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{detalhe}")
        else:
            self._set_status("Não foi possível converter", "erro")
            messagebox.showerror("Erro", detalhe)


def main() -> None:
    janela = tk.Tk()
    ConversorApp(janela)
    janela.mainloop()


if __name__ == "__main__":
    main()

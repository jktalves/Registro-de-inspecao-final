"""
Inspeção de Qualidade — 920.002.438
Tabela 1: grade com todos os 14 itens (página 1 do Word).
Tabela 2: características por item selecionado (página 2 do Word).
Tabela 3: grade F3 com todos os 14 itens × 8 posições (página 3 do Word).
"""

import os, tkinter as tk
from tkinter import ttk, messagebox
class CustomComboBox(tk.Frame):
    def __init__(self, master, values, command=None, width=22, font=None, **kwargs):
        super().__init__(master, **kwargs)
        self.values = values
        self.command = command
        self.var = tk.StringVar()
        self.btn = tk.Button(self, textvariable=self.var, width=width, font=font, relief="groove", anchor="w", command=self._show_list)
        self.btn.pack(fill="x")
        self.selected_index = 0
        self.var.set(self.values[0] if self.values else "")
        self.font = font
        self.listbox = None
        self.toplevel = None

    def _show_list(self):
        if self.toplevel:
            return
        self.toplevel = tk.Toplevel(self)
        self.toplevel.wm_overrideredirect(True)
        self.toplevel.lift()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height()
        self.toplevel.geometry(f"{self.winfo_width()}x180+{x}+{y}")
        frame = tk.Frame(self.toplevel, bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        scrollbar = tk.Scrollbar(frame, orient="vertical")
        self.listbox = tk.Listbox(frame, selectmode="browse", font=self.font, yscrollcommand=scrollbar.set, activestyle="dotbox")
        for v in self.values:
            self.listbox.insert("end", v)
        self.listbox.selection_set(self.selected_index)
        self.listbox.see(self.selected_index)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.bind("<ButtonRelease-1>", self._on_select)
        self.listbox.bind("<Return>", self._on_select)
        self.listbox.bind("<Escape>", lambda e: self._close_list())
        self.listbox.bind("<FocusOut>", lambda e: self._close_list())
        self.listbox.bind("<MouseWheel>", self._on_mousewheel)
        self.listbox.focus_set()

    def _on_mousewheel(self, event):
        self.listbox.yview_scroll(-1*(event.delta//120), "units")
        return "break"

    def _on_select(self, event=None):
        idxs = self.listbox.curselection()
        if idxs:
            idx = idxs[0]
            self.selected_index = idx
            self.var.set(self.values[idx])
            if self.command:
                self.command(idx)
        self._close_list()

    def _close_list(self):
        if self.toplevel:
            self.toplevel.destroy()
            self.toplevel = None
            self.listbox = None

    def set_values(self, values):
        self.values = values
        self.selected_index = 0
        self.var.set(self.values[0] if self.values else "")

    def current(self, idx):
        if 0 <= idx < len(self.values):
            self.selected_index = idx
            self.var.set(self.values[idx])

    def get(self):
        return self.var.get()

    def bind(self, event, handler):
        # For compatibility with ttk.Combobox usage
        pass
import sys
import win32com.client

# ── Arquivo ───────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)
else:
    _BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO = os.path.join(_BASE_DIR, "Planilha.doc")

# ── Linhas no documento ───────────────────────────────────────────────────────
T1_BASE = 3
T2_BASE = 5
T3_BASE = 5
N_ITENS = 14

# ── Colunas Tabela 1 ──────────────────────────────────────────────────────────
C1 = {"nqa":2,"nivel":3,"lote":4,"ord_servico":5,"nota_fiscal":6,
      "quant":7,"tam_amostra":8,"defeitos":9,"laudo":10,
      "instrumentos":11,"data":12}

# ── Colunas Tabela 2 (Word) ───────────────────────────────────────────────────
C2 = {"comp_min":2,"comp_max":3,"dext_min":4,"dext_max":5,
      "sent_min":6,"sent_max":7,"esp_min":8,"esp_max":9,
      "ret_min":10,"ret_max":11,"dfio_min":12,"dfio_max":13,
      "f1_min":14,"f1_max":15,"f2_min":16,"f2_max":17,
      "visto2":18}

# ── Colunas Tabela 3 (Word) — 8 posições F3 ──────────────────────────────────
C3 = {
    "f3_1min":2, "f3_1max":3,
    "f3_2min":4, "f3_2max":5,
    "f3_3min":6, "f3_3max":7,
    "f3_4min":8, "f3_4max":9,
    "f3_5min":10,"f3_5max":11,
    "f3_6min":12,"f3_6max":13,
    "f3_7min":14,"f3_7max":15,
    "f3_8min":16,"f3_8max":17,
    "visto3":18,
}

_F3 = (0.45, 0.55)

# ── Tolerâncias ───────────────────────────────────────────────────────────────
TOL = {
    "f1_min":(0.10,0.20), "f1_max":(0.10,0.20),
    "f2_min":(0.25,0.35), "f2_max":(0.25,0.35),
    **{f"f3_{i}min": _F3 for i in range(1,9)},
    **{f"f3_{i}max": _F3 for i in range(1,9)},
}

# ── Cores ─────────────────────────────────────────────────────────────────────
AZUL  = "#1C3D6B"
VERM  = "#B00020"
VERDE = "#1B6B35"
CINZA = "#777777"
BG    = "#EEF1F6"
WHITE = "#FFFFFF"
FORA  = "#FFE0E0"

CAMPOS_T2 = set(C2.keys())
CAMPOS_T3 = set(C3.keys())

# ── Helpers ───────────────────────────────────────────────────────────────────

def auto_formatar(txt):
    txt = txt.strip()
    if not txt:
        return txt
    if "," not in txt and "." not in txt and txt.isdigit():
        return "0," + txt
    return txt

def limpar(t):
    return t.replace("\r","").replace("\x07","").strip()

def to_float(t):
    return float(t.replace(",",".").strip())

def eh_num(t):
    try: to_float(t); return True
    except: return False

def fora_tol(campo, txt):
    if campo not in TOL or not eh_num(txt): return False
    v = to_float(txt); mn, mx = TOL[campo]
    return v < mn or v > mx

def ler_cel(tbl, r, c):
    try:
        txt = limpar(tbl.Cell(r, c).Range.Text)
        if txt in ("\x01", "\x07", ""):
            return ""
        return txt
    except:
        return ""

def ler_linha_cells(tbl, r):
    mapa = {}
    try:
        for cell in tbl.Range.Cells:
            if cell.RowIndex == r:
                mapa[cell.ColumnIndex] = limpar(cell.Range.Text)
    except:
        pass
    return mapa

def gravar_cel(tbl, r, c, val, verm=False):
    try:
        cel = tbl.Cell(r, c)
        cel.Range.Text = str(val)
        cel.Range.Font.Bold       = verm
        cel.Range.Font.ColorIndex = 6 if verm else 0  # 6=wdRed, 0=wdAutomatic
    except:
        pass

# ── Word ──────────────────────────────────────────────────────────────────────

def abrir_word():
    w = win32com.client.Dispatch("Word.Application")
    w.Visible = False
    d = w.Documents.Open(ARQUIVO)
    return w, d

def ler_item(t1, t2, t3, idx):
    r1, r2, r3 = T1_BASE+idx, T2_BASE+idx, T3_BASE+idx
    d = {"_r1":r1, "_r2":r2, "_r3":r3, "item":f"{idx+1:02d}"}
    for k, c in C1.items():
        d[k] = ler_cel(t1, r1, c)
    for k, c in C2.items():
        d[k] = ler_cel(t2, r2, c)
    for k, c in C3.items():
        d[k] = ler_cel(t3, r3, c)
    return d

def item_preenchido(d):
    return any(d.get(c,"") for c in ("lote","nota_fiscal","f1_min","f2_min","f3_1min"))

# ═════════════════════════════════════════════════════════════════════════════
# APP
# ═════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Inspeção de Qualidade — 920.002.438")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.minsize(900, 660)

        self.dados      = []
        self.idx        = 0
        self.vars       = {}
        self.entries    = {}
        self.t1_vars    = [{} for _ in range(N_ITENS)]
        self.t1_entries = [{} for _ in range(N_ITENS)]
        self.t3_vars    = [{} for _ in range(N_ITENS)]
        self.t3_entries = [{} for _ in range(N_ITENS)]
        self._inner_cvs = []   # canvases internos (T1 e T3) — bloqueiam scroll vertical
        self._word      = None
        self._doc       = None

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build()
        self._load()

    def _combo_mousewheel(self, event):
        # Impede propagação do scroll do mouse para o canvas principal
        return "break"

    # ── Layout principal ──────────────────────────────────────────────────────

    def _build(self):
        hdr = tk.Frame(self, bg=AZUL, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="REGISTRO DE INSPEÇÃO FINAL  —  920.002.438",
                 bg=AZUL, fg="white", font=("Segoe UI",13,"bold")).pack()
        tk.Label(hdr,
                 text="Mola Compressão  •  Material: Inox Ø 0,38–0,42  •  Cliente: UNIPAC / Jactos / Integra",
                 bg=AZUL, fg="#9ABBD4", font=("Segoe UI",8)).pack(pady=(1,0))

        # Navegação para Tabela 2 (per-item)
        nav = tk.Frame(self, bg=BG, pady=6)
        nav.pack(fill="x", padx=12)
        tk.Label(nav, text="Item (Tabela 2):", bg=BG,
                 font=("Segoe UI",9,"bold"), fg="#333").pack(side="left")
        self._combo = CustomComboBox(nav, values=[], width=22, font=("Segoe UI",9), command=self._on_combo_custom)
        self._combo.pack(side="left", padx=(4,10))
        btn = dict(bg=AZUL, fg="white", relief="flat",
                   font=("Segoe UI",9), padx=8, pady=2, cursor="hand2")
        tk.Button(nav, text="◀", command=self._prev, **btn).pack(side="left", padx=1)
        tk.Button(nav, text="▶", command=self._next, **btn).pack(side="left", padx=1)
        self._lbl_status = tk.Label(nav, text="", bg=BG,
                                     font=("Segoe UI",9,"bold"))
        self._lbl_status.pack(side="left", padx=10)

        self._build_form()

        rod = tk.Frame(self, bg=BG, pady=8)
        rod.pack(fill="x", padx=12)
        tk.Button(rod, text="   SALVAR NO WORD   ", command=self._save,
                  bg=VERM, fg="white", relief="flat",
                  font=("Segoe UI",11,"bold"), padx=16, pady=8,
                  cursor="hand2").pack(side="left")

    def _build_lista_btns(self):
        opcoes = []
        for d in self.dados:
            preen = item_preenchido(d)
            opcoes.append(f"Item {d['item']}  {'● Preenchido' if preen else '○ Em branco'}")
        self._combo.set_values(opcoes)
        self._combo.current(self.idx)

    def _atualizar_lista(self):
        opcoes = []
        for d in self.dados:
            preen = item_preenchido(d)
            opcoes.append(f"Item {d['item']}  {'● Preenchido' if preen else '○ Em branco'}")
        self._combo.set_values(opcoes)
        self._combo.current(self.idx)


    def _on_combo_custom(self, idx):
        if idx >= 0:
            self._select(idx)

    # ── Formulário ────────────────────────────────────────────────────────────

    def _build_form(self):
        wrap = tk.Frame(self, bg=BG)
        wrap.pack(fill="both", expand=True, padx=8)
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)

        self._cv = tk.Canvas(wrap, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=vsb.set)
        vsb.grid(row=0, column=1, sticky="ns")
        self._cv.grid(row=0, column=0, sticky="nsew")

        self._form = tk.Frame(self._cv, bg=BG)
        self._wid  = self._cv.create_window((0,0), window=self._form, anchor="nw")
        self._form.bind("<Configure>",
            lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>",
            lambda e: self._cv.itemconfig(self._wid, width=e.width))

        self._sec_t1()
        self._sec_t2()
        self._sec_t3()

        # Bind mousewheel após todos os widgets existirem
        self._cv.bind_all("<MouseWheel>", self._on_scroll)

    def _on_scroll(self, event):
        self._cv.yview_scroll(-1 * (event.delta // 120), "units")

    # ── Tabela 1 — grade 14 itens ─────────────────────────────────────────────

    def _sec_t1(self):
        self._titulo("TABELA 1  —  REGISTRO")
        outer = tk.Frame(self._form, bg=WHITE, relief="solid", bd=1)
        outer.pack(fill="x", pady=(0,2))

        cv = tk.Canvas(outer, bg=WHITE, highlightthickness=0, height=370)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=cv.xview)
        cv.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        cv.pack(fill="both", expand=True, padx=4, pady=4)

        inner = tk.Frame(cv, bg=WHITE)
        cv.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self._inner_cvs.append(cv)

        COLUNAS = [
            ("ITEM",                           None,           4),
            ("NQA",                            "nqa",          5),
            ("NIVEL",                          "nivel",        5),
            ("LOTE",                           "lote",         9),
            ("Nº ORD.\nSERVIÇO",               "ord_servico",  7),
            ("Nº NOTA\nFISCAL",                "nota_fiscal",  8),
            ("QUANT.\nLIBERADA",               "quant",        8),
            ("TAM.\nAMOSTRA",                  "tam_amostra",  6),
            ("Nº\nDEFEITOS",                   "defeitos",     6),
            ("LAUDO",                          "laudo",        5),
            ("VALIDADE E Nº DOS INSTRUMENTOS", "instrumentos", 28),
            ("DATA",                           "data",         8),
        ]

        for ci, (hdr, _, w) in enumerate(COLUNAS):
            tk.Label(inner, text=hdr, bg=AZUL, fg="white",
                     font=("Segoe UI",8,"bold"), width=w,
                     anchor="center", justify="center",
                     relief="flat", padx=2, pady=3
                    ).grid(row=0, column=ci, padx=1, pady=1, sticky="ew")

        for ii in range(N_ITENS):
            tk.Label(inner, text=f"{ii+1:02d}", bg=WHITE, fg=AZUL,
                     font=("Segoe UI",9,"bold"), width=4, anchor="center"
                    ).grid(row=ii+1, column=0, padx=1, pady=1)
            for ci, (_, campo, w) in enumerate(COLUNAS[1:], start=1):
                v = tk.StringVar()
                self.t1_vars[ii][campo] = v
                e = tk.Entry(inner, textvariable=v, width=w,
                             font=("Segoe UI",9), relief="solid", bd=1,
                             bg=WHITE, justify="center")
                e.grid(row=ii+1, column=ci, padx=1, pady=1)
                self.t1_entries[ii][campo] = e

        leg = tk.Frame(outer, bg=WHITE)
        leg.pack(fill="x", padx=8, pady=(2,6))
        tk.Label(leg,
                 text="LEGENDA DE SÍMBOLOS PREENCHIMENTO EM CAMPO – LAUDO  E  INSTRUMENTOS DE MEDIÇÃO",
                 bg=WHITE, fg=AZUL, font=("Segoe UI",8,"bold")).pack(anchor="w")
        tk.Label(leg,
                 text="APROVADO    SELECIONAR    REPROVADO    APROVADO CONDICIONAL"
                      "    PQ= Paquímetro;  B= Balança;  PP= Projetor de Perfil",
                 bg=WHITE, fg="#444", font=("Segoe UI",8)).pack(anchor="w")

    # ── Tabela 2 — características por item ──────────────────────────────────

    def _sec_t2(self):
        self._titulo("TABELA 2  —  CARACTERÍSTICAS ENCONTRADAS")
        f = self._card()

        for col, txt, w in [(0,"Característica",48),(1,"MIN",11),(2,"MAX",11)]:
            tk.Label(f, text=txt, bg=WHITE, fg=AZUL,
                     font=("Segoe UI",9,"bold"), width=w,
                     anchor="w" if col==0 else "center"
                     ).grid(row=0, column=col, padx=6 if col==0 else 3, sticky="w")

        medicoes = [
            ("1  |  L0 53,00 +1,00mm",                 "comp_min","comp_max", None),
            ("2  |  Ø Ext. 4,60 ±0,1mm",               "dext_min","dext_max", None),
            ("3  |  Sentido Direito",                   "sent_min","sent_max", None),
            ("4  |  Ig. 20° / If. 18° Espiras",        "esp_min", "esp_max",  None),
            ("5  |  Sem Retifica",                      "ret_min", "ret_max",  None),
            ("6  |  Ø Fio 0,38 – 0,42mm",              "dfio_min","dfio_max", None),
            ("7  |  L1: 43,10   F1: 0,10 – 0,20 Kgf",  "f1_min",  "f1_max",  (0.10,0.20)),
            ("8  |  L2: 32,50   F2: 0,25 – 0,35 Kgf",  "f2_min",  "f2_max",  (0.25,0.35)),
        ]
        self._grid_med(f, medicoes)

        r = len(medicoes) + 1
        tk.Label(f, text="VISTO DE INSPEÇÃO", bg=WHITE, fg=AZUL,
                 font=("Segoe UI",9,"bold"), anchor="w", width=48
                 ).grid(row=r, column=0, padx=6, sticky="w", pady=(6,2))
        self._ent_grid(f, "visto2", 26, r, 1, colspan=2)

    # ── Tabela 3 — grade F3 (14 itens × 8 posições) ──────────────────────────

    def _sec_t3(self):
        self._titulo("TABELA 3  —  CARACTERÍSTICAS ENCONTRADAS")
        outer = tk.Frame(self._form, bg=WHITE, relief="solid", bd=1)
        outer.pack(fill="x", pady=(0,2))

        # Info do valor de referência
        info = tk.Frame(outer, bg=WHITE)
        info.pack(fill="x", padx=10, pady=(6,0))
        tk.Label(info, text="L3: 22,70  |  F3: 0,45 – 0,55 Kgf",
                 bg=WHITE, fg=VERM, font=("Segoe UI",9,"bold")).pack(anchor="w")

        # Canvas com scroll horizontal
        cv = tk.Canvas(outer, bg=WHITE, highlightthickness=0, height=390)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=cv.xview)
        cv.configure(xscrollcommand=hsb.set)
        hsb.pack(side="bottom", fill="x")
        cv.pack(fill="both", expand=True, padx=4, pady=4)

        inner = tk.Frame(cv, bg=WHITE)
        cv.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        self._inner_cvs.append(cv)

        W_ITEM = 4
        W_MED  = 7
        W_VISTO = 8

        # ── Cabeçalho linha 0: ITEM | 1..8 (colspan=2 cada) | VISTO (rowspan=2)
        tk.Label(inner, text="ITEM", bg=AZUL, fg="white",
                 font=("Segoe UI",8,"bold"), width=W_ITEM,
                 anchor="center", padx=2, pady=3
                ).grid(row=0, column=0, padx=1, pady=1, rowspan=2, sticky="nsew")

        for pos in range(1, 9):
            col = (pos - 1) * 2 + 1
            tk.Label(inner, text=str(pos), bg=AZUL, fg="white",
                     font=("Segoe UI",8,"bold"), width=W_MED*2,
                     anchor="center", padx=2, pady=2
                    ).grid(row=0, column=col, columnspan=2, padx=1, pady=1, sticky="ew")
            for j, sub in enumerate(("MIN","MAX")):
                tk.Label(inner, text=sub, bg=AZUL, fg="white",
                         font=("Segoe UI",7), width=W_MED,
                         anchor="center"
                        ).grid(row=1, column=col+j, padx=1, pady=1, sticky="ew")

        tk.Label(inner, text="VISTO DE\nINSPEÇÃO", bg=AZUL, fg="white",
                 font=("Segoe UI",7,"bold"), width=W_VISTO,
                 anchor="center", justify="center", padx=2, pady=3
                ).grid(row=0, column=17, padx=1, pady=1, rowspan=2, sticky="nsew")

        # ── Linhas de dados
        for ii in range(N_ITENS):
            tk.Label(inner, text=f"{ii+1:02d}", bg=WHITE, fg=AZUL,
                     font=("Segoe UI",9,"bold"), width=W_ITEM, anchor="center"
                    ).grid(row=ii+2, column=0, padx=1, pady=1)

            for pos in range(1, 9):
                col = (pos - 1) * 2 + 1
                for j, suf in enumerate(("min","max")):
                    campo = f"f3_{pos}{suf}"
                    v = tk.StringVar()
                    self.t3_vars[ii][campo] = v
                    v.trace_add("write", lambda *a, i=ii, c=campo: self._chk_t3(i, c))
                    e = tk.Entry(inner, textvariable=v, width=W_MED,
                                 font=("Segoe UI",9), relief="solid", bd=1,
                                 bg=WHITE, justify="center")
                    e.grid(row=ii+2, column=col+j, padx=1, pady=1)
                    e.bind("<FocusOut>", lambda ev, i=ii, c=campo: self._auto_fmt_t3(i, c))
                    e.bind("<Return>",   lambda ev, i=ii, c=campo: self._auto_fmt_t3(i, c))
                    e.bind("<Tab>",      lambda ev, i=ii, c=campo: self._auto_fmt_t3(i, c))
                    self.t3_entries[ii][campo] = e

            # VISTO
            v = tk.StringVar()
            self.t3_vars[ii]["visto3"] = v
            e = tk.Entry(inner, textvariable=v, width=W_VISTO,
                         font=("Segoe UI",9), relief="solid", bd=1,
                         bg=WHITE, justify="center")
            e.grid(row=ii+2, column=17, padx=1, pady=1)
            self.t3_entries[ii]["visto3"] = e

        tk.Label(outer,
                 text="⚠  Valores fora de 0,45 – 0,55 ficam VERMELHOS na tela e no Word.",
                 bg=WHITE, fg=CINZA, font=("Segoe UI",8)
                ).pack(anchor="w", padx=8, pady=(2,6))

    # ── Helpers de construção ─────────────────────────────────────────────────

    def _titulo(self, txt):
        f = tk.Frame(self._form, bg=AZUL, pady=4)
        f.pack(fill="x", pady=(10,0))
        tk.Label(f, text=txt, bg=AZUL, fg="white",
                 font=("Segoe UI",9,"bold"), padx=10).pack(anchor="w")

    def _card(self):
        f = tk.Frame(self._form, bg=WHITE, relief="solid", bd=1, padx=12, pady=10)
        f.pack(fill="x", pady=(0,2))
        return f

    def _ent_grid(self, p, campo, w, row, col, colspan=1):
        v = self._var(campo)
        e = tk.Entry(p, textvariable=v, width=w,
                     font=("Segoe UI",9), relief="solid", bd=1,
                     justify="center", bg=WHITE)
        e.grid(row=row, column=col, pady=3, padx=3, columnspan=colspan)
        if campo in CAMPOS_T2:
            e.bind("<FocusOut>", lambda ev, c=campo: self._auto_fmt(c))
            e.bind("<Return>",   lambda ev, c=campo: self._auto_fmt(c))
            e.bind("<Tab>",      lambda ev, c=campo: self._auto_fmt(c))
        self.entries[campo] = e

    def _var(self, campo):
        if campo not in self.vars:
            sv = tk.StringVar()
            sv.trace_add("write", lambda *a, c=campo: self._chk(c))
            self.vars[campo] = sv
        return self.vars[campo]

    def _grid_med(self, parent, linhas):
        for i, (lbl, cmin, cmax, tol) in enumerate(linhas):
            r = i + 1
            tem = tol is not None
            fg  = VERM if tem else "#222"
            fnt = ("Segoe UI",9,"bold") if tem else ("Segoe UI",9)
            tk.Label(parent, text=lbl, bg=WHITE, fg=fg, font=fnt,
                     anchor="w", width=48
                     ).grid(row=r, column=0, sticky="w", pady=2, padx=6)
            self._ent_grid(parent, cmin, 11, r, 1)
            self._ent_grid(parent, cmax, 11, r, 2)
            if tem:
                tk.Label(parent,
                         text=f"({tol[0]:.2f} – {tol[1]:.2f})",
                         bg=WHITE, fg=CINZA, font=("Segoe UI",7)
                         ).grid(row=r, column=3, sticky="w", padx=4)

    # ── Auto-formatação ───────────────────────────────────────────────────────

    def _auto_fmt(self, campo):
        txt = self.vars[campo].get()
        novo = auto_formatar(txt)
        if novo != txt:
            self.vars[campo].set(novo)
        else:
            self._chk(campo)

    def _auto_fmt_t3(self, ii, campo):
        txt = self.t3_vars[ii][campo].get()
        novo = auto_formatar(txt)
        if novo != txt:
            self.t3_vars[ii][campo].set(novo)
        else:
            self._chk_t3(ii, campo)

    # ── Verificação de tolerância ─────────────────────────────────────────────

    def _chk(self, campo):
        e = self.entries.get(campo)
        if e is None or campo not in TOL: return
        txt = self.vars[campo].get()
        fora = fora_tol(campo, txt)
        e.config(bg=FORA if fora else WHITE, fg=VERM if fora else "black")

    def _chk_t3(self, ii, campo):
        e = self.t3_entries[ii].get(campo)
        if e is None or campo not in TOL: return
        txt = self.t3_vars[ii][campo].get()
        fora = fora_tol(campo, txt)
        e.config(bg=FORA if fora else WHITE, fg=VERM if fora else "black")

    def _chk_todos(self):
        for c in TOL:
            self._chk(c)

    # ── Carregar ──────────────────────────────────────────────────────────────

    def _load(self):
        self.title("Inspeção — lendo documento…")
        self.update()

        if not os.path.exists(ARQUIVO):
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{ARQUIVO}")
            return
        try:
            self._word, self._doc = abrir_word()
            t1 = self._doc.Tables(1)
            t2 = self._doc.Tables(2)
            t3 = self._doc.Tables(3)
            self.dados = [ler_item(t1, t2, t3, i) for i in range(N_ITENS)]
        except Exception as ex:
            messagebox.showerror("Erro ao ler", str(ex))
            return

        # Preenche Tabela 1 (grade)
        for ii in range(N_ITENS):
            di = self.dados[ii]
            for k in C1.keys():
                if k in self.t1_vars[ii]:
                    self.t1_vars[ii][k].set(di.get(k, ""))

        # Preenche Tabela 3 (grade F3)
        for ii in range(N_ITENS):
            di = self.dados[ii]
            for k in C3.keys():
                if k in self.t3_vars[ii]:
                    self.t3_vars[ii][k].set(di.get(k, ""))
                    if k in TOL:
                        self._chk_t3(ii, k)

        self._build_lista_btns()
        self._select(0)

    def _select(self, idx):
        self.idx = idx
        d = self.dados[idx]
        for campo, var in self.vars.items():
            var.set(d.get(campo, ""))
        preen = item_preenchido(d)
        self._lbl_status.config(
            text="● Preenchido" if preen else "○ Em branco",
            fg=VERDE if preen else CINZA)
        self._combo.current(idx)
        self._chk_todos()
        self.title(f"Inspeção 920.002.438  —  Item {d['item']}")

    def _on_close(self):
        if self._word is not None:
            try:
                self._doc.Close(False)
                self._word.Quit()
            except:
                pass
        self.destroy()

    def _prev(self):
        if self.idx > 0:
            self._select(self.idx - 1)

    def _next(self):
        if self.idx < N_ITENS - 1:
            self._select(self.idx + 1)

    # ── Salvar ────────────────────────────────────────────────────────────────

    def _save(self):
        if not self.dados: return

        # Coleta Tabela 1: todos os itens
        for ii in range(N_ITENS):
            di = dict(self.dados[ii])
            for k in C1.keys():
                if k in self.t1_vars[ii]:
                    di[k] = self.t1_vars[ii][k].get()
            self.dados[ii] = di

        # Coleta Tabela 2: item selecionado
        d = dict(self.dados[self.idx])
        for campo, var in self.vars.items():
            d[campo] = var.get()
        self.dados[self.idx] = d

        # Coleta Tabela 3: todos os itens
        for ii in range(N_ITENS):
            di = dict(self.dados[ii])
            for k in C3.keys():
                if k in self.t3_vars[ii]:
                    di[k] = self.t3_vars[ii][k].get()
            self.dados[ii] = di

        self.title("Salvando…")
        self.update()

        def _executar_save():
            t1 = self._doc.Tables(1)
            t2 = self._doc.Tables(2)
            t3 = self._doc.Tables(3)

            for ii in range(N_ITENS):
                di = self.dados[ii]
                for k, c in C1.items():
                    gravar_cel(t1, di["_r1"], c, di.get(k, ""))

            di = self.dados[self.idx]
            for k, c in C2.items():
                v = di.get(k, "")
                gravar_cel(t2, di["_r2"], c, v, verm=fora_tol(k, v))

            for ii in range(N_ITENS):
                di = self.dados[ii]
                for k, c in C3.items():
                    v = di.get(k, "")
                    gravar_cel(t3, di["_r3"], c, v, verm=fora_tol(k, v))

            self._doc.Save()

        try:
            if self._doc is None:
                self._word, self._doc = abrir_word()
            _executar_save()
        except Exception as ex:
            # RPC indisponível (-2147023174): Word fechou — reconecta e tenta de novo
            codigo = getattr(ex, 'args', [None])[0] if ex.args else None
            if codigo == -2147023174:
                try:
                    self._word, self._doc = abrir_word()
                    _executar_save()
                except Exception as ex2:
                    messagebox.showerror("Erro ao salvar", str(ex2))
                    return
            else:
                messagebox.showerror("Erro ao salvar", str(ex))
                return

        self._atualizar_lista()
        self.title("Inspeção 920.002.438  —  Salvo")
        messagebox.showinfo("Sucesso",
            "Salvo no Word!\n\n"
            "Valores do documento atualizados.")

# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()

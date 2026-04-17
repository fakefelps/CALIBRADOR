# -*- coding: utf-8 -*-
"""
CALIBRADOR DO MEMORIAL — Morais Engenharia
Ajuste visual de assinatura e checkboxes do Memorial Excel.

4 estados de checkbox:
  1. Esgoto SIM
  2. Esgoto NÃO
  3. Geminados condomínio = SIM
  4. Geminados condomínio = Não se aplica
"""

import multiprocessing
multiprocessing.freeze_support()

import os, sys, shutil, zipfile, tempfile, threading, traceback, time
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import win32com.client
import pythoncom
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ── Paleta ────────────────────────────────────────────────────────────
BG     = "#0f1923"
CAMPO  = "#243447"
ACENTO = "#2e86de"
VERDE  = "#27ae60"
VERM   = "#e74c3c"
TEXTO  = "#e8edf3"
TEXTO2 = "#7a99b8"
BORDA  = "#2a3f58"
LOG_BG = "#0a1118"
LOG_FG = "#4ecb8d"
LARANJ = "#e67e22"

# ── Shapes confirmados via drawing1.xml ──────────────────────────────
SHAPE_ESGOTO_SIM = "QO012,12.L0C0;L0C-34^"    # AM70
SHAPE_ESGOTO_NAO = "QO012,22.L0C0;L0C-37^"    # AP70
SHAPE_COND_SIM   = "QOCN,13.L0C-32;L0C-34^"   # AM65
SHAPE_COND_NAO   = "QOCN,23.L0C-35;L0C-37^"   # AP65
SHAPE_COND_NSA   = "QOCN,33.L0C-38;L0C-40^"   # AS65
SHAPE_LOT_NSA    = "QOCI,33.L0C-38;L0C-40^"   # AS64 — loteamentos fixo NSA

ESTADOS = {
    1: ("Esgoto — SIM",              ACENTO),
    2: ("Esgoto — NÃO",              VERM),
    3: ("Condomínio — SIM",          VERDE),
    4: ("Condomínio — Não se aplica",LARANJ),
}

# ── Defaults ──────────────────────────────────────────────────────────
# chkN_ancora / chkN_off_x / chkN_off_y / chkN_larg / chkN_alt
DEFAULT = {
    "ass_ancora": "AE72", "ass_off_x": 10, "ass_off_y": -5,
    "ass_larg": 170, "ass_alt": 55,
    "chk1_ancora": "AM70", "chk1_off_x": 0, "chk1_off_y": 0, "chk1_larg": 12, "chk1_alt": 12,
    "chk2_ancora": "AP70", "chk2_off_x": 0, "chk2_off_y": 0, "chk2_larg": 12, "chk2_alt": 12,
    "chk3_ancora": "AM65", "chk3_off_x": 0, "chk3_off_y": 0, "chk3_larg": 12, "chk3_alt": 12,
    "chk4_ancora": "AS65", "chk4_off_x": 0, "chk4_off_y": 0, "chk4_larg": 12, "chk4_alt": 12,
}


# ── Helpers COM ────────────────────────────────────────────────────────

def _fechar_excel(xl, wb):
    if wb:
        try: wb.Close(SaveChanges=False)
        except: pass
    if xl:
        try: xl.Quit()
        except: pass
        try:
            import subprocess
            subprocess.run(["taskkill","/F","/IM","EXCEL.EXE"],
                           capture_output=True, creationflags=0x08000000)
        except: pass


def _criar_xl():
    import subprocess
    def _try():
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except: pass
        try: xl.DisplayAlerts = False
        except: pass
        try: xl.ScreenUpdating = False
        except: pass
        _ = xl.Workbooks.Count
        return xl
    try:
        return _try()
    except:
        pass
    try:
        subprocess.run(["taskkill","/F","/IM","EXCEL.EXE"],
                       capture_output=True, creationflags=0x08000000)
    except: pass
    time.sleep(1)
    return _try()


def _xls_para_xlsx(path):
    """Converte .xls → .xlsx temp. Retorna (path, criou_temp)."""
    if str(path).lower().endswith(".xlsx"):
        return str(path), False
    tmp = tempfile.mktemp(suffix=".xlsx")
    pythoncom.CoInitialize()
    xl = None; wb = None
    try:
        xl = _criar_xl()
        wb = xl.Workbooks.Open(os.path.abspath(str(path)))
        wb.SaveAs(os.path.abspath(tmp), FileFormat=51)
        return tmp, True
    finally:
        _fechar_excel(xl, wb)
        pythoncom.CoUninitialize()


def _quadrado_preto():
    tmp = tempfile.mktemp(suffix=".png")
    Image.new("RGBA", (20, 20), (0, 0, 0, 255)).save(tmp)
    return tmp


def _inserir_img(ws, img_path, ancora, off_x, off_y, larg, alt):
    cell = ws.Range(ancora)
    ws.Shapes.AddPicture(
        os.path.abspath(img_path), False, True,
        cell.Left + off_x, cell.Top + off_y, larg, alt,
    )


# ── Método nativo ─────────────────────────────────────────────────────

def _nativo(xlsx_path, esgoto_sim, cond_val, log):
    """
    Pinta shapes:
      - esgoto linha 70
      - condomínios linha 65
      - loteamentos linha 64 → sempre Não se aplica
    """
    import shutil as _sh

    mapa = {
        SHAPE_ESGOTO_SIM: "000000" if esgoto_sim          else "FFFFFF",
        SHAPE_ESGOTO_NAO: "000000" if not esgoto_sim      else "FFFFFF",
        SHAPE_COND_SIM:   "000000" if cond_val == "sim"   else "FFFFFF",
        SHAPE_COND_NAO:   "000000" if cond_val == "nao"   else "FFFFFF",
        SHAPE_COND_NSA:   "000000" if cond_val == "nao_se_aplica" else "FFFFFF",
        SHAPE_LOT_NSA:    "000000",
    }

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    _sh.copy2(xlsx_path, tmp)
    ok = 0

    try:
        with zipfile.ZipFile(tmp, "r") as zi, \
             zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                data = zi.read(item.filename)
                if item.filename.startswith("xl/drawings/drawing") and \
                   item.filename.endswith(".xml"):
                    try:
                        root = etree.fromstring(data)
                        XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                        A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
                        for sp in root.iter(f"{{{XDR}}}sp"):
                            cnv  = sp.find(f".//{{{XDR}}}cNvPr")
                            nome = cnv.get("name","") if cnv is not None else ""
                            if nome not in mapa:
                                continue
                            cor   = mapa[nome]
                            sp_pr = sp.find(f".//{{{XDR}}}spPr")
                            if sp_pr is None:
                                continue
                            for ft in [f"{{{A}}}solidFill",
                                       f"{{{A}}}noFill",
                                       f"{{{A}}}gradFill"]:
                                ex = sp_pr.find(ft)
                                if ex is not None:
                                    sp_pr.remove(ex)
                            sf = etree.SubElement(sp_pr, f"{{{A}}}solidFill")
                            sc = etree.SubElement(sf,    f"{{{A}}}srgbClr")
                            sc.set("val", cor)
                            xfrm = sp_pr.find(f"{{{A}}}xfrm")
                            sp_pr.remove(sf)
                            if xfrm is not None:
                                xfrm.addnext(sf)
                            else:
                                sp_pr.insert(0, sf)
                            ok += 1
                        data = etree.tostring(root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                    except Exception as e:
                        log(f"  ⚠ XML: {e}")
                zo.writestr(item, data)
        os.unlink(tmp)
        log(f"  ✓ {ok} shape(s) modificado(s)")
        return ok > 0
    except Exception as e:
        log(f"  ✗ {e}")
        try: os.unlink(tmp)
        except: pass
        return False


# ── gerar_preview ─────────────────────────────────────────────────────

def gerar_preview(memorial_path, ass_img, cfg, estado, modo, saida, log):
    pythoncom.CoInitialize()
    xl = None; wb = None

    # mapa estado → parâmetros
    esgoto_sim = estado == 1
    cond_map   = {3: "sim", 4: "nao_se_aplica"}
    cond_val   = cond_map.get(estado, "nao_se_aplica")
    pfx        = f"chk{estado}_"

    try:
        # Preparar xlsx
        log("• Preparando arquivo...")
        xlsx_tmp, criou = _xls_para_xlsx(memorial_path)
        src_r  = os.path.realpath(xlsx_tmp)
        dest_r = os.path.realpath(saida)
        if src_r == dest_r:
            inter = tempfile.mktemp(suffix=".xlsx")
            shutil.copy2(xlsx_tmp, inter)
            if criou:
                try: os.unlink(xlsx_tmp)
                except: pass
            shutil.move(inter, saida)
        else:
            shutil.copy2(xlsx_tmp, saida)
            if criou:
                try: os.unlink(xlsx_tmp)
                except: pass

        # Abrir
        log("• Abrindo no Excel...")
        xl = _criar_xl()
        wb = xl.Workbooks.Open(os.path.abspath(saida))
        try: ws = wb.Worksheets("ElemConstrutivos")
        except: ws = wb.Worksheets(1)

        # Assinatura
        if ass_img and os.path.exists(ass_img):
            log(f"• Assinatura em {cfg['ass_ancora']} "
                f"({cfg['ass_larg']}×{cfg['ass_alt']}pt)...")
            _inserir_img(ws, ass_img,
                         cfg["ass_ancora"],
                         cfg["ass_off_x"], cfg["ass_off_y"],
                         cfg["ass_larg"],  cfg["ass_alt"])
        else:
            log("  ⚠ Assinatura não selecionada")

        # Checkbox
        if modo == "nativo":
            log(f"• Checkboxes NATIVO (estado {estado}: {ESTADOS[estado][0]})...")
            wb.Save()
            _fechar_excel(xl, wb)
            xl = None; wb = None
            pythoncom.CoUninitialize()
            time.sleep(1)
            _nativo(saida, esgoto_sim, cond_val, log)
            pythoncom.CoInitialize()
            xl = _criar_xl()
            wb = xl.Workbooks.Open(os.path.abspath(saida))
            try: ws = wb.Worksheets("ElemConstrutivos")
            except: ws = wb.Worksheets(1)
        else:
            log(f"• Checkbox IMAGEM {pfx}ancora={cfg[pfx+'ancora']} "
                f"({cfg[pfx+'larg']}×{cfg[pfx+'alt']}pt)...")
            q = _quadrado_preto()
            _inserir_img(ws, q,
                         cfg[pfx+"ancora"],
                         cfg[pfx+"off_x"], cfg[pfx+"off_y"],
                         cfg[pfx+"larg"],  cfg[pfx+"alt"])

        # PDF
        wb.Save()
        pdf = saida.replace(".xlsx", ".pdf")
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf), 0, True, False)
        log(f"✅ PDF gerado!")

        # Valores
        log("─" * 44)
        log("VALORES PARA app.py:")
        log(f'  ASSINATURA_EXCEL_ANCORA      = "{cfg["ass_ancora"]}"')
        log(f'  ASSINATURA_EXCEL_OFFSET_X_PT = {cfg["ass_off_x"]}')
        log(f'  ASSINATURA_EXCEL_OFFSET_Y_PT = {cfg["ass_off_y"]}')
        log(f'  ASSINATURA_EXCEL_LARGURA_PT  = {cfg["ass_larg"]}')
        log(f'  ASSINATURA_EXCEL_ALTURA_PT   = {cfg["ass_alt"]}')
        for n in range(1, 5):
            p = f"chk{n}_"
            log(f'  # {n}. {ESTADOS[n][0]}')
            log(f'  CHK{n} = ancora={DEFAULT[p+"ancora"]} '
                f'off=({DEFAULT[p+"off_x"]},{DEFAULT[p+"off_y"]}) '
                f'tam={DEFAULT[p+"larg"]}x{DEFAULT[p+"alt"]}')

        try: os.startfile(pdf)
        except: pass

    except Exception as e:
        log(f"✗ ERRO: {e}")
        log(traceback.format_exc())
    finally:
        _fechar_excel(xl, wb)
        pythoncom.CoUninitialize()


# ── Interface ─────────────────────────────────────────────────────────

class Calibrador(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Calibrador do Memorial — Morais Engenharia")
        self.geometry("940x840")
        self.configure(bg=BG)
        self._vars = {}
        self._estado_atual = 1
        self._criar_ui()

    # ── Construção da UI ──────────────────────

    def _criar_ui(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=BG, pady=12)
        hdr.pack(fill="x", padx=22)
        tk.Label(hdr, text="CALIBRADOR DO MEMORIAL",
                 font=("Segoe UI",16,"bold"), fg=TEXTO, bg=BG).pack(anchor="w")
        tk.Label(hdr, text="Ajuste assinatura e checkboxes • copie os valores para o app.py",
                 font=("Segoe UI",9), fg=TEXTO2, bg=BG).pack(anchor="w")
        tk.Frame(self, bg=BORDA, height=1).pack(fill="x")

        # Botão fixo no rodapé
        rodape = tk.Frame(self, bg=BG, pady=10)
        rodape.pack(side="bottom", fill="x", padx=22)
        tk.Frame(rodape, bg=BORDA, height=1).pack(fill="x", pady=(0,8))
        self.btn = tk.Button(rodape,
                             text="⚡  GERAR PREVIEW (abre PDF automaticamente)",
                             command=self._iniciar,
                             bg=ACENTO, fg=TEXTO, relief="flat",
                             font=("Segoe UI",12,"bold"), pady=10)
        self.btn.pack(fill="x")
        self.var_status = tk.StringVar(value="Aguardando...")
        tk.Label(rodape, textvariable=self.var_status,
                 bg=BG, fg=TEXTO2, font=("Segoe UI",9)).pack(anchor="w", pady=(4,0))

        # Corpo
        tk.Frame(self, bg=BORDA, height=1).pack(fill="x")
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)

        # Coluna esquerda com scroll
        outer = tk.Frame(body, bg=BG, width=350)
        outer.pack(side="left", fill="y")
        outer.pack_propagate(False)
        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        sb = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        col_esq = tk.Frame(canvas, bg=BG)
        wid = canvas.create_window((0,0), window=col_esq, anchor="nw")
        col_esq.bind("<Configure>", lambda e: (
            canvas.configure(scrollregion=canvas.bbox("all")),
            canvas.itemconfig(wid, width=canvas.winfo_width()),
        ))
        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

        # Coluna direita
        col_dir = tk.Frame(body, bg=BG)
        col_dir.pack(side="right", fill="both", expand=True, padx=(4,22))

        # ── Arquivos ──
        self._T(col_esq, "ARQUIVOS")
        self.var_memorial   = tk.StringVar()
        self.var_assinatura = tk.StringVar()
        self._arq(col_esq, self.var_memorial,   "Memorial (.xls/.xlsx)", [("Excel","*.xls *.xlsx")])
        self._arq(col_esq, self.var_assinatura, "Assinatura (.png/.jpg)", [("Imagem","*.png *.jpg *.jpeg")])

        # ── Assinatura ──
        self._T(col_esq, "ASSINATURA NO MEMORIAL")
        self._C(col_esq, "Célula âncora", "ass_ancora", DEFAULT["ass_ancora"])
        self._S(col_esq, "Offset X (pt)", "ass_off_x",  DEFAULT["ass_off_x"],  -200, 200)
        self._S(col_esq, "Offset Y (pt)", "ass_off_y",  DEFAULT["ass_off_y"],  -200, 200)
        self._S(col_esq, "Largura (pt)",  "ass_larg",   DEFAULT["ass_larg"],    10,  400)
        self._S(col_esq, "Altura (pt)",   "ass_alt",    DEFAULT["ass_alt"],     10,  200)
        tk.Label(col_esq, text="1 cm ≈ 28 pt  |  1\" = 72 pt",
                 bg=BG, fg=TEXTO2, font=("Segoe UI",8)).pack(anchor="w", pady=(2,8))

        # ── Seleção de estado ──
        self._T(col_esq, "ESTADO DO CHECKBOX A CALIBRAR")
        self.var_estado = tk.IntVar(value=1)
        for n, (lbl, cor) in ESTADOS.items():
            tk.Radiobutton(col_esq, text=f"  {n}. {lbl}",
                           variable=self.var_estado, value=n,
                           bg=BG, fg=cor, selectcolor=CAMPO,
                           activebackground=BG, activeforeground=cor,
                           font=("Segoe UI",9,"bold"),
                           command=self._trocar_estado,
                           ).pack(anchor="w", pady=2)

        # ── Campos do checkbox (modo imagem) ──
        self._T(col_esq, "POSIÇÃO DO CHECKBOX (modo imagem)")
        self.lbl_estado = tk.Label(col_esq,
                                   text=f"Estado 1 — {ESTADOS[1][0]}",
                                   bg=BG, fg=ACENTO, font=("Segoe UI",8,"bold"))
        self.lbl_estado.pack(anchor="w", pady=(0,6))
        self._C(col_esq, "Célula âncora", "chk_ancora", DEFAULT["chk1_ancora"])
        self._S(col_esq, "Offset X (pt)", "chk_off_x",  DEFAULT["chk1_off_x"], -200, 200)
        self._S(col_esq, "Offset Y (pt)", "chk_off_y",  DEFAULT["chk1_off_y"], -200, 200)
        self._S(col_esq, "Largura (pt)",  "chk_larg",   DEFAULT["chk1_larg"],   2,  200)
        self._S(col_esq, "Altura (pt)",   "chk_alt",    DEFAULT["chk1_alt"],    2,  200)

        # ── Método ──
        self._T(col_esq, "MÉTODO")
        self.var_modo = tk.StringVar(value="nativo")
        for txt, val in [("Nativo — pinta shape XML (recomendado)", "nativo"),
                          ("Imagem — sobrepõe quadrado preto",       "imagem")]:
            tk.Radiobutton(col_esq, text=txt, variable=self.var_modo, value=val,
                           bg=BG, fg=TEXTO, selectcolor=CAMPO,
                           activebackground=BG, activeforeground=TEXTO,
                           font=("Segoe UI",9)).pack(anchor="w", pady=2)

        # ── Log ──
        self._T(col_dir, "LOG")
        fr_log = tk.Frame(col_dir, bg=LOG_BG)
        fr_log.pack(fill="both", expand=True, pady=4)
        sb_log = tk.Scrollbar(fr_log)
        sb_log.pack(side="right", fill="y")
        self.txt_log = tk.Text(fr_log, bg=LOG_BG, fg=LOG_FG,
                               font=("Consolas",9), relief="flat",
                               yscrollcommand=sb_log.set, wrap="word")
        self.txt_log.pack(fill="both", expand=True)
        sb_log.config(command=self.txt_log.yview)

        # ── Valores para copiar ──
        self._T(col_dir, "VALORES PARA COPIAR")
        self.txt_copy = tk.Text(col_dir, height=16, bg=CAMPO, fg=VERDE,
                                font=("Consolas",8), relief="flat")
        self.txt_copy.pack(fill="x", pady=4)
        tk.Button(col_dir, text="📋  COPIAR TODOS OS VALORES",
                  command=self._copiar,
                  bg=BORDA, fg=TEXTO, relief="flat",
                  font=("Segoe UI",9,"bold"), pady=5).pack(fill="x")

        self._atualizar_copy()
        for v in self._vars.values():
            v.trace_add("write", lambda *_: self.after(150, self._atualizar_copy))

    # ── Helpers de widget ─────────────────────

    def _T(self, p, t):
        tk.Label(p, text=t, font=("Segoe UI",10,"bold"),
                 fg=TEXTO, bg=BG).pack(anchor="w", pady=(12,4))

    def _arq(self, p, var, hint, ft):
        tk.Label(p, text=hint, fg=TEXTO2, bg=BG,
                 font=("Segoe UI",8)).pack(anchor="w")
        fr = tk.Frame(p, bg=BG)
        fr.pack(fill="x", pady=(0,4))
        tk.Entry(fr, textvariable=var, bg=CAMPO, fg=TEXTO,
                 insertbackground=TEXTO, relief="flat",
                 font=("Segoe UI",9)).pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="📁",
                  command=lambda: var.set(
                      filedialog.askopenfilename(filetypes=ft) or var.get()),
                  bg=ACENTO, fg=TEXTO, relief="flat",
                  font=("Segoe UI",9)).pack(side="right", padx=(4,0))

    def _C(self, p, label, key, default):
        fr = tk.Frame(p, bg=BG)
        fr.pack(fill="x", pady=2)
        tk.Label(fr, text=label, width=16, anchor="w",
                 fg=TEXTO2, bg=BG, font=("Segoe UI",9)).pack(side="left")
        var = tk.StringVar(value=str(default))
        self._vars[key] = var
        tk.Entry(fr, textvariable=var, width=10, bg=CAMPO, fg=TEXTO,
                 insertbackground=TEXTO, relief="flat",
                 font=("Segoe UI",10,"bold")).pack(side="left")

    def _S(self, p, label, key, default, mn, mx):
        fr = tk.Frame(p, bg=BG)
        fr.pack(fill="x", pady=2)
        tk.Label(fr, text=label, width=16, anchor="w",
                 fg=TEXTO2, bg=BG, font=("Segoe UI",9)).pack(side="left")
        var = tk.StringVar(value=str(default))
        self._vars[key] = var
        tk.Spinbox(fr, textvariable=var, from_=mn, to=mx, width=7,
                   bg=CAMPO, fg=TEXTO, buttonbackground=BORDA,
                   relief="flat", insertbackground=TEXTO,
                   font=("Segoe UI",10,"bold")).pack(side="left")
        tk.Button(fr, text="-5",
                  command=lambda v=var: self._nudge(v,-5),
                  bg=BORDA, fg=TEXTO2, relief="flat",
                  font=("Segoe UI",8), padx=4).pack(side="left", padx=(6,1))
        tk.Button(fr, text="+5",
                  command=lambda v=var: self._nudge(v,+5),
                  bg=BORDA, fg=TEXTO2, relief="flat",
                  font=("Segoe UI",8), padx=4).pack(side="left", padx=1)

    def _nudge(self, var, d):
        try: var.set(str(int(var.get())+d))
        except: pass

    # ── Lógica ────────────────────────────────

    def _salvar_estado_atual(self):
        """Salva os valores dos campos chk_* no DEFAULT para o estado atual."""
        n = self._estado_atual
        pfx = f"chk{n}_"
        v = self._vars
        try:
            DEFAULT[pfx+"ancora"] = v["chk_ancora"].get().strip()
            DEFAULT[pfx+"off_x"]  = int(v["chk_off_x"].get())
            DEFAULT[pfx+"off_y"]  = int(v["chk_off_y"].get())
            DEFAULT[pfx+"larg"]   = int(v["chk_larg"].get())
            DEFAULT[pfx+"alt"]    = int(v["chk_alt"].get())
        except: pass

    def _carregar_estado(self, n):
        """Carrega os valores do DEFAULT para os campos chk_*."""
        pfx = f"chk{n}_"
        v = self._vars
        v["chk_ancora"].set(DEFAULT[pfx+"ancora"])
        v["chk_off_x"].set(str(DEFAULT[pfx+"off_x"]))
        v["chk_off_y"].set(str(DEFAULT[pfx+"off_y"]))
        v["chk_larg"].set(str(DEFAULT[pfx+"larg"]))
        v["chk_alt"].set(str(DEFAULT[pfx+"alt"]))

    def _trocar_estado(self):
        self._salvar_estado_atual()
        n = self.var_estado.get()
        self._estado_atual = n
        lbl, cor = ESTADOS[n]
        self.lbl_estado.configure(text=f"Estado {n} — {lbl}", fg=cor)
        self._carregar_estado(n)
        self._atualizar_copy()

    def _cfg(self):
        self._salvar_estado_atual()
        v = self._vars
        s = lambda k: v[k].get().strip()
        i = lambda k: int(v[k].get())
        return {
            "ass_ancora": s("ass_ancora"),
            "ass_off_x":  i("ass_off_x"),
            "ass_off_y":  i("ass_off_y"),
            "ass_larg":   i("ass_larg"),
            "ass_alt":    i("ass_alt"),
            **{k: v for k, v in DEFAULT.items() if k.startswith("chk")},
        }

    def _atualizar_copy(self):
        try:
            self._salvar_estado_atual()
        except: pass
        v = self._vars
        try:
            txt = (
                f'ASSINATURA_EXCEL_ANCORA      = "{v["ass_ancora"].get()}"\n'
                f'ASSINATURA_EXCEL_OFFSET_X_PT = {v["ass_off_x"].get()}\n'
                f'ASSINATURA_EXCEL_OFFSET_Y_PT = {v["ass_off_y"].get()}\n'
                f'ASSINATURA_EXCEL_LARGURA_PT  = {v["ass_larg"].get()}\n'
                f'ASSINATURA_EXCEL_ALTURA_PT   = {v["ass_alt"].get()}\n\n'
            )
            for n in range(1, 5):
                p = f"chk{n}_"
                lbl = ESTADOS[n][0]
                txt += (
                    f'# {n}. {lbl}\n'
                    f'CHK{n}_ANCORA  = "{DEFAULT[p+"ancora"]}"\n'
                    f'CHK{n}_OFF_X   = {DEFAULT[p+"off_x"]}\n'
                    f'CHK{n}_OFF_Y   = {DEFAULT[p+"off_y"]}\n'
                    f'CHK{n}_LARGURA = {DEFAULT[p+"larg"]}\n'
                    f'CHK{n}_ALTURA  = {DEFAULT[p+"alt"]}\n\n'
                )
            self.txt_copy.delete("1.0","end")
            self.txt_copy.insert("1.0", txt)
        except: pass

    def _copiar(self):
        self.clipboard_clear()
        self.clipboard_append(self.txt_copy.get("1.0","end").strip())
        self.var_status.set("✓ Valores copiados!")

    def _iniciar(self):
        memorial = self.var_memorial.get().strip()
        if not memorial or not os.path.exists(memorial):
            messagebox.showerror("Erro","Selecione um Memorial válido.")
            return
        if "PREVIEW_MEMORIAL_CALIBRADOR" in os.path.basename(memorial):
            messagebox.showerror("Arquivo inválido",
                "Selecione o memorial ORIGINAL, não o arquivo de preview.")
            return
        try:
            cfg = self._cfg()
        except ValueError as e:
            messagebox.showerror("Valor inválido", str(e))
            return

        ass   = self.var_assinatura.get().strip() or None
        estado = self.var_estado.get()
        modo   = self.var_modo.get()
        saida  = str(Path(memorial).parent / "PREVIEW_MEMORIAL_CALIBRADOR.xlsx")

        self.btn.configure(state="disabled", text="⏳  Gerando preview...")
        self.txt_log.delete("1.0","end")
        self.var_status.set("Processando...")

        threading.Thread(
            target=self._worker,
            args=(memorial, ass, cfg, estado, modo, saida),
            daemon=True,
        ).start()

    def _worker(self, memorial, ass, cfg, estado, modo, saida):
        def log(msg):
            self.after(0, self._log_insert, msg)
        try:
            gerar_preview(memorial, ass, cfg, estado, modo, saida, log)
            self.after(0, self.var_status.set, "✅ Preview gerado — verifique o PDF!")
        except Exception as e:
            self.after(0, self.var_status.set, f"✗ {e}")
            self.after(0, self._log_insert, f"✗ ERRO:\n{traceback.format_exc()}")
            try:
                import subprocess
                subprocess.run(["taskkill","/F","/IM","EXCEL.EXE"],
                               capture_output=True, creationflags=0x08000000)
            except: pass
        finally:
            self.after(0, self.btn.configure,
                       {"state":"normal",
                        "text":"⚡  GERAR PREVIEW (abre PDF automaticamente)"})

    def _log_insert(self, msg):
        self.txt_log.insert("end", msg+"\n")
        self.txt_log.see("end")


if __name__ == "__main__":
    Calibrador().mainloop()

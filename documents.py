"""
comparar_tiquetes_gui.py  —  Herramienta de cruce COMFACHOCO
═══════════════════════════════════════════════════════════════
PROPÓSITO (reutilizable para cualquier mes):

Cruza dos archivos Excel periódicos de COMFACHOCO:

  Archivo MES (ej. FEBRERO_2026_PARA_LAINER.xlsx)
    Col A: Nro orden de compra   <- clave de comparación
    Col B: NUMERO DEL TIQUETE    <- se llena con este proceso
    Col C: FECHA                 <- se llena con este proceso
    Col D: HORA                  <- se llena con este proceso
    Col J: Número de Documento   <- clave de comparación (cédula)
    Col N: Cantidad de Pasajes   <- controla la expansión de filas

  Archivo TIQUETES (ej. TIQUETES_EXPEDIDIOS_EN_FEBRERO_2026.xls)
    Col K: CEDULA PASAJERO       <- clave de comparación con col J de MES
    Col M: FEC.SALIDA            <- fecha + hora del viaje (se separa)
    Col N: NRO ORDEN CREDITO     <- clave de comparación con col A de MES

ESTRATEGIA DE COMPARACIÓN:
  Coincide si: CEDULA (K==J)  Y  orden compatible (N~=A):
    1. Match exacto:  str(orden_tiq) == str(orden_mes)
    2. Match prefijo: str(orden_tiq).startswith(str(orden_mes))
       (NRO ORDEN CREDITO puede tener dígitos extra al final,
        ej. FEBRERO=2026017179 <-> TIQUETES=20260171799)

  NO se usa matching numérico aproximado porque órdenes de ida y vuelta
  del mismo usuario tienen números consecutivos (ej. 2026026318 y 2026026328)
  y mezclarlas sería incorrecto.

NOTAS:
  - Cada tiquete se asigna una sola vez (no se repite).
  - Los tiquetes se asignan en orden cronológico (FEC.SALIDA).
"""

import subprocess, os, shutil, threading
import pandas as pd
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


# ══════════════════════════════════════════════════════════════════
#  UTILIDADES
# ══════════════════════════════════════════════════════════════════

def buscar_libreoffice():
    import shutil as sh
    found = sh.which("libreoffice") or sh.which("soffice")
    if found:
        return found
    for r in [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]:
        if os.path.isfile(r):
            return r
    return None


def convertir_xls(ruta_xls, ejecutable, log_fn):
    """
    Convierte .xls a .xlsx usando LibreOffice.

    CORRECCIONES vs versión anterior:
      1. Invalida la caché si el .xls es más nuevo que el _converted.xlsx.
      2. Busca el archivo de salida en múltiples directorios porque
         LibreOffice a veces ignora --outdir y escribe en el directorio
         de trabajo actual o en la raíz del sistema de archivos.
    """
    carpeta    = os.path.dirname(os.path.abspath(ruta_xls))
    base       = os.path.splitext(os.path.basename(ruta_xls))[0]
    destino    = os.path.join(carpeta, base + "_converted.xlsx")

    # ── CORRECCIÓN 1: invalidar caché si el xls es más reciente ──
    if os.path.exists(destino):
        mtime_xls     = os.path.getmtime(ruta_xls)
        mtime_destino = os.path.getmtime(destino)
        if mtime_xls > mtime_destino:
            log_fn("INFO", "  Caché de conversión obsoleta, reconvirtiendo...")
            os.remove(destino)
        else:
            log_fn("INFO", "  Usando conversión en caché...")
            return destino

    # ── Ejecutar LibreOffice ──────────────────────────────────────
    cwd = os.getcwd()
    result = subprocess.run(
        [ejecutable, "--headless", "--convert-to", "xlsx",
         os.path.abspath(ruta_xls), "--outdir", carpeta],
        capture_output=True, text=True,
        cwd=carpeta   # forzar directorio de trabajo = carpeta del archivo
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice no pudo convertir el archivo.\n{result.stderr}")

    # ── CORRECCIÓN 2: buscar el archivo en múltiples ubicaciones ──
    nombre_xlsx = base + ".xlsx"
    posibles    = [
        os.path.join(carpeta, nombre_xlsx),          # lo normal
        os.path.join(cwd, nombre_xlsx),              # directorio de trabajo
        os.path.join("/", nombre_xlsx),              # raíz (Linux/Mac)
        os.path.join(os.path.expanduser("~"), nombre_xlsx),  # home
    ]

    encontrado = None
    for p in posibles:
        if os.path.exists(p):
            encontrado = p
            break

    if encontrado is None:
        raise RuntimeError(
            f"LibreOffice completó sin errores pero no se encontró el archivo "
            f"convertido '{nombre_xlsx}'.\n"
            f"Se buscó en: {posibles}\n"
            f"Salida de LibreOffice: {result.stdout}"
        )

    if encontrado != destino:
        shutil.move(encontrado, destino)

    return destino


def leer_tiquetes(ruta, log_fn):
    """Lee TIQUETES EXPEDIDOS (.xls o .xlsx). El encabezado real está en la fila 6."""
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".xlsx":
        df = pd.read_excel(ruta, header=5)
        df.columns = df.columns.str.strip()
        return df

    lo = buscar_libreoffice()
    if lo:
        log_fn("INFO", "  Convirtiendo .xls con LibreOffice...")
        ruta_conv = convertir_xls(ruta, lo, log_fn)
        df = pd.read_excel(ruta_conv, header=5)
        df.columns = df.columns.str.strip()
        # Verificar que el archivo tiene datos reales
        if df.empty or "CEDULA PASAJERO" not in df.columns:
            raise RuntimeError(
                "El archivo de tiquetes convertido no tiene la estructura esperada.\n"
                f"Columnas encontradas: {list(df.columns)}\n"
                "Asegúrese de que el archivo .xls es el correcto."
            )
        return df

    log_fn("INFO", "  Usando xlrd para leer .xls...")
    try:
        import xlrd  # noqa
    except ImportError:
        raise ImportError(
            "Para leer .xls instala xlrd:  pip install xlrd\n"
            "O instala LibreOffice (gratis)."
        )
    df_raw = pd.read_excel(ruta, header=None, engine="xlrd")
    header_row = next(
        i for i, row in df_raw.iterrows()
        if any(str(v).strip().upper() == "TIQUETE" for v in row)
    )
    df_raw.columns = df_raw.iloc[header_row].str.strip()
    return df_raw.iloc[header_row + 1:].reset_index(drop=True)


def to_str_orden(v):
    """Convierte un número de orden a string entero limpio."""
    try:
        return str(int(float(v))) if v is not None and not pd.isna(v) else ""
    except (ValueError, TypeError):
        return str(v).strip() if v else ""


def to_str_doc(v):
    """Convierte un número de documento/cédula a string entero limpio."""
    try:
        return str(int(float(v))) if v is not None and not pd.isna(v) else ""
    except (ValueError, TypeError):
        return str(v).strip() if v else ""


def separar_fecha_hora(valor):
    """Separa FEC.SALIDA en (fecha 'DD/MM/YYYY', hora 'HH:MM:SS')."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return None, None
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y"), valor.strftime("%H:%M:%S")
    try:
        dt = datetime.strptime(str(valor).strip(), "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S")
    except ValueError:
        return str(valor), None


def orden_coincide(o_mes, o_tiq):
    """
    True si o_mes y o_tiq se consideran la misma orden:
      - Exacto: o_tiq == o_mes
      - Prefijo: o_tiq.startswith(o_mes)  (dígitos extra en TIQUETES)
    """
    if not o_mes or not o_tiq:
        return False
    return o_tiq == o_mes or o_tiq.startswith(o_mes)


# ══════════════════════════════════════════════════════════════════
#  NÚCLEO DEL PROCESO
# ══════════════════════════════════════════════════════════════════

_wb_orig_ws = None   # referencia global al ws original (para copiar estilos)


def procesar(ruta_mes, ruta_tiquetes, log_fn, progress_fn):
    global _wb_orig_ws

    # ── Leer archivos ──────────────────────────────────────────────
    progress_fn(5)
    log_fn("INFO", "Leyendo archivo MES...")
    df_mes    = pd.read_excel(ruta_mes)
    wb_orig   = load_workbook(ruta_mes)
    _wb_orig_ws = wb_orig.active

    progress_fn(15)
    log_fn("INFO", "Leyendo archivo TIQUETES EXPEDIDOS...")
    df_tiq = leer_tiquetes(ruta_tiquetes, log_fn)
    log_fn("INFO", f"MES: {len(df_mes)} filas  |  TIQUETES: {len(df_tiq)} filas")

    # ── Validación temprana ────────────────────────────────────────
    cols_tiq = set(df_tiq.columns.tolist())
    for col_req in ("TIQUETE", "CEDULA PASAJERO", "FEC.SALIDA", "NRO ORDEN CREDITO"):
        if col_req not in cols_tiq:
            raise RuntimeError(
                f"No se encontró la columna '{col_req}' en el archivo de tiquetes.\n"
                f"Columnas disponibles: {sorted(cols_tiq)}"
            )

    # ── Paso 1: índice de tiquetes con fecha/hora ya separadas ────
    progress_fn(28)
    log_fn("INFO", "Paso 1 — Separando col M (FEC.SALIDA) en FECHA y HORA...")

    tiq_index = {}   # str_cedula -> [{ orden_str, tiquete, fecha, hora, raw_fec }]
    filas_sin_cedula = 0
    for _, row in df_tiq.iterrows():
        o_str = to_str_orden(row.get("NRO ORDEN CREDITO"))   # col N de TIQUETES
        d_str = to_str_doc(row.get("CEDULA PASAJERO"))       # col K de TIQUETES
        if not o_str or not d_str or o_str == "0" or d_str == "0":
            filas_sin_cedula += 1
            continue
        fecha, hora = separar_fecha_hora(row.get("FEC.SALIDA"))
        tiq_index.setdefault(d_str, []).append({
            "orden_str": o_str,
            "tiquete":   str(row["TIQUETE"]).strip(),
            "fecha":     fecha,
            "hora":      hora,
            "raw_fec":   row.get("FEC.SALIDA"),
        })

    for d in tiq_index:   # ordenar cronológicamente
        tiq_index[d].sort(key=lambda x: str(x["raw_fec"]) if x["raw_fec"] else "")

    log_fn("INFO", f"  Cédulas únicas en TIQUETES: {len(tiq_index)}")
    log_fn("INFO", f"  Filas de tiquetes sin cédula/orden válida: {filas_sin_cedula}")

    # ── Pasos 2 & 3: planificar expansión y asignación ────────────
    progress_fn(40)
    log_fn("INFO", "Paso 2 — Planificando expansión por cantidad de pasajes...")
    log_fn("INFO", "Paso 3 — Comparando col K+N (TIQUETES) con col J+A (MES)...")

    col_map_orig  = {_wb_orig_ws.cell(1, c).value: c
                     for c in range(1, _wb_orig_ws.max_column + 1)}
    col_cant_orig = col_map_orig.get("Cantidad de Pasajes", 14)

    plan = []
    filas_ok = filas_no = 0

    for df_idx in df_mes.index:
        excel_row_orig = df_idx + 2
        o_mes_str = to_str_orden(df_mes.loc[df_idx, "Nro orden de compra"])  # col A
        d_mes_str = to_str_doc(df_mes.loc[df_idx, "Número de Documento"])    # col J

        val_cant = _wb_orig_ws.cell(excel_row_orig, col_cant_orig).value
        try:
            cantidad = int(val_cant) if val_cant and int(val_cant) > 0 else 1
        except (TypeError, ValueError):
            cantidad = 1

        candidatos = [
            e for e in tiq_index.get(d_mes_str, [])
            if orden_coincide(o_mes_str, e["orden_str"])
        ]

        plan.append((excel_row_orig, cantidad, candidatos))
        if candidatos:
            filas_ok += 1
        else:
            filas_no += 1

    log_fn("OK",   f"Filas del MES con tiquetes encontrados: {filas_ok}")
    log_fn("WARN", f"Filas del MES sin tiquetes: {filas_no}")

    # ── Preparar archivo de salida ────────────────────────────────
    progress_fn(55)
    log_fn("INFO", "Copiando formato del archivo MES original...")

    ruta_salida = os.path.join(
        os.path.dirname(os.path.abspath(ruta_mes)),
        os.path.splitext(os.path.basename(ruta_mes))[0] + "_ACTUALIZADO.xlsx"
    )
    shutil.copy2(ruta_mes, ruta_salida)

    wb = load_workbook(ruta_salida)
    ws = wb.active

    col_map  = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    col_tiq  = col_map["NUMERO DEL TIQUETE"]
    col_fec  = col_map["FECHA"]
    col_hora = col_map["HORA"]
    col_cant = col_map["Cantidad de Pasajes"]
    col_tar  = col_map["Tarifa"]
    col_tot  = col_map["TOTAL"]
    n_cols   = ws.max_column
    amarillo = PatternFill("solid", fgColor="FFFF00")

    # ── Expandir filas y escribir B, C, D ─────────────────────────
    progress_fn(68)
    log_fn("INFO", "Expandiendo filas y escribiendo datos en col B, C, D...")

    ws.delete_rows(2, ws.max_row - 1)

    tiquetes_usados  = set()
    filas_expandidas = 0
    write_row        = 2

    for (orig_row, cantidad, candidatos) in plan:
        disponibles = [e for e in candidatos if e["tiquete"] not in tiquetes_usados]

        for _rep in range(cantidad):
            ws.append([None] * n_cols)

            for c in range(1, n_cols + 1):
                src = _wb_orig_ws.cell(orig_row, c)
                dst = ws.cell(write_row, c)
                if src.has_style:
                    dst.font          = copy(src.font)
                    dst.fill          = copy(src.fill)
                    dst.border        = copy(src.border)
                    dst.alignment     = copy(src.alignment)
                    dst.number_format = src.number_format
                if c == col_cant:
                    dst.value = 1
                elif c == col_tot:
                    dst.value = _wb_orig_ws.cell(orig_row, col_tar).value or 0
                elif c in (col_tiq, col_fec, col_hora):
                    dst.value = None   # limpiar valores previos
                else:
                    dst.value = src.value

            if disponibles:
                entry = disponibles.pop(0)
                tiquetes_usados.add(entry["tiquete"])
                ws.cell(write_row, col_tiq).value = entry["tiquete"]
                ws.cell(write_row, col_fec).value  = entry["fecha"]
                ws.cell(write_row, col_hora).value = entry["hora"]
                for col in [col_tiq, col_fec, col_hora]:
                    ws.cell(write_row, col).fill = amarillo

            if cantidad > 1:
                filas_expandidas += 1
            write_row += 1

    progress_fn(90)
    log_fn("INFO", "Guardando archivo actualizado...")
    wb.save(ruta_salida)
    progress_fn(100)

    total_filas   = write_row - 2
    tiq_asignados = len(tiquetes_usados)
    log_fn("OK", f"Tiquetes asignados en col B: {tiq_asignados}")
    log_fn("OK", f"Filas expandidas (pasajes > 1): {filas_expandidas}")
    log_fn("OK", f"Total filas en archivo final: {total_filas}")

    return ruta_salida, filas_ok, filas_no, tiq_asignados, total_filas


# ══════════════════════════════════════════════════════════════════
#  PALETA Y TIPOGRAFÍA
# ══════════════════════════════════════════════════════════════════

C = {
    "bg":            "#F7F6F3",
    "surface":       "#FFFFFF",
    "surface2":      "#F0EEE9",
    "border":        "#E2DED6",
    "border_focus":  "#1A1A1A",
    "text":          "#1A1A1A",
    "text2":         "#6B6860",
    "text3":         "#A09D97",
    "accent":        "#1A1A1A",
    "accent_hover":  "#333333",
    "ok":            "#1B7A4A",
    "ok_bg":         "#EBF7F1",
    "warn":          "#9A6500",
    "warn_bg":       "#FFF8E6",
    "err":           "#C0392B",
    "err_bg":        "#FDECEA",
    "info":          "#1A4A7A",
    "info_bg":       "#EBF2FB",
    "progress_bg":   "#E8E6E1",
    "progress_fill": "#1A1A1A",
}

FT = {
    "title": ("Georgia", 18, "bold"),
    "sub":   ("Georgia", 10, "italic"),
    "label": ("Helvetica", 9, "bold"),
    "body":  ("Helvetica", 10),
    "small": ("Helvetica", 8),
    "log":   ("Courier", 9),
    "btn":   ("Helvetica", 11, "bold"),
    "stat":  ("Georgia", 22, "bold"),
    "statl": ("Helvetica", 8),
}


# ══════════════════════════════════════════════════════════════════
#  COMPONENTES UI
# ══════════════════════════════════════════════════════════════════

class FileCard(tk.Frame):
    def __init__(self, parent, number, label, extensions, **kw):
        super().__init__(parent, bg=C["surface"],
                         highlightbackground=C["border"], highlightthickness=1, **kw)
        self.var = tk.StringVar()

        nf = tk.Frame(self, bg=C["accent"], width=32, height=32)
        nf.pack_propagate(False)
        nf.pack(side="left", padx=(16, 12), pady=16)
        tk.Label(nf, text=str(number), font=("Helvetica", 11, "bold"),
                 fg=C["surface"], bg=C["accent"]).place(relx=0.5, rely=0.5, anchor="center")

        center = tk.Frame(self, bg=C["surface"])
        center.pack(side="left", fill="x", expand=True, pady=12)
        tk.Label(center, text=label.upper(), font=FT["label"],
                 fg=C["text2"], bg=C["surface"], anchor="w").pack(fill="x")
        self.path_lbl = tk.Label(center, text="Haz clic para seleccionar archivo...",
                                 font=FT["body"], fg=C["text3"], bg=C["surface"], anchor="w")
        self.path_lbl.pack(fill="x")

        tk.Button(self, text="Examinar", font=FT["small"],
                  fg=C["text"], bg=C["surface2"],
                  activebackground=C["border"], activeforeground=C["text"],
                  relief="flat", bd=0, padx=14, pady=6, cursor="hand2",
                  command=lambda: self._browse(extensions)
                  ).pack(side="right", padx=16, pady=16)

        self.dot = tk.Label(self, text="*", font=("Helvetica", 14, "bold"),
                            fg=C["border"], bg=C["surface"])
        self.dot.place(relx=1.0, rely=0.0, x=-10, y=6, anchor="ne")

        for w in [self, center, self.path_lbl]:
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", lambda e, ext=extensions: self._browse(ext))

    def _browse(self, extensions):
        path = filedialog.askopenfilename(filetypes=[("Excel", extensions)])
        if path:
            self.var.set(path)
            self.path_lbl.configure(text=os.path.basename(path), fg=C["text"])
            self.dot.configure(fg=C["ok"])
            self.configure(highlightbackground=C["ok"])

    def _on_enter(self, e):
        if not self.var.get():
            self.configure(highlightbackground=C["border_focus"])

    def _on_leave(self, e):
        if not self.var.get():
            self.configure(highlightbackground=C["border"])

    def get(self):
        return self.var.get().strip()


class StatCard(tk.Frame):
    def __init__(self, parent, label, color_key="ok", **kw):
        bg, fg = C[color_key + "_bg"], C[color_key]
        super().__init__(parent, bg=bg,
                         highlightbackground=C["border"], highlightthickness=1, **kw)
        self._var = tk.StringVar(value="--")
        tk.Label(self, textvariable=self._var, font=FT["stat"], fg=fg, bg=bg).pack(pady=(14, 2))
        tk.Label(self, text=label.upper(), font=FT["statl"], fg=fg, bg=bg,
                 wraplength=110, justify="center").pack(pady=(0, 14))

    def set(self, val):
        self._var.set(str(val))


class LogLine(tk.Frame):
    _ICONS  = {"OK": "+", "WARN": "!", "ERR": "X", "INFO": "."}
    _COLORS = {
        "OK":   (C["ok"],    C["ok_bg"]),
        "WARN": (C["warn"],  C["warn_bg"]),
        "ERR":  (C["err"],   C["err_bg"]),
        "INFO": (C["text2"], C["surface"]),
    }

    def __init__(self, parent, level, message, **kw):
        fg, bg = self._COLORS.get(level, (C["text2"], C["surface"]))
        super().__init__(parent, bg=bg, **kw)
        tk.Label(self, text=self._ICONS.get(level, "."),
                 font=("Helvetica", 9, "bold"), fg=fg, bg=bg, width=2
                 ).pack(side="left", padx=(8, 4), pady=3)
        tk.Label(self, text=message, font=FT["log"],
                 fg=fg if level != "INFO" else C["text2"],
                 bg=bg, anchor="w"
                 ).pack(side="left", fill="x", pady=3, padx=(0, 8))


class AnimatedButton(tk.Button):
    def __init__(self, parent, **kw):
        super().__init__(parent,
                         bg=C["accent"], fg=C["surface"],
                         activebackground=C["accent_hover"], activeforeground=C["surface"],
                         relief="flat", bd=0, cursor="hand2", font=FT["btn"], pady=14, **kw)
        self.bind("<Enter>", lambda e: self.configure(bg=C["accent_hover"]))
        self.bind("<Leave>", lambda e: self.configure(bg=C["accent"]))


class ProgressBar(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, height=4, bg=C["progress_bg"], highlightthickness=0, **kw)
        self._pct = self._target = 0.0
        self._bar = self.create_rectangle(0, 0, 0, 4, fill=C["progress_fill"], width=0)
        self.bind("<Configure>", self._draw)

    def _draw(self, e=None):
        self.coords(self._bar, 0, 0, int(self.winfo_width() * self._pct / 100), 4)

    def set_target(self, pct):
        self._target = float(pct)
        self._step()

    def _step(self):
        if abs(self._pct - self._target) < 0.5:
            self._pct = self._target
            self._draw()
            return
        self._pct += (self._target - self._pct) * 0.15
        self._draw()
        self.after(16, self._step)

    def reset(self):
        self._pct = self._target = 0.0
        self._draw()


# ══════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cruce COMFACHOCO — Tiquetes vs Autorizaciones")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self._build()

    def _build(self):
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="both", expand=True, padx=32, pady=28)

        # Encabezado
        hdr = tk.Frame(outer, bg=C["bg"])
        hdr.pack(fill="x", pady=(0, 6))
        tk.Label(hdr, text="Cruce", font=FT["title"],
                 fg=C["text"], bg=C["bg"]).pack(side="left")
        tk.Label(hdr, text=" COMFACHOCO", font=("Georgia", 18),
                 fg=C["text2"], bg=C["bg"]).pack(side="left")
        tk.Label(hdr, text="Tiquetes vs Autorizaciones",
                 font=FT["sub"], fg=C["text3"], bg=C["bg"]
                 ).pack(side="right", anchor="s", pady=4)

        tk.Label(outer,
                 text="Compara cedula (col K<->J) y orden (col N<->A) entre archivos, "
                      "y completa las columnas B, C y D del archivo MES.",
                 font=FT["small"], fg=C["text3"], bg=C["bg"],
                 anchor="w", justify="left", wraplength=480).pack(fill="x", pady=(0, 16))

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(0, 20))

        # Archivos de entrada
        tk.Label(outer, text="ARCHIVOS DE ENTRADA", font=FT["label"],
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 8))

        self.card_mes = FileCard(outer, 1,
                                 "Archivo MES - Autorizaciones COMFACHOCO",
                                 "*.xlsx *.xls")
        self.card_mes.pack(fill="x", pady=(0, 8))

        self.card_tiq = FileCard(outer, 2,
                                 "Archivo TIQUETES EXPEDIDOS",
                                 "*.xls *.xlsx")
        self.card_tiq.pack(fill="x")

        nota = tk.Frame(outer, bg=C["info_bg"],
                        highlightbackground=C["border"], highlightthickness=1)
        nota.pack(fill="x", pady=(10, 0))
        tk.Label(nota,
                 text="El archivo actualizado se generara automaticamente "
                      "en la misma carpeta que el archivo MES.",
                 font=FT["small"], fg=C["info"], bg=C["info_bg"],
                 anchor="w", pady=8, padx=12).pack(fill="x")

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(20, 16))

        # Barra de progreso
        self.progress = ProgressBar(outer, width=480)
        self.progress.pack(fill="x", pady=(0, 6))
        self.status_var = tk.StringVar(value="Listo para procesar")
        tk.Label(outer, textvariable=self.status_var,
                 font=FT["small"], fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x")

        self.btn = AnimatedButton(outer, text="Procesar archivos", command=self._run)
        self.btn.pack(fill="x", pady=(14, 0))

        # Tarjetas de resultados
        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(20, 16))
        tk.Label(outer, text="RESULTADOS", font=FT["label"],
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 10))

        sf = tk.Frame(outer, bg=C["bg"])
        sf.pack(fill="x")
        self.s_ok    = StatCard(sf, "Filas con tiquete",  "ok")
        self.s_no    = StatCard(sf, "Sin tiquete",        "warn")
        self.s_tiq   = StatCard(sf, "Tiquetes asignados", "ok")
        self.s_total = StatCard(sf, "Total filas",        "info")
        for i, s in enumerate([self.s_ok, self.s_no, self.s_tiq, self.s_total]):
            s.grid(row=0, column=i, padx=(0, 8) if i < 3 else 0, sticky="nsew")
            sf.columnconfigure(i, weight=1)

        # Registro de actividad
        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(20, 16))
        tk.Label(outer, text="REGISTRO DE ACTIVIDAD", font=FT["label"],
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 8))

        log_wrap = tk.Frame(outer, bg=C["surface"],
                            highlightbackground=C["border"], highlightthickness=1)
        log_wrap.pack(fill="both")
        self.log_inner = tk.Frame(log_wrap, bg=C["surface"])
        self.log_inner.pack(fill="both", expand=True)
        self._placeholder = tk.Label(
            self.log_inner,
            text="El registro aparecera aqui una vez iniciado el proceso.",
            font=FT["small"], fg=C["text3"], bg=C["surface"], pady=20)
        self._placeholder.pack()
        self._log_lines = []

    # ── Helpers UI ────────────────────────────────────────────────

    def _log(self, level, msg):
        if self._placeholder.winfo_ismapped():
            self._placeholder.pack_forget()
        line = LogLine(self.log_inner, level, msg)
        line.pack(fill="x")
        self._log_lines.append(line)
        self.log_inner.update_idletasks()

    def _clear_log(self):
        for w in self._log_lines:
            w.destroy()
        self._log_lines = []
        self._placeholder.pack()

    def _restaurar_btn(self):
        self.btn.configure(state="normal", text="Procesar archivos")
        self.progress.set_target(0)

    # ── Acción principal ──────────────────────────────────────────

    def _run(self):
        mes = self.card_mes.get()
        tiq = self.card_tiq.get()
        if not mes or not tiq:
            messagebox.showwarning("Archivos requeridos",
                                   "Selecciona ambos archivos antes de continuar.")
            return

        self.btn.configure(state="disabled", text="Procesando...")
        self.progress.reset()
        self._clear_log()
        for s in [self.s_ok, self.s_no, self.s_tiq, self.s_total]:
            s.set("--")

        def log_fn(level, msg):
            self.after(0, lambda lv=level, m=msg: self._log(lv, m))

        def progress_fn(pct):
            self.after(0, lambda p=pct: self.progress.set_target(p))
            self.after(0, lambda p=pct: self.status_var.set(f"Procesando... {p}%"))

        def worker():
            try:
                ruta, ok, no, tiq_n, total = procesar(mes, tiq, log_fn, progress_fn)
                ruta_c, ok_c, no_c, tiq_c, total_c = ruta, ok, no, tiq_n, total
                def done():
                    self.s_ok.set(ok_c)
                    self.s_no.set(no_c)
                    self.s_tiq.set(tiq_c)
                    self.s_total.set(total_c)
                    self.status_var.set(f"Completado — {os.path.basename(ruta_c)}")
                    self._restaurar_btn()
                    messagebox.showinfo(
                        "Proceso completado",
                        f"Archivo generado exitosamente.\n\nUbicacion:\n{ruta_c}"
                    )
                self.after(0, done)
            except Exception as exc:
                err_msg = str(exc)
                def on_err(msg=err_msg):
                    self._log("ERR", msg)
                    self.status_var.set("Error durante el proceso")
                    self._restaurar_btn()
                    messagebox.showerror("Error", msg)
                self.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
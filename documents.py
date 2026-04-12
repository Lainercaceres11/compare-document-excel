"""
comparar_tiquetes_gui.py
UI profesional con tkinter — tema claro refinado.

Estrategia de comparación:
  - Clave: (orden, cedula)
  - Match exacto: str(orden_feb) == str(orden_tiq)
  - Match prefijo: str(orden_tiq).startswith(str(orden_feb))  [NRO ORDEN CREDITO puede tener dígitos extra]
  - Cada fila expandida recibe su propio tiquete individual en orden de FEC.SALIDA
"""

import subprocess, os, shutil, threading
import pandas as pd
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ══════════════════════════════════════════════════════════════════
#  LÓGICA DE PROCESAMIENTO
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


def convertir_xls(ruta_xls, ejecutable):
    carpeta = os.path.dirname(os.path.abspath(ruta_xls))
    base    = os.path.splitext(os.path.basename(ruta_xls))[0]
    destino = os.path.join(carpeta, base + "_converted.xlsx")
    if os.path.exists(destino):
        return destino
    result = subprocess.run(
        [ejecutable, "--headless", "--convert-to", "xlsx",
         os.path.abspath(ruta_xls), "--outdir", carpeta],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice no pudo convertir el archivo.\n{result.stderr}")
    convertido = os.path.join(carpeta, base + ".xlsx")
    if os.path.exists(convertido) and convertido != destino:
        os.rename(convertido, destino)
    return destino


def leer_tiquetes(ruta, log_fn):
    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".xlsx":
        df = pd.read_excel(ruta, header=5)
        df.columns = df.columns.str.strip()
        return df
    lo = buscar_libreoffice()
    if lo:
        log_fn("INFO", "  Convirtiendo .xls con LibreOffice...")
        ruta = convertir_xls(ruta, lo)
        df = pd.read_excel(ruta, header=5)
        df.columns = df.columns.str.strip()
        return df
    log_fn("INFO", "  LibreOffice no encontrado, usando xlrd...")
    try:
        import xlrd  # noqa
    except ImportError:
        raise ImportError("Instala xlrd: pip install xlrd\nO instala LibreOffice.")
    df_raw = pd.read_excel(ruta, header=None, engine="xlrd")
    header_row = next(
        i for i, row in df_raw.iterrows()
        if any(str(v).strip().upper() == "TIQUETE" for v in row)
    )
    df_raw.columns = df_raw.iloc[header_row].str.strip()
    return df_raw.iloc[header_row + 1:].reset_index(drop=True)


def extraer_fecha_hora(valor):
    if pd.isna(valor):
        return None, None
    dt = valor if isinstance(valor, datetime) else datetime.strptime(str(valor).strip(), "%Y-%m-%d %H:%M:%S")
    return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S")


def str_orden(v):
    """Convierte un valor de orden a string limpio para comparación."""
    try:
        return str(int(float(v))) if v and not pd.isna(v) else ""
    except (ValueError, TypeError):
        return str(v).strip() if v else ""


def str_doc(v):
    try:
        return str(int(float(v))) if v and not pd.isna(v) else ""
    except (ValueError, TypeError):
        return str(v).strip() if v else ""


wb_orig_ws = None


def procesar(route_mes, ruta_tiquetes, log_fn, progress_fn):
    global wb_orig_ws

    progress_fn(5)
    log_fn("INFO", "Leyendo archivo MES 2026...")
    df_feb     = pd.read_excel(route_mes)
    wb_orig    = load_workbook(route_mes)
    wb_orig_ws = wb_orig.active

    progress_fn(15)
    log_fn("INFO", "Leyendo archivo TIQUETES...")
    df_tiq = leer_tiquetes(ruta_tiquetes, log_fn)
    log_fn("INFO", f"MES: {len(df_feb)} filas  |  TIQUETES: {len(df_tiq)} filas")

    progress_fn(30)
    log_fn("INFO", "Construyendo índice de tiquetes...")

    # ── Índice de tiquetes: (str_orden_tiq, str_doc) -> lista ordenada por FEC.SALIDA ──
    # Normalizamos la cédula exacta y el número de orden como string
    # Para el match: buscamos (str_orden_feb == str_orden_tiq) OR str_orden_tiq.startswith(str_orden_feb)
    tiq_index = {}   # str_doc -> lista de dicts con orden_str, tiquete, fecha, hora, raw_fec
    for _, row in df_tiq.iterrows():
        o_str = str_orden(row.get("NRO ORDEN CREDITO"))
        d_str = str_doc(row.get("CEDULA PASAJERO"))
        if not o_str or not d_str:
            continue
        f, h = extraer_fecha_hora(row.get("FEC.SALIDA"))
        entry = {
            "orden_str": o_str,
            "tiquete":   row["TIQUETE"],
            "fecha":     f,
            "hora":      h,
            "raw_fec":   row.get("FEC.SALIDA"),
        }
        tiq_index.setdefault(d_str, []).append(entry)

    # Ordenar cada lista por FEC.SALIDA
    for d_str in tiq_index:
        tiq_index[d_str].sort(key=lambda x: str(x["raw_fec"]) if x["raw_fec"] else "")

    progress_fn(40)
    log_fn("INFO", "Comparando registros...")

    col_map_feb = {}
    for c in range(1, wb_orig_ws.max_column + 1):
        col_map_feb[wb_orig_ws.cell(1, c).value] = c
    col_cant_orig = col_map_feb.get("Cantidad de Pasajes", 14)

    # Para cada fila de FEBRERO, buscar tiquetes coincidentes
    # Match: misma cédula + (orden exacta O orden TIQUETES empieza con orden FEBRERO)
    plan = []
    filas_ok = filas_no = 0

    for df_idx in df_feb.index:
        excel_row_orig = df_idx + 2
        o_feb = str_orden(df_feb.loc[df_idx, "Nro orden de compra"])
        d_feb = str_doc(df_feb.loc[df_idx, "Número de Documento"])

        val_cant = wb_orig_ws.cell(excel_row_orig, col_cant_orig).value
        try:
            cantidad = int(val_cant) if val_cant and int(val_cant) > 0 else 1
        except (TypeError, ValueError):
            cantidad = 1

        # Buscar tiquetes para esta cédula que coincidan con la orden
        candidatos = []
        for entry in tiq_index.get(d_feb, []):
            o_tiq = entry["orden_str"]
            if o_tiq == o_feb or o_tiq.startswith(o_feb):
                candidatos.append(entry)

        plan.append((excel_row_orig, cantidad, candidatos))
        if candidatos:
            filas_ok += 1
        else:
            filas_no += 1

    log_fn("OK",   f"Filas con tiquetes encontrados: {filas_ok}")
    log_fn("WARN", f"Filas sin tiquetes: {filas_no}")

    progress_fn(55)
    log_fn("INFO", "Copiando formato del archivo original...")

    ruta_salida = os.path.join(
        os.path.dirname(os.path.abspath(route_mes)),
        os.path.splitext(os.path.basename(route_mes))[0] + "_ACTUALIZADO.xlsx"
    )
    shutil.copy2(route_mes, ruta_salida)

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

    progress_fn(70)
    log_fn("INFO", "Expandiendo filas y escribiendo datos...")

    ws.delete_rows(2, ws.max_row - 1)

    # Rastrear tiquetes ya usados globalmente: un tiquete no se repite
    tiquetes_usados = set()

    filas_expandidas = 0
    write_row        = 2

    for (orig_row, cantidad, candidatos) in plan:
        # Filtrar candidatos que aún no han sido usados
        disponibles = [e for e in candidatos if e["tiquete"] not in tiquetes_usados]

        for rep in range(cantidad):
            ws.append([None] * n_cols)
            # Copiar estilo y valor de la fila original
            for c in range(1, n_cols + 1):
                src = wb_orig_ws.cell(orig_row, c)
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
                    tarifa = wb_orig_ws.cell(orig_row, col_tar).value or 0
                    dst.value = tarifa
                else:
                    dst.value = src.value

            # Asignar el siguiente tiquete disponible (sin repetir)
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
    log_fn("INFO", "Guardando archivo...")
    wb.save(ruta_salida)
    progress_fn(100)

    total_filas = write_row - 2
    log_fn("OK", f"Filas expandidas (pasajes > 1): {filas_expandidas}")
    log_fn("OK", f"Total filas en archivo final:   {total_filas}")

    return ruta_salida, filas_ok, filas_no, filas_expandidas, total_filas


# ══════════════════════════════════════════════════════════════════
#  PALETA Y TIPOGRAFÍA
# ══════════════════════════════════════════════════════════════════

C = {
    "bg":           "#F7F6F3",
    "surface":      "#FFFFFF",
    "surface2":     "#F0EEE9",
    "border":       "#E2DED6",
    "border_focus": "#1A1A1A",
    "text":         "#1A1A1A",
    "text2":        "#6B6860",
    "text3":        "#A09D97",
    "accent":       "#1A1A1A",
    "accent_hover": "#333333",
    "ok":           "#1B7A4A",
    "ok_bg":        "#EBF7F1",
    "warn":         "#9A6500",
    "warn_bg":      "#FFF8E6",
    "err":          "#C0392B",
    "err_bg":       "#FDECEA",
    "info":         "#1A4A7A",
    "info_bg":      "#EBF2FB",
    "progress_bg":  "#E8E6E1",
    "progress_fill":"#1A1A1A",
}

FONT_TITLE  = ("Georgia", 18, "bold")
FONT_SUB    = ("Georgia", 10, "italic")
FONT_LABEL  = ("Helvetica", 9, "bold")
FONT_BODY   = ("Helvetica", 10)
FONT_SMALL  = ("Helvetica", 8)
FONT_LOG    = ("Courier", 9)
FONT_BTN    = ("Helvetica", 11, "bold")
FONT_STAT   = ("Georgia", 22, "bold")
FONT_STAT_L = ("Helvetica", 8)


# ══════════════════════════════════════════════════════════════════
#  COMPONENTES UI
# ══════════════════════════════════════════════════════════════════

class FileCard(tk.Frame):
    def __init__(self, parent, number, label, extensions, **kw):
        super().__init__(parent, bg=C["surface"],
                         highlightbackground=C["border"],
                         highlightthickness=1, **kw)
        self.var = tk.StringVar()

        num_frame = tk.Frame(self, bg=C["accent"], width=32, height=32)
        num_frame.pack_propagate(False)
        num_frame.pack(side="left", padx=(16, 12), pady=16)
        tk.Label(num_frame, text=str(number), font=("Helvetica", 11, "bold"),
                 fg=C["surface"], bg=C["accent"]).place(relx=0.5, rely=0.5, anchor="center")

        center = tk.Frame(self, bg=C["surface"])
        center.pack(side="left", fill="x", expand=True, pady=12)
        tk.Label(center, text=label.upper(), font=FONT_LABEL,
                 fg=C["text2"], bg=C["surface"], anchor="w").pack(fill="x")
        self.path_label = tk.Label(center, text="Haz clic para seleccionar archivo…",
                                   font=FONT_BODY, fg=C["text3"],
                                   bg=C["surface"], anchor="w")
        self.path_label.pack(fill="x")

        self.btn = tk.Button(self, text="Examinar", font=FONT_SMALL,
                             fg=C["text"], bg=C["surface2"],
                             activebackground=C["border"],
                             activeforeground=C["text"],
                             relief="flat", bd=0, padx=14, pady=6,
                             cursor="hand2",
                             command=lambda: self._browse(extensions))
        self.btn.pack(side="right", padx=16, pady=16)

        self.dot = tk.Label(self, text="●", font=("Helvetica", 8),
                            fg=C["border"], bg=C["surface"])
        self.dot.place(relx=1.0, rely=0.0, x=-8, y=8, anchor="ne")

        for w in [self, center, self.path_label]:
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<Button-1>", lambda e, ext=extensions: self._browse(ext))

    def _browse(self, extensions):
        path = filedialog.askopenfilename(filetypes=[("Excel", extensions)])
        if path:
            self.var.set(path)
            self.path_label.configure(text=os.path.basename(path), fg=C["text"])
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
        bg = C[color_key + "_bg"]
        fg = C[color_key]
        super().__init__(parent, bg=bg,
                         highlightbackground=C["border"],
                         highlightthickness=1, **kw)
        self.value_var = tk.StringVar(value="—")
        tk.Label(self, textvariable=self.value_var,
                 font=FONT_STAT, fg=fg, bg=bg).pack(pady=(14, 2))
        tk.Label(self, text=label.upper(),
                 font=FONT_STAT_L, fg=fg, bg=bg,
                 wraplength=110, justify="center").pack(pady=(0, 14))

    def set(self, val):
        self.value_var.set(str(val))


class LogLine(tk.Frame):
    ICONS  = {"OK": "✓", "WARN": "⚠", "ERR": "✕", "INFO": "·"}
    COLORS = {
        "OK":   (C["ok"],   C["ok_bg"]),
        "WARN": (C["warn"], C["warn_bg"]),
        "ERR":  (C["err"],  C["err_bg"]),
        "INFO": (C["text2"], C["surface"]),
    }

    def __init__(self, parent, level, message, **kw):
        fg, bg = self.COLORS.get(level, (C["text2"], C["surface"]))
        super().__init__(parent, bg=bg, **kw)
        tk.Label(self, text=self.ICONS.get(level, "·"),
                 font=("Helvetica", 9, "bold"),
                 fg=fg, bg=bg, width=2).pack(side="left", padx=(8, 4), pady=3)
        tk.Label(self, text=message, font=FONT_LOG,
                 fg=fg if level != "INFO" else C["text2"],
                 bg=bg, anchor="w").pack(side="left", fill="x", pady=3, padx=(0, 8))


class AnimatedButton(tk.Button):
    def __init__(self, parent, **kw):
        super().__init__(parent,
                         bg=C["accent"], fg=C["surface"],
                         activebackground=C["accent_hover"],
                         activeforeground=C["surface"],
                         relief="flat", bd=0, cursor="hand2",
                         font=FONT_BTN, pady=14, **kw)
        self.bind("<Enter>", lambda e: self.configure(bg=C["accent_hover"]))
        self.bind("<Leave>", lambda e: self.configure(bg=C["accent"]))


class ProgressBar(tk.Canvas):
    def __init__(self, parent, **kw):
        super().__init__(parent, height=4, bg=C["progress_bg"],
                         highlightthickness=0, **kw)
        self._pct    = 0
        self._target = 0
        self._bar    = self.create_rectangle(0, 0, 0, 4,
                                             fill=C["progress_fill"], width=0)
        self.bind("<Configure>", self._redraw)

    def _redraw(self, e=None):
        w = self.winfo_width()
        self.coords(self._bar, 0, 0, int(w * self._pct / 100), 4)

    def set_target(self, pct):
        self._target = pct
        self._animate()

    def _animate(self):
        if abs(self._pct - self._target) < 0.5:
            self._pct = self._target
            self._redraw()
            return
        self._pct += (self._target - self._pct) * 0.15
        self._redraw()
        self.after(16, self._animate)

    def reset(self):
        self._pct = self._target = 0
        self._redraw()


# ══════════════════════════════════════════════════════════════════
#  VENTANA PRINCIPAL
# ══════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparador de Tiquetes")
        self.configure(bg=C["bg"])
        self.resizable(False, False)
        self._build()

    def _build(self):
        outer = tk.Frame(self, bg=C["bg"])
        outer.pack(fill="both", expand=True, padx=32, pady=28)

        header = tk.Frame(outer, bg=C["bg"])
        header.pack(fill="x", pady=(0, 24))
        tk.Label(header, text="Comparador", font=FONT_TITLE,
                 fg=C["text"], bg=C["bg"]).pack(side="left")
        tk.Label(header, text=" de Tiquetes",
                 font=("Georgia", 18), fg=C["text2"], bg=C["bg"]).pack(side="left")
        tk.Label(header, text="Febrero 2026",
                 font=FONT_SUB, fg=C["text3"], bg=C["bg"]).pack(side="right", anchor="s", pady=4)

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(0, 24))

        tk.Label(outer, text="ARCHIVOS DE ENTRADA", font=FONT_LABEL,
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 8))

        self.card_feb = FileCard(outer, 1, "Archivo MES 2026", "*.xlsx *.xls")
        self.card_feb.pack(fill="x", pady=(0, 8))
        self.card_tiq = FileCard(outer, 2, "Archivo de Tiquetes", "*.xls *.xlsx")
        self.card_tiq.pack(fill="x")

        nota = tk.Frame(outer, bg=C["info_bg"],
                        highlightbackground=C["border"], highlightthickness=1)
        nota.pack(fill="x", pady=(12, 0))
        tk.Label(nota,
                 text="El archivo actualizado se generará en la misma carpeta que el archivo MES.",
                 font=FONT_SMALL, fg=C["info"], bg=C["info_bg"],
                 anchor="w", pady=8, padx=12).pack(fill="x")

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(24, 16))

        self.progress = ProgressBar(outer, width=480)
        self.progress.pack(fill="x", pady=(0, 8))
        self.status_var = tk.StringVar(value="Listo para procesar")
        tk.Label(outer, textvariable=self.status_var,
                 font=FONT_SMALL, fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x")

        self.btn = AnimatedButton(outer, text="Procesar archivos", command=self._run)
        self.btn.pack(fill="x", pady=(16, 0))

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(24, 16))
        tk.Label(outer, text="RESULTADOS", font=FONT_LABEL,
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 10))

        stats_frame = tk.Frame(outer, bg=C["bg"])
        stats_frame.pack(fill="x")
        self.stat_ok    = StatCard(stats_frame, "Con tiquete",   "ok")
        self.stat_no    = StatCard(stats_frame, "Sin tiquete",   "warn")
        self.stat_exp   = StatCard(stats_frame, "Expandidas",    "info")
        self.stat_total = StatCard(stats_frame, "Filas totales", "info")
        for i, s in enumerate([self.stat_ok, self.stat_no, self.stat_exp, self.stat_total]):
            s.grid(row=0, column=i, padx=(0, 8) if i < 3 else 0, sticky="nsew")
            stats_frame.columnconfigure(i, weight=1)

        tk.Frame(outer, bg=C["border"], height=1).pack(fill="x", pady=(24, 16))
        tk.Label(outer, text="REGISTRO DE ACTIVIDAD", font=FONT_LABEL,
                 fg=C["text3"], bg=C["bg"], anchor="w").pack(fill="x", pady=(0, 8))

        log_container = tk.Frame(outer, bg=C["surface"],
                                 highlightbackground=C["border"], highlightthickness=1)
        log_container.pack(fill="both")
        self.log_inner = tk.Frame(log_container, bg=C["surface"])
        self.log_inner.pack(fill="both", expand=True)
        self.log_placeholder = tk.Label(
            self.log_inner,
            text="El registro de actividad aparecerá aquí una vez iniciado el proceso.",
            font=FONT_SMALL, fg=C["text3"], bg=C["surface"], pady=20)
        self.log_placeholder.pack()
        self._log_lines = []

    def _add_log(self, level, message):
        if self.log_placeholder.winfo_ismapped():
            self.log_placeholder.pack_forget()
        line = LogLine(self.log_inner, level, message)
        line.pack(fill="x")
        self._log_lines.append(line)
        self.log_inner.update_idletasks()

    def _clear_log(self):
        for w in self._log_lines:
            w.destroy()
        self._log_lines = []
        self.log_placeholder.pack()

    def _run(self):
        feb = self.card_feb.get()
        tiq = self.card_tiq.get()
        if not feb or not tiq:
            messagebox.showwarning("Archivos requeridos",
                                   "Por favor selecciona los dos archivos antes de continuar.")
            return

        self.btn.configure(state="disabled", text="Procesando…")
        self.progress.reset()
        self._clear_log()
        for s in [self.stat_ok, self.stat_no, self.stat_exp, self.stat_total]:
            s.set("—")

        def log_fn(level, msg):
            self.after(0, lambda l=level, m=msg: self._add_log(l, m))

        def progress_fn(pct):
            self.after(0, lambda p=pct: self.progress.set_target(p))
            self.after(0, lambda p=pct: self.status_var.set(f"Procesando… {p}%"))

        def worker():
            try:
                ruta, ok, no, exp, total = procesar(feb, tiq, log_fn, progress_fn)
                def done():
                    self.stat_ok.set(ok)
                    self.stat_no.set(no)
                    self.stat_exp.set(exp)
                    self.stat_total.set(total)
                    self.status_var.set(f"Completado — {os.path.basename(ruta)}")
                    self.btn.configure(state="normal", text="Procesar archivos")
                    messagebox.showinfo("Proceso completado",
                                        f"Archivo generado exitosamente.\n\nUbicación:\n{ruta}")
                self.after(0, done)
            except Exception as e:
                def on_err():
                    self._add_log("ERR", str(e))
                    self.status_var.set("Error durante el proceso")
                    self.btn.configure(state="normal", text="Procesar archivos")
                    messagebox.showerror("Error", str(e))
                self.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.progress.stop()
        self.btn.configure(state="normal", text="Procesar archivos")


if __name__ == "__main__":
    app = App()
    app.mainloop()

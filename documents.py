"""
comparar_tiquetes_gui.py
Preserva el formato completo del archivo FEBRERO original y llena
NUMERO DEL TIQUETE, FECHA y HORA cuando hay exactamente 1 tiquete por orden+cédula.
"""

import subprocess, os, shutil, threading
import pandas as pd
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ── Utilidades ─────────────────────────────────────────────────────────────────

def buscar_libreoffice():
    found = shutil.which("libreoffice") or shutil.which("soffice")
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
        log_fn("  Convirtiendo .xls con LibreOffice...")
        ruta = convertir_xls(ruta, lo)
        df = pd.read_excel(ruta, header=5)
        df.columns = df.columns.str.strip()
        return df
    log_fn("  LibreOffice no encontrado, usando xlrd...")
    try:
        import xlrd  # noqa
    except ImportError:
        raise ImportError("Instala xlrd:  pip install xlrd\nO instala LibreOffice.")
    df_raw = pd.read_excel(ruta, header=None, engine="xlrd")
    header_row = next(
        i for i, row in df_raw.iterrows()
        if any(str(v).strip().upper() == "TIQUETE" for v in row)
    )
    df_raw.columns = df_raw.iloc[header_row].str.strip()
    return df_raw.iloc[header_row + 1:].reset_index(drop=True)


def normalizar(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0).astype("int64")


def extraer_fecha_hora(valor):
    if pd.isna(valor):
        return None, None
    dt = valor if isinstance(valor, datetime) else datetime.strptime(str(valor).strip(), "%Y-%m-%d %H:%M:%S")
    return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S")


def copiar_estilo(src, dst):
    """Copia fuente, relleno, borde y alineación de src a dst."""
    if src.has_style:
        dst.font      = copy(src.font)
        dst.fill      = copy(src.fill)
        dst.border    = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


# ── Procesamiento ──────────────────────────────────────────────────────────────

def procesar(ruta_febrero, ruta_tiquetes, log_fn):
    log_fn("Leyendo archivos...")
    df_feb = pd.read_excel(ruta_febrero)
    df_tiq = leer_tiquetes(ruta_tiquetes, log_fn)
    log_fn(f"  FEBRERO:  {len(df_feb)} filas")
    log_fn(f"  TIQUETES: {len(df_tiq)} filas")

    df_feb["_orden"] = normalizar(df_feb["Nro orden de compra"])
    df_feb["_doc"]   = normalizar(df_feb["Número de Documento"])
    df_tiq["_orden"] = normalizar(df_tiq["NRO ORDEN CREDITO"])
    df_tiq["_doc"]   = normalizar(df_tiq["CEDULA PASAJERO"])

    conteo           = df_tiq.groupby(["_orden", "_doc"])["TIQUETE"].count()
    claves_unicas    = set(conteo[conteo == 1].index)
    claves_multiples = set(conteo[conteo > 1].index)
    log_fn(f"  Con 1 tiquete exacto:   {len(claves_unicas)}")
    log_fn(f"  Con múltiples tiquetes: {len(claves_multiples)}")

    mask      = df_tiq.apply(lambda r: (r["_orden"], r["_doc"]) in claves_unicas, axis=1)
    df_unicos = df_tiq[mask][["_orden", "_doc", "TIQUETE", "FEC.SALIDA"]].copy()
    df_unicos.set_index(["_orden", "_doc"], inplace=True)

    # Construir mapa fila → (tiquete, fecha, hora)
    resultados = {}
    filas_ok = filas_multi = filas_no = 0
    for idx, row in df_feb.iterrows():
        clave = (row["_orden"], row["_doc"])
        if clave in df_unicos.index:
            reg = df_unicos.loc[clave]
            f, h = extraer_fecha_hora(reg["FEC.SALIDA"])
            resultados[idx] = (reg["TIQUETE"], f, h)
            filas_ok += 1
        else:
            resultados[idx] = (None, None, None)
            if clave in claves_multiples:
                filas_multi += 1
            else:
                filas_no += 1

    # ── Escribir sobre una COPIA del original para preservar formato ───────────
    log_fn("Copiando formato del archivo original...")
    ruta_salida = os.path.join(
        os.path.dirname(os.path.abspath(ruta_febrero)),
        os.path.splitext(os.path.basename(ruta_febrero))[0] + "_ACTUALIZADO.xlsx"
    )
    shutil.copy2(ruta_febrero, ruta_salida)

    wb = load_workbook(ruta_salida)
    ws = wb.active

    # Mapear nombre de columna → índice Excel (1-based)
    col_map = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    col_tiq  = col_map.get("NUMERO DEL TIQUETE")
    col_fec  = col_map.get("FECHA")
    col_hora = col_map.get("HORA")

    # Estilo de referencia para celdas llenas (tomado de fila 2 si existe)
    ref_row = 2 if ws.max_row >= 2 else 1

    amarillo = PatternFill("solid", fgColor="FFFF00")

    log_fn("Llenando celdas...")
    for df_idx, (tiquete, fecha, hora) in resultados.items():
        excel_row = df_idx + 2  # +1 header, +1 base-1
        if tiquete is not None:
            ws.cell(excel_row, col_tiq).value = tiquete
            ws.cell(excel_row, col_fec).value = fecha
            ws.cell(excel_row, col_hora).value = hora
            # Resaltar en amarillo solo las celdas recién llenadas
            for col in [col_tiq, col_fec, col_hora]:
                ws.cell(excel_row, col).fill = amarillo

    wb.save(ruta_salida)

    log_fn("─" * 46)
    log_fn(f"✅  Filas actualizadas              : {filas_ok}")
    log_fn(f"⚠️   Múltiples tiquetes (sin llenar): {filas_multi}")
    log_fn(f"❌  Sin coincidencia                : {filas_no}")
    log_fn("─" * 46)
    log_fn(f"📄  Archivo guardado en:")
    log_fn(f"    {ruta_salida}")
    return ruta_salida, filas_ok, filas_multi, filas_no


# ── Interfaz gráfica ───────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparador de Tiquetes – FEBRERO 2026")
        self.resizable(False, False)
        self.configure(bg="#1a1a2e")
        self._build_ui()

    def _build_ui(self):
        BG    = "#1a1a2e"
        CARD  = "#16213e"
        ACC   = "#e94560"
        ACC2  = "#0f3460"
        TXT   = "#eaeaea"
        MUTED = "#8892a4"
        PAD   = 18

        hdr = tk.Frame(self, bg=ACC2, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="COMPARADOR DE TIQUETES",
                 font=("Courier New", 15, "bold"), fg=ACC, bg=ACC2).pack()
        tk.Label(hdr, text="Febrero 2026  •  Actualización automática",
                 font=("Courier New", 9), fg=MUTED, bg=ACC2).pack()

        body = tk.Frame(self, bg=BG, padx=PAD, pady=PAD)
        body.pack(fill="both")

        def file_row(parent, label, var, tipos):
            row = tk.Frame(parent, bg=CARD, padx=12, pady=10)
            row.pack(fill="x", pady=5)
            tk.Label(row, text=label, font=("Courier New", 8, "bold"),
                     fg=MUTED, bg=CARD, anchor="w").pack(fill="x")
            inner = tk.Frame(row, bg=CARD)
            inner.pack(fill="x")
            e = tk.Entry(inner, textvariable=var, font=("Courier New", 9),
                         bg="#0d1b2a", fg=TXT, insertbackground=TXT,
                         relief="flat", bd=0, state="readonly",
                         readonlybackground="#0d1b2a")
            e.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 8))
            tk.Button(inner, text="EXAMINAR",
                      font=("Courier New", 8, "bold"),
                      bg=ACC2, fg=ACC, activebackground=ACC,
                      activeforeground="#fff", relief="flat", bd=0,
                      padx=10, cursor="hand2",
                      command=lambda v=var, t=tipos: self._browse(v, t)
                      ).pack(side="right")

        self.var_feb = tk.StringVar()
        self.var_tiq = tk.StringVar()
        file_row(body, "1.  ARCHIVO FEBRERO 2026  (.xlsx)",
                 self.var_feb, [("Excel", "*.xlsx *.xls")])
        file_row(body, "2.  ARCHIVO TIQUETES  (.xls / .xlsx)",
                 self.var_tiq, [("Excel", "*.xls *.xlsx")])

        nota = tk.Frame(body, bg="#0d1b2a", padx=12, pady=8)
        nota.pack(fill="x", pady=(0, 5))
        tk.Label(nota,
                 text="📄  El archivo actualizado se generará automáticamente\n"
                      "    en la misma carpeta que el archivo FEBRERO 2026.",
                 font=("Courier New", 8), fg=MUTED, bg="#0d1b2a",
                 justify="left").pack(anchor="w")

        self.btn_run = tk.Button(body, text="▶  PROCESAR",
                                 font=("Courier New", 11, "bold"),
                                 bg=ACC, fg="#fff",
                                 activebackground="#c73652",
                                 activeforeground="#fff",
                                 relief="flat", bd=0, pady=10,
                                 cursor="hand2", command=self._run)
        self.btn_run.pack(fill="x", pady=(6, 5))

        log_frame = tk.Frame(body, bg=CARD, padx=8, pady=8)
        log_frame.pack(fill="both", expand=True, pady=(4, 0))
        tk.Label(log_frame, text="REGISTRO", font=("Courier New", 7, "bold"),
                 fg=MUTED, bg=CARD, anchor="w").pack(fill="x")
        self.log_text = tk.Text(log_frame, height=13, width=62,
                                font=("Courier New", 9),
                                bg="#0d1b2a", fg="#7effa0",
                                insertbackground="#7effa0",
                                relief="flat", bd=0, state="disabled",
                                wrap="word")
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.progress = ttk.Progressbar(body, mode="indeterminate", length=400)
        self.progress.pack(fill="x", pady=(6, 0))

        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TScrollbar", background=ACC2, troughcolor=CARD,
                        arrowcolor=MUTED, bordercolor=CARD)
        style.configure("TProgressbar", troughcolor=CARD,
                        background=ACC, bordercolor=CARD)

    def _browse(self, var, tipos):
        path = filedialog.askopenfilename(filetypes=tipos)
        if path:
            var.set(path)

    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _run(self):
        feb = self.var_feb.get().strip()
        tiq = self.var_tiq.get().strip()
        if not feb or not tiq:
            messagebox.showwarning("Campos vacíos",
                                   "Por favor selecciona los dos archivos.")
            return
        self.btn_run.configure(state="disabled", text="Procesando...")
        self.progress.start(12)
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        def worker():
            try:
                ruta, ok, multi, no = procesar(feb, tiq, self._log)
                self.after(0, lambda: messagebox.showinfo(
                    "Completado",
                    f"¡Proceso terminado!\n\n"
                    f"✅  Filas actualizadas: {ok}\n"
                    f"⚠️   Múltiples tiquetes: {multi}\n"
                    f"❌  Sin coincidencia: {no}\n\n"
                    f"Archivo guardado en:\n{ruta}"
                ))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
                self._log(f"\n❌ ERROR: {e}")
            finally:
                self.after(0, self._done)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self):
        self.progress.stop()
        self.btn_run.configure(state="normal", text="▶  PROCESAR")


if __name__ == "__main__":
    app = App()
    app.mainloop()

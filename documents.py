"""
comparar_tiquetes_gui.py
UI profesional con tkinter — tema claro refinado, tipografía elegante,
animaciones CSS-equivalentes via after(), feedback visual completo.
"""

import subprocess, os, shutil
import pandas as pd
from copy import copy
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime


from ui import *


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
        log_fn("  Convirtiendo .xls con LibreOffice...")
        ruta = convertir_xls(ruta, lo)
        df = pd.read_excel(ruta, header=5)
        df.columns = df.columns.str.strip()
        return df
    log_fn("  LibreOffice no encontrado, usando xlrd...")
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


def normalizar(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0).astype("int64")


def extraer_fecha_hora(valor):
    if pd.isna(valor):
        return None, None
    dt = valor if isinstance(valor, datetime) else datetime.strptime(str(valor).strip(), "%Y-%m-%d %H:%M:%S")
    return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S")


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
    df_tiq = leer_tiquetes(ruta_tiquetes, lambda m: log_fn("INFO", m))

    log_fn("INFO", f"MES: {len(df_feb)} filas  |  TIQUETES: {len(df_tiq)} filas")

    progress_fn(30)
    log_fn("INFO", "Comparando registros...")

    df_feb["_orden"] = normalizar(df_feb["Nro orden de compra"])
    df_feb["_doc"]   = normalizar(df_feb["Número de Documento"])
    df_tiq["_orden"] = normalizar(df_tiq["NRO ORDEN CREDITO"])
    df_tiq["_doc"]   = normalizar(df_tiq["CEDULA PASAJERO"])

    conteo           = df_tiq.groupby(["_orden", "_doc"])["TIQUETE"].count()
    claves_unicas    = set(conteo[conteo == 1].index)
    claves_multiples = set(conteo[conteo > 1].index)

    log_fn("OK",   f"Combinaciones únicas encontradas: {len(claves_unicas)}")
    log_fn("WARN", f"Combinaciones con múltiples tiquetes: {len(claves_multiples)}")

    mask      = df_tiq.apply(lambda r: (r["_orden"], r["_doc"]) in claves_unicas, axis=1)
    df_unicos = df_tiq[mask][["_orden", "_doc", "TIQUETE", "FEC.SALIDA"]].copy()
    df_unicos.set_index(["_orden", "_doc"], inplace=True)

    progress_fn(45)
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

    progress_fn(60)
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

    plan = []
    for df_idx in df_feb.index:
        excel_row_orig = df_idx + 2
        cantidad = wb_orig_ws.cell(excel_row_orig, col_cant).value
        try:
            cantidad = int(cantidad) if cantidad and cantidad > 0 else 1
        except (TypeError, ValueError):
            cantidad = 1
        tiquete, fecha, hora = resultados[df_idx]
        plan.append((excel_row_orig, cantidad, tiquete, fecha, hora))

    progress_fn(70)
    log_fn("INFO", "Expandiendo filas y escribiendo datos...")

    ws.delete_rows(2, ws.max_row - 1)

    filas_expandidas = 0
    write_row = 2

    for (orig_row, cantidad, tiquete, fecha, hora) in plan:
        for rep in range(cantidad):
            ws.append([None] * n_cols)
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
            if tiquete is not None:
                ws.cell(write_row, col_tiq).value = tiquete
                ws.cell(write_row, col_fec).value = fecha
                ws.cell(write_row, col_hora).value = hora
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
    log_fn("OK",   f"Filas actualizadas con tiquete: {filas_ok}")
    log_fn("WARN", f"Sin llenar (múltiples tiquetes): {filas_multi}")
    log_fn("ERR",  f"Sin coincidencia: {filas_no}")
    log_fn("OK",   f"Filas expandidas (pasajes > 1): {filas_expandidas}")
    log_fn("OK",   f"Total filas en archivo final: {total_filas}")

    return ruta_salida, filas_ok, filas_multi, filas_no, filas_expandidas, total_filas


if __name__ == "__main__":
    app = App()
    app.mainloop()

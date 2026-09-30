#!/usr/bin/env python3
"""Limpiador de Excel/CSV.

Recibe un archivo tabular desordenado y devuelve un .xlsx limpio con:
  - hoja "Limpio": datos normalizados (+ columna "avisos" con lo que hay que revisar)
  - hoja "Resumen": qué se ha cambiado, cuántas filas se han descartado, etc.

Uso:
    python limpiar_excel.py entrada.xlsx            -> entrada_limpio.xlsx
    python limpiar_excel.py entrada.csv -o salida.xlsx
    python limpiar_excel.py entrada.xlsx --hoja "Clientes"

El script NO modifica el archivo original.
"""
import argparse
import re
import sys
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FORMATOS_FECHA = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
    "%d/%m/%y", "%d-%m-%y",
]
PARTICULAS = {"de", "del", "la", "las", "los", "y", "e", "da", "di", "van", "von"}
EMAIL_RE = re.compile(r"^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)+$")


# ---------------------------------------------------------------- utilidades
def vacio(valor):
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    return isinstance(valor, str) and valor.strip() == ""


def sin_acentos(texto):
    return "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )


def normalizar_cabecera(texto):
    """' Fecha alta ' -> 'fecha_alta'; 'Teléfono' -> 'telefono'."""
    t = sin_acentos(str(texto)).strip().lower()
    t = re.sub(r"[^a-z0-9]+", "_", t)
    return t.strip("_")


def tipo_de_columna(cabecera):
    """Decide cómo limpiar una columna según su nombre."""
    c = cabecera
    if "mail" in c or "correo" in c:
        return "email"
    if any(k in c for k in ("tel", "movil", "celular", "phone")):
        return "telefono"
    if "fecha" in c or c in ("date", "alta", "nacimiento"):
        return "fecha"
    if any(k in c for k in ("importe", "precio", "total", "monto", "coste", "cantidad_eur")):
        return "importe"
    if any(k in c for k in ("nombre", "apellido", "ciudad", "localidad", "provincia", "cliente")):
        return "titulo"
    return "texto"


# ------------------------------------------------------------ limpiadores
def limpiar_texto(valor):
    if vacio(valor):
        return None
    return re.sub(r"\s+", " ", str(valor)).strip()


def limpiar_titulo(valor):
    t = limpiar_texto(valor)
    if t is None:
        return None
    palabras = t.lower().split(" ")
    salida = []
    for i, p in enumerate(palabras):
        salida.append(p if (i > 0 and p in PARTICULAS) else p.capitalize())
    return " ".join(salida)


def limpiar_email(valor):
    t = limpiar_texto(valor)
    if t is None:
        return None, None
    t = t.replace(" ", "").lower()
    if EMAIL_RE.match(t):
        return t, None
    return t, "email no válido"


def limpiar_telefono(valor):
    if vacio(valor):
        return None, None
    crudo = str(valor).strip()
    if re.fullmatch(r"\d+\.0", crudo):  # Excel lo leyó como número
        crudo = crudo[:-2]
    tiene_mas = crudo.startswith("+")
    digitos = re.sub(r"\D", "", crudo)
    if digitos.startswith("0034") and len(digitos) == 13:
        digitos = digitos[4:]
    elif tiene_mas and digitos.startswith("34") and len(digitos) == 11:
        digitos = digitos[2:]
    elif tiene_mas:
        return "+" + digitos, None  # extranjero: se deja tal cual
    if len(digitos) == 9:
        return f"{digitos[:3]} {digitos[3:6]} {digitos[6:]}", None
    return digitos, "teléfono con longitud no válida"


def limpiar_fecha(valor):
    if vacio(valor):
        return None, None
    if isinstance(valor, datetime):
        return valor.date(), None
    if isinstance(valor, date):
        return valor, None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        if 20000 < valor < 80000:  # número de serie de Excel
            return (datetime(1899, 12, 30) + timedelta(days=float(valor))).date(), None
        return None, "fecha no reconocida"
    t = str(valor).strip()
    if re.fullmatch(r"\d{5}(\.0+)?", t):  # número de serie de Excel guardado como texto
        return limpiar_fecha(float(t))
    t = re.sub(r"[ T]\d{1,2}:\d{2}(:\d{2})?$", "", t)  # quita una hora final (00:00:00)
    for fmt in FORMATOS_FECHA:
        try:
            return datetime.strptime(t, fmt).date(), None
        except ValueError:
            continue
    return None, f"fecha no válida ({t})"


def limpiar_importe(valor):
    """Acepta '1.234,50 €', '1234.5', '89 EUR', '1 200,00', etc."""
    if vacio(valor):
        return None, None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return round(float(valor), 2), None
    t = str(valor).strip()
    t = re.sub(r"[^\d,.\-]", "", t)  # quita €, EUR, espacios, etc.
    if not re.search(r"\d", t):
        return None, f"importe no válido ({valor})"
    if "," in t and "." in t:
        # el último separador es el decimal
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):
        # '2.500' se interpreta como miles (hipótesis habitual en España)
        t = t.replace(".", "")
    try:
        return round(float(t), 2), None
    except ValueError:
        return None, f"importe no válido ({valor})"


# ------------------------------------------------------------ lectura
def leer(ruta, hoja=None):
    ruta = Path(ruta)
    if ruta.suffix.lower() in (".csv", ".tsv"):
        sep = "\t" if ruta.suffix.lower() == ".tsv" else None
        try:
            return pd.read_csv(ruta, sep=sep, engine="python", dtype=object, encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(ruta, sep=sep, engine="python", dtype=object, encoding="latin-1")
    return pd.read_excel(ruta, sheet_name=hoja or 0, dtype=object)


# ------------------------------------------------------------ limpieza
def limpiar(df):
    resumen = {}
    filas_inicio = len(df)

    # 1) cabeceras y filas/columnas vacías
    df = df.copy()
    df.columns = [normalizar_cabecera(c) or f"columna_{i+1}" for i, c in enumerate(df.columns)]
    df = df.dropna(how="all").dropna(axis=1, how="all")
    resumen["Filas vacías eliminadas"] = filas_inicio - len(df)

    # 2) limpieza por tipo de columna
    avisos = [[] for _ in range(len(df))]
    cambios = 0
    for col in df.columns:
        tipo = tipo_de_columna(col)
        nuevos = []
        for i, original in enumerate(df[col].tolist()):
            aviso = None
            if tipo == "email":
                nuevo, aviso = limpiar_email(original)
            elif tipo == "telefono":
                nuevo, aviso = limpiar_telefono(original)
            elif tipo == "fecha":
                nuevo, aviso = limpiar_fecha(original)
            elif tipo == "importe":
                nuevo, aviso = limpiar_importe(original)
            elif tipo == "titulo":
                nuevo = limpiar_titulo(original)
            else:
                nuevo = limpiar_texto(original)
            if aviso:
                avisos[i].append(f"{col}: {aviso}")
            if str(nuevo) != str(None if vacio(original) else original):
                cambios += 1
            nuevos.append(nuevo)
        df[col] = pd.Series(nuevos, index=df.index, dtype=object)
    resumen["Celdas modificadas"] = cambios

    df["avisos"] = ["; ".join(a) for a in avisos]

    # 3) duplicados: por email si existe; si no, por toda la fila
    antes = len(df)
    col_email = next((c for c in df.columns if tipo_de_columna(c) == "email"), None)
    if col_email:
        con_email = df[col_email].notna()
        dup = df[con_email].duplicated(subset=[col_email], keep="first")
        indices_dup = dup[dup].index
        df = df.drop(index=indices_dup)
        sin_email = df[col_email].isna()
        dup2 = df[sin_email].drop(columns=["avisos"]).duplicated(keep="first")
        df = df.drop(index=dup2[dup2].index)
    else:
        dup = df.drop(columns=["avisos"]).duplicated(keep="first")
        df = df.drop(index=dup[dup].index)
    resumen["Duplicados eliminados"] = antes - len(df)

    df = df.reset_index(drop=True)
    resumen["Filas finales"] = len(df)
    resumen["Filas con avisos (revisar a mano)"] = int((df["avisos"] != "").sum())
    return df, resumen


# ------------------------------------------------------------ escritura
def guardar(df, resumen, ruta_salida):
    wb = Workbook()
    ws = wb.active
    ws.title = "Limpio"

    fuente = Font(name="Arial", size=10)
    cabecera = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    relleno_cab = PatternFill("solid", start_color="305496")
    relleno_aviso = PatternFill("solid", start_color="FFF2CC")

    ws.append(list(df.columns))
    for celda in ws[1]:
        celda.font = cabecera
        celda.fill = relleno_cab
        celda.alignment = Alignment(horizontal="center", vertical="center")

    for _, fila in df.iterrows():
        ws.append([None if vacio(v) else v for v in fila.tolist()])

    for fila in ws.iter_rows(min_row=2):
        con_aviso = bool(fila[-1].value)
        for celda in fila:
            celda.font = fuente
            if con_aviso:
                celda.fill = relleno_aviso
            if isinstance(celda.value, date):
                celda.number_format = "DD/MM/YYYY"
            elif isinstance(celda.value, float):
                celda.number_format = "#,##0.00"

    for i, col in enumerate(df.columns, start=1):
        valores = [len(str(col))] + [len(str(v)) for v in df[col] if not vacio(v)]
        ws.column_dimensions[get_column_letter(i)].width = min(max(valores) + 3, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    rs = wb.create_sheet("Resumen")
    rs.append(["Concepto", "Valor"])
    for k, v in resumen.items():
        rs.append([k, v])
    rs.append([])
    rs.append(["Notas"])
    notas = [
        "Las filas en amarillo tienen avisos: revísalas a mano.",
        "Teléfonos: se asume formato español de 9 dígitos (se quita el prefijo +34).",
        "Importes: '2.500' se interpreta como 2500 (miles), no como 2,5.",
        "Duplicados: se conserva la primera aparición (por email, o fila completa si no hay email).",
        "El archivo original no se modifica.",
    ]
    for n in notas:
        rs.append([n])
    for celda in rs[1]:
        celda.font = cabecera
        celda.fill = relleno_cab
    for fila in rs.iter_rows(min_row=2):
        for celda in fila:
            celda.font = fuente
    rs["A" + str(len(resumen) + 3)].font = Font(name="Arial", size=10, bold=True)
    rs.column_dimensions["A"].width = 70
    rs.column_dimensions["B"].width = 12

    wb.save(ruta_salida)


# ------------------------------------------------------------ principal
def main():
    p = argparse.ArgumentParser(description="Limpia un Excel/CSV desordenado.")
    p.add_argument("entrada", help="archivo .xlsx, .csv o .tsv")
    p.add_argument("-o", "--salida", help="archivo de salida (.xlsx)")
    p.add_argument("--hoja", help="nombre de la hoja (por defecto, la primera)")
    args = p.parse_args()

    entrada = Path(args.entrada)
    if not entrada.exists():
        sys.exit(f"No existe el archivo: {entrada}")
    salida = Path(args.salida) if args.salida else entrada.with_name(entrada.stem + "_limpio.xlsx")
    if salida.resolve() == entrada.resolve():
        sys.exit("La salida no puede ser el mismo archivo que la entrada.")

    df = leer(entrada, args.hoja)
    limpio, resumen = limpiar(df)
    guardar(limpio, resumen, salida)

    print(f"Listo: {salida}")
    for k, v in resumen.items():
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()

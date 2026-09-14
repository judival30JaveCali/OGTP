"""Genera el Reporte de Escalafón Profesoral a partir de los excels fuente.

Reemplaza, de forma simplificada, las fórmulas de la hoja Profesores_General
del "REPORTE MUESTRA ESCALAFON PROFESORAL.xlsx":

- Edad, Tiempo Categoría/Nivel        -> resta de fechas vectorizada con pandas.
- Productos Investigación/Docencia    -> equivalente a los SUMIFS del Excel,
  hecho con groupby + pivot_table.
- VALIDADOR (dentro del archivo de productos) -> comparación vectorizada de
  años en vez de un VLOOKUP + IF por cada fila.
- Promocion, Escala Intermedia        -> tablas de reglas en reglas.py, en
  vez de los IF anidados de 8-9 niveles del Excel.

Columnas que en el Excel de muestra son texto pegado a mano (Cargo, Género,
Actividad Preponderante, Último nivel de formación, Nivel de idioma, Fecha
última escala/nivel, Portafolio Creado, Tiempo Directivo) NO tienen fuente en
ninguno de los 5 archivos de esta carpeta, así que quedan vacías: se llenan
manualmente después de generar el reporte.
"""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

import reglas

COLUMNAS_FECHA = {"Fecha nacimiento", "Fecha ingreso Categoría/Nivel", "Fecha ultima escala / nivel"}

# El script vive fuera de la carpeta de datos (que no se versiona: tiene
# información personal/salarial). Por defecto se asume una carpeta hermana
# llamada "Proyecto_1_RegistroMensual", pero --data-dir permite apuntar a
# cualquier otra ruta si esa carpeta se mueve o se renombra.
CARPETA_DATOS_POR_DEFECTO = Path(__file__).resolve().parent / "Proyecto_1_RegistroMensual"

# Nombres (no rutas) de los 5 excels fuente dentro de la carpeta de datos.
NOMBRE_NOMINA = "Nomina Corte Julio.xlsx"
NOMBRE_PRODUCTOS = "PRODUCTOS TRAYECTORIA ACAD - 0309206.xlsx"
NOMBRE_VALORACION_2024 = "VALORACION GENERAL 2024.xlsx"
NOMBRE_VALORACION_2025 = "VALORACION GENERAL 2025.xlsx"
NOMBRE_DOCENCIA = "VALORACION DOCENCIA 2024-2025.xlsx"

# Nombres de columna reales de "Nomina Corte Julio.xlsx", en orden posicional.
# El archivo repite el encabezado "Descripción" varias veces (una por cada
# columna de código que describe), así que se renombra por posición en vez
# de por nombre para no depender del sufijo que pandas le agregue.
COLUMNAS_NOMINA = [
    "Tp Doc ID", "Doc ID", "ID Empleado", "Nº Reg Empl", "Nombre", "Correo-E",
    "Fecha Inicio Contrato", "Desc Estado", "Gp Pago", "Desc Gp Pago",
    "Puesto", "Desc Puesto", "Posición", "Desc Posición", "Hrs Estnd/Semn",
    "Tm Cmpl/Parc", "Unidad de Negocio", "Desc Unidad de Negocio",
    "Departamento", "Desc Departamento", "Proyecto", "Familia de Puesto",
    "Desc Familia de Puesto", "Desc Familia de Puesto 2", "Nombramiento",
    "Categoria", "ID Dpto", "Desc ID Dpto", "Nivel Salarial", "F Nacimiento",
    "Posc Jefe directo", "Cargo jefe directo", "CC Jefe directo",
    "Nombre Jefe Directo", "Directivos y coordinadores con personal a cargo",
]

COLUMNAS_SALIDA = [
    "Doc ID", "ID Empleado", "Nombre", "Correo-E", "Unidad",
    "Departamento adscrito", "Cargo", "Tiempo Directivo", "Fecha nacimiento",
    "Edad", "Genero", "Categoría", "Categ/Escala Intermedia/Nivel",
    "Fecha ingreso Categoría/Nivel", "Tiempo Categoría / Nivel",
    "Fecha ultima escala / nivel", "Actividad Preponderante",
    "Productos Investigación", "Productos Docencia", "Total",
    "Último nivel de formación", "Nivel de idioma",
    "Valoración Desempeño 2024", "Valoración Desempeño 2025",
    "Valoración Docencia 2024", "Valoración Docencia 2025", "Promocion",
    "Escala Intermedia", "Tiempo Directivo_Num", "Portafolio Creado",
]

# Columnas que ningún archivo de la carpeta puede llenar: quedan vacías para
# diligenciarlas a mano.
COLUMNAS_SIN_FUENTE = [
    "Cargo", "Tiempo Directivo", "Genero", "Fecha ultima escala / nivel",
    "Actividad Preponderante", "Último nivel de formación", "Nivel de idioma",
    "Portafolio Creado",
]


def normalizar_cedula(valor) -> str | None:
    """Deja solo dígitos para poder cruzar cédulas sin importar el formato
    (texto/número, espacios, prefijos) con el que vino cada archivo."""
    if pd.isna(valor):
        return None
    digitos = re.sub(r"\D", "", str(valor))
    return digitos or None


def normalizar_categ_escala(valor) -> str | None:
    """Corrige inconsistencias de texto vistas en la Nomina:
    'Titular- Escala 2' (falta espacio) y 'Instructor - Escala' (falta el
    número; en la práctica todos los instructores son Escala 2)."""
    if pd.isna(valor) or str(valor).strip() in ("", "N/A"):
        return None
    texto = re.sub(r"\s+", " ", str(valor).strip())
    texto = texto.replace("- Escala", " - Escala").replace("  ", " ")
    if texto == "Instructor - Escala":
        texto = "Instructor - Escala 2"
    return texto


def categoria_amplia(categ_escala: str | None) -> str:
    """'Asistente - Escala 3' -> 'Asistente'; cualquier cosa que no empiece
    con Instructor/Asistente/Asociado/Titular (p.ej. 'Nivel I Planta
    Catedráticos') cae en 'No categorizado', igual que en la muestra."""
    if categ_escala:
        for prefijo in ("Instructor", "Asistente", "Asociado", "Titular"):
            if categ_escala.startswith(prefijo):
                return prefijo
    return "No categorizado"


def parsear_tiempo_directivo(texto) -> float | None:
    """'15 año(s) y 0 mes(es)' -> 15.0. Se deja disponible por si en el
    futuro se agrega la columna 'Tiempo Directivo' con este formato."""
    if not isinstance(texto, str):
        return None
    m = re.search(r"(\d+)\s*año", texto)
    n = re.search(r"(\d+)\s*mes", texto)
    if not m:
        return None
    anios = int(m.group(1))
    meses = int(n.group(1)) if n else 0
    return round(anios + meses / 12, 1)


def cargar_roster(path: Path) -> pd.DataFrame:
    """Filtra la Nomina completa a solo profesores (Categoria != N/A) y arma
    las columnas base del roster: identificación, unidad/departamento y
    categoría/escala."""
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df.columns = COLUMNAS_NOMINA

    df["Categ/Escala Intermedia/Nivel"] = df["Categoria"].map(normalizar_categ_escala)
    df = df[df["Categ/Escala Intermedia/Nivel"].notna()].copy()

    roster = pd.DataFrame({
        "Doc ID": df["Doc ID"].map(normalizar_cedula),
        "ID Empleado": df["ID Empleado"].str.strip(),
        "Nombre": df["Nombre"].str.strip().str.title(),
        "Correo-E": df["Correo-E"].str.strip(),
        "Unidad": df["Unidad de Negocio"].str.strip(),
        # "Desc Departamento" (junto a la columna "Departamento") resultó ser
        # el nombre de la Carrera/Programa, no el departamento académico.
        # El departamento real está en "ID Dpto" / "Desc ID Dpto".
        "Departamento adscrito": df["Desc ID Dpto"].str.strip(),
        "Fecha nacimiento": pd.to_datetime(df["F Nacimiento"], errors="coerce"),
        "Categ/Escala Intermedia/Nivel": df["Categ/Escala Intermedia/Nivel"],
    })
    roster["Categoría"] = roster["Categ/Escala Intermedia/Nivel"].map(categoria_amplia)
    roster = roster.dropna(subset=["Doc ID"]).drop_duplicates(subset=["Doc ID"])
    return roster


def cargar_productos(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=0)
    df["Doc ID"] = df["Cedula"].map(normalizar_cedula)
    df["Fecha ingreso a la categoria"] = pd.to_datetime(
        df["Fecha ingreso a la categoria"], errors="coerce"
    )
    df["Año de publicación"] = pd.to_numeric(df["Año de publicación"], errors="coerce")
    df["Equivalencia"] = pd.to_numeric(df["Equivalencia"], errors="coerce").fillna(0)
    return df


def derivar_fecha_ingreso_categoria(productos: pd.DataFrame) -> pd.Series:
    """Un profesor sin roster propio de fechas de categoría: se toma la
    primera fecha no nula registrada en sus productos (debería ser la misma
    en todos sus productos). Si nunca tuvo productos, queda sin fecha."""
    return (
        productos.dropna(subset=["Fecha ingreso a la categoria"])
        .groupby("Doc ID")["Fecha ingreso a la categoria"]
        .first()
    )


def calcular_productos_validos(productos: pd.DataFrame, fecha_ingreso: pd.Series) -> pd.DataFrame:
    """Reemplaza VALIDADOR (VLOOKUP+IF) y los SUMIFS de Productos
    Investigación/Docencia por operaciones vectorizadas de pandas."""
    anio_ingreso = productos["Doc ID"].map(fecha_ingreso).dt.year
    valido = productos["Año de publicación"] >= anio_ingreso

    resumen = (
        productos[valido]
        .pivot_table(
            index="Doc ID", columns="Actividad académica", values="Equivalencia",
            aggfunc="sum", fill_value=0,
        )
    )
    for col in ("Investigación", "Docencia"):
        if col not in resumen.columns:
            resumen[col] = 0.0
    resumen = resumen.rename(columns={
        "Investigación": "Productos Investigación", "Docencia": "Productos Docencia",
    })[["Productos Investigación", "Productos Docencia"]]
    resumen["Total"] = resumen["Productos Investigación"] + resumen["Productos Docencia"]
    return resumen


def cargar_valoracion_desempeno(path: Path, columna_cedula: str, anio: int) -> pd.Series:
    df = pd.read_excel(path, sheet_name=0)
    cedulas = df[columna_cedula].map(normalizar_cedula)
    serie = df["escaladeresultadofinal"].groupby(cedulas).first()
    serie.name = f"Valoración Desempeño {anio}"
    return serie


def cargar_valoracion_docencia(path: Path) -> pd.DataFrame:
    """Promedio simple de totalgeneral*100 por cédula y año (año =
    Periodo // 10, p.ej. 20241 y 20242 -> 2024)."""
    df = pd.read_excel(path, sheet_name="FE 2024 - 2025")
    df["Doc ID"] = df["cedula"].map(normalizar_cedula)
    df["anio"] = (pd.to_numeric(df["Periodo"], errors="coerce") // 10).astype("Int64")
    df["totalgeneral"] = pd.to_numeric(df["totalgeneral"], errors="coerce") * 100

    promedio = df.groupby(["Doc ID", "anio"])["totalgeneral"].mean().unstack("anio").round(2)
    promedio.columns = [f"Valoración Docencia {int(a)}" for a in promedio.columns]
    return promedio


def construir_reporte(data_dir: Path) -> pd.DataFrame:
    roster = cargar_roster(data_dir / NOMBRE_NOMINA)

    productos = cargar_productos(data_dir / NOMBRE_PRODUCTOS)
    fecha_ingreso_categoria = derivar_fecha_ingreso_categoria(productos)
    productos_resumen = calcular_productos_validos(productos, fecha_ingreso_categoria)

    desempeno_2024 = cargar_valoracion_desempeno(data_dir / NOMBRE_VALORACION_2024, "Cédula", 2024)
    desempeno_2025 = cargar_valoracion_desempeno(data_dir / NOMBRE_VALORACION_2025, "ccevaluado", 2025)
    docencia = cargar_valoracion_docencia(data_dir / NOMBRE_DOCENCIA)

    reporte = roster.set_index("Doc ID", drop=False)
    reporte["Fecha ingreso Categoría/Nivel"] = fecha_ingreso_categoria
    reporte = reporte.join(productos_resumen).join(desempeno_2024).join(desempeno_2025).join(docencia)

    for col in ("Productos Investigación", "Productos Docencia", "Total"):
        reporte[col] = reporte[col].fillna(0)

    hoy = pd.Timestamp.today().normalize()
    reporte["Edad"] = ((hoy - reporte["Fecha nacimiento"]).dt.days / 365).round(1)
    reporte["Tiempo Categoría / Nivel"] = (
        (hoy - reporte["Fecha ingreso Categoría/Nivel"]).dt.days / 365
    ).round(1)

    for col in COLUMNAS_SIN_FUENTE:
        reporte[col] = pd.NA
    reporte["Tiempo Directivo_Num"] = reporte["Tiempo Directivo"].map(parsear_tiempo_directivo)

    filas = reporte.to_dict("records")
    reporte["Promocion"] = [reglas.evaluar_promocion(f) for f in filas]
    reporte["Escala Intermedia"] = [reglas.evaluar_escala_intermedia(f) for f in filas]

    reporte = reporte.reset_index(drop=True)
    return reporte[COLUMNAS_SALIDA]


def construir_resumen(reporte: pd.DataFrame) -> pd.DataFrame:
    """Equivalente simplificado a la tabla dinámica 'TD Prof Planta' de la
    muestra: profesores, productos de investigación y de docencia por
    departamento y categoría."""
    resumen = (
        reporte.groupby(["Departamento adscrito", "Categoría"], dropna=False)
        .agg(
            **{
                "Num Profesores": ("Doc ID", "count"),
                "Productos Investigación": ("Productos Investigación", "sum"),
                "Productos Docencia": ("Productos Docencia", "sum"),
            }
        )
        .reset_index()
        .sort_values(["Departamento adscrito", "Categoría"])
    )
    return resumen


def autoajustar_columnas(worksheet) -> None:
    for i, columna in enumerate(worksheet.columns, start=1):
        ancho = max((len(str(c.value)) for c in columna if c.value is not None), default=10)
        worksheet.column_dimensions[get_column_letter(i)].width = min(ancho + 2, 45)


def dar_formato(worksheet, df: pd.DataFrame) -> None:
    """Encabezado legible, panel de títulos congelado y formato de fecha
    sin la hora (que por defecto añade openpyxl a las columnas datetime)."""
    for celda in worksheet[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="4472C4")
    worksheet.freeze_panes = "A2"

    for i, columna in enumerate(df.columns, start=1):
        if columna in COLUMNAS_FECHA:
            for fila in range(2, worksheet.max_row + 1):
                worksheet.cell(row=fila, column=i).number_format = "yyyy-mm-dd"

    autoajustar_columnas(worksheet)


def guardar_reporte(reporte: pd.DataFrame, resumen: pd.DataFrame, destino: Path) -> None:
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        reporte.to_excel(writer, sheet_name="Profesores_General", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        dar_formato(writer.sheets["Profesores_General"], reporte)
        dar_formato(writer.sheets["Resumen"], resumen)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--data-dir", type=Path, default=CARPETA_DATOS_POR_DEFECTO,
        help=(
            "Carpeta con los 5 excels fuente (Nomina, Productos, Valoraciones). "
            f"Por defecto: {CARPETA_DATOS_POR_DEFECTO}"
        ),
    )
    parser.add_argument(
        "--output", type=Path, default=None,
        help="Ruta del Excel de salida. Por defecto se guarda dentro de --data-dir.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    data_dir = args.data_dir
    if not data_dir.is_dir():
        raise SystemExit(f"No existe la carpeta de datos: {data_dir}")
    destino = args.output or data_dir / f"Reporte_Escalafon_{datetime.today():%Y%m%d}.xlsx"

    reporte = construir_reporte(data_dir)
    resumen = construir_resumen(reporte)
    guardar_reporte(reporte, resumen, destino)
    print(f"Reporte generado: {destino}")
    print(f"  Profesores: {len(reporte)}")
    print(f"  Con productos válidos > 0: {(reporte['Total'] > 0).sum()}")


if __name__ == "__main__":
    main()

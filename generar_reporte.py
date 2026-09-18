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

Cargo, Categoría/Escala, Departamento/Facultad, fechas de ingreso, Actividad
Preponderante, Último nivel de formación, Nivel de idioma y Tiempo Directivo
se resuelven con la Nómina y con "REPORTE PROFESORES TRAYECTORIA
ACAD.xlsx" (fuente autorizada para casi todos estos, confirmada por el
vicerrectorado - ver Proyecto_1_RegistroMensual/AUDITORIA_COMPARACION.md).
Género y Portafolio Creado siguen sin fuente y quedan vacíos: se llenan
manualmente después de generar el reporte. Por pedido del vicerrectorado,
"Nivel de idioma" vacío se etiqueta "Sin diagnóstico" y "Valoración
Desempeño/Docencia" vacías se etiquetan "No aplica" en vez de dejarse en
blanco.
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
# VALORACION GENERAL 2024/2025.xlsx y VALORACION DOCENCIA 2024-2025.xlsx
# quedaron obsoletos: el vicerrectorado pidió usar en su lugar las bases
# oficiales de Valora de abajo, que traen desempeño y docencia en un solo
# archivo por año (ver punto 7 de AUDITORIA_COMPARACION.md).
NOMBRE_NOMINA = "Nomina Corte Julio.xlsx"
NOMBRE_PRODUCTOS = "PRODUCTOS TRAYECTORIA ACAD - 0309206.xlsx"
NOMBRE_TRAYECTORIA = "REPORTE PROFESORES TRAYECTORIA ACAD.xlsx"
NOMBRE_VALORACION_2024 = "Reporte Valoración desempeño 2024 Evaluación Profesores 3.xlsx"
NOMBRE_VALORACION_2025 = "Reporte Valoración desempeño 2025 Evaluación Profesores 3.xlsx"

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
    "Departamento adscrito", "Facultad", "Cargo", "Tiempo Directivo",
    "Fecha nacimiento", "Edad", "Genero", "Categoría",
    "Categ/Escala Intermedia/Nivel", "Fecha ingreso Categoría/Nivel",
    "Tiempo Categoría / Nivel", "Fecha ultima escala / nivel",
    "Actividad Preponderante", "Productos Investigación",
    "Productos Docencia", "Total", "Último nivel de formación",
    "Nivel de idioma", "Valoración Desempeño 2024", "Valoración Desempeño 2025",
    "Valoración Docencia 2024", "Valoración Docencia 2025", "Promocion",
    "Escala Intermedia", "Portafolio Creado",
]

# Columnas que ningún archivo de la carpeta puede llenar: quedan vacías para
# diligenciarlas a mano. La creación del portafolio quedó pendiente a
# propósito (hay columnas para eso en la Trayectoria Académica, pero el
# vicerrectorado pidió dejarlo para más adelante).
COLUMNAS_SIN_FUENTE = ["Genero", "Portafolio Creado"]

# Para estas columnas, una celda vacía no es un hueco de datos sino una
# etiqueta con significado propio (pedido explícito del vicerrectorado):
# "Sin diagnóstico" = no hay un nivel de idioma reconocible (A1-C2) en la
# Trayectoria Académica; "No aplica" = el profesor no aparece en el archivo
# de Valora de ese año (ingresó después, licencia, etc.).
ETIQUETA_SIN_IDIOMA = "Sin diagnóstico"
ETIQUETA_SIN_VALORACION = "No aplica"
COLUMNAS_VALORACION = [
    "Valoración Desempeño 2024", "Valoración Desempeño 2025",
    "Valoración Docencia 2024", "Valoración Docencia 2025",
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
    'Titular- Escala 2' (falta espacio), 'Instructor - Escala' o
    'Instructor' solos (falta el número; en la práctica todos los
    instructores son Escala 2) y 'Asistente'/'Asociado'/'Titular' solos, sin
    escala (el vicerrectorado pidió asumir Escala 1 en ese caso).

    Solo se usa como respaldo para profesores que no aparecen en
    "REPORTE PROFESORES TRAYECTORIA ACAD.xlsx", que es la fuente preferida
    de Categoría/Escala (ver `cargar_trayectoria_academica`)."""
    if pd.isna(valor) or str(valor).strip() in ("", "N/A"):
        return None
    texto = re.sub(r"\s+", " ", str(valor).strip())
    texto = texto.replace("- Escala", " - Escala").replace("  ", " ")
    if texto in ("Instructor - Escala", "Instructor"):
        texto = "Instructor - Escala 2"
    elif texto in ("Asistente", "Asociado", "Titular"):
        texto = f"{texto} - Escala 1"
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
    """'15 año(s) y 0 mes(es)' -> 15.0. Convierte el texto de "Tiempo en
    cargos directivos" de la Trayectoria Académica al número decimal que
    lleva la columna "Tiempo Directivo" del reporte."""
    if not isinstance(texto, str):
        return None
    m = re.search(r"(\d+)\s*año", texto)
    n = re.search(r"(\d+)\s*mes", texto)
    if not m:
        return None
    anios = int(m.group(1))
    meses = int(n.group(1)) if n else 0
    return round(anios + meses / 12, 1)


NIVELES_IDIOMA_ORDEN = ["A1", "A2", "B1", "B2", "C1", "C2"]


def normalizar_nivel_idioma(texto) -> str | None:
    """La Trayectoria Académica a veces repite el nivel varias veces
    separado por comas -uno por cada registro de idioma del profesor,
    p.ej. 'B2,C1'-; se toma el más alto de los que sea un nivel CEFR válido
    (A1-C2). Texto libre sin ese formato ('Alto', 'Sin asignar') no cuenta
    porque el motor de reglas compara contra A1..C2 exactos."""
    if not isinstance(texto, str):
        return None
    tokens = [t.strip().upper() for t in texto.split(",")]
    validos = [t for t in tokens if t in NIVELES_IDIOMA_ORDEN]
    if not validos:
        return None
    return max(validos, key=NIVELES_IDIOMA_ORDEN.index)


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
        "Cargo": df["Desc Puesto"].str.strip(),
        # "Desc Departamento" (junto a la columna "Departamento") resultó ser
        # el nombre de la Carrera/Programa, no el departamento académico.
        # El departamento real está en "ID Dpto" / "Desc ID Dpto". Se usa
        # como respaldo: la fuente preferida es "Nombre Unidad" de la
        # Trayectoria Académica (ver `cargar_trayectoria_academica`), que
        # además ya trae la Facultad y no confunde oficina con facultad de
        # origen para el personal directivo (Rectoría, Vicerrectorías).
        "Departamento adscrito": df["Desc ID Dpto"].str.strip(),
        "Fecha nacimiento": pd.to_datetime(df["F Nacimiento"], errors="coerce"),
        "Categ/Escala Intermedia/Nivel": df["Categ/Escala Intermedia/Nivel"],
    })
    roster["Categoría"] = roster["Categ/Escala Intermedia/Nivel"].map(categoria_amplia)
    roster = roster.dropna(subset=["Doc ID"]).drop_duplicates(subset=["Doc ID"])
    return roster


def cargar_trayectoria_academica(path: Path) -> pd.DataFrame:
    """Lee "REPORTE PROFESORES TRAYECTORIA ACAD.xlsx": la fuente que el
    vicerrectorado confirmó como autorizada para Departamento/Facultad,
    Categoría/Escala, fechas de ingreso a la categoría y a la escala,
    Actividad Preponderante, Último nivel de formación, Nivel de idioma y
    Tiempo Directivo (columna "Tiempo en cargos directivos", AH) (ver
    preguntas 1, 2, 3 y 4 en Proyecto_1_RegistroMensual/
    AUDITORIA_COMPARACION.md). No todos los profesores de la Nómina están
    acá (~14 de 333); para esos, `construir_reporte` conserva lo que
    `cargar_roster` ya derivó de la Nómina como respaldo."""
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df["Doc ID"] = df["Número de documento"].map(normalizar_cedula)
    df = df.dropna(subset=["Doc ID"]).drop_duplicates(subset=["Doc ID"])

    # "Profesor Asistente" -> "Asistente" (mismo formato que `categoria_amplia`).
    categoria = df["Categoría"].str.strip().str.replace(r"^Profesor\s+", "", regex=True)
    escala = df["Escala"].str.strip()
    escala_numerica = escala.str.fullmatch(r"\d+").fillna(False)
    categ_escala = categoria.copy()
    categ_escala[escala_numerica] = categoria[escala_numerica] + " - Escala " + escala[escala_numerica]

    trayectoria = pd.DataFrame({
        "Doc ID": df["Doc ID"],
        "Categoría": categoria,
        "Categ/Escala Intermedia/Nivel": categ_escala,
        "Departamento adscrito": df["Nombre Unidad"].str.strip(),
        "Facultad": df["Facultad"].str.strip(),
        "Fecha ingreso Categoría/Nivel": pd.to_datetime(
            df["Fecha de ingreso a la categoría"], errors="coerce"
        ),
        "Fecha ultima escala / nivel": pd.to_datetime(
            df["Fecha de ingreso a la escala"], errors="coerce"
        ),
        "Actividad Preponderante": df["Actividad preponderante"].str.strip(),
        "Último nivel de formación": df["Máximo nivel académico"].str.strip(),
        "Nivel de idioma": df["Nivel"].map(normalizar_nivel_idioma),
        "Tiempo Directivo": df["Tiempo en cargos directivos"].str.strip().map(parsear_tiempo_directivo),
    })
    return trayectoria.set_index("Doc ID")


def cargar_productos(path: Path) -> pd.DataFrame:
    """Lee el excel de productos académicos y normaliza los tipos (cédula,
    fechas, año y equivalencia numérica) para poder cruzarlo y sumarlo más
    adelante."""
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


def cargar_valoracion(path: Path, columna_cedula: str, anio: int) -> pd.DataFrame:
    """Cruce directo por cédula contra la base oficial que exporta Valora
    (un profesor, un resultado por año, nada que promediar): trae en el
    mismo archivo tanto el desempeño general ("escaladeresultadofinal",
    columna AR en 2024 / AW en 2025) como la valoración de Docencia
    ("consolidadofuenteestudiantes", ya en escala 0-100: columna Y en 2024 /
    AB en 2025). Reemplaza las VALORACION GENERAL 2024/2025.xlsx y
    VALORACION DOCENCIA 2024-2025.xlsx anteriores (ver punto 7 de
    AUDITORIA_COMPARACION.md); `columna_cedula` varía porque cada archivo
    anual la nombra distinto."""
    df = pd.read_excel(path, sheet_name=0)
    cedulas = df[columna_cedula].map(normalizar_cedula)
    return pd.DataFrame({
        f"Valoración Desempeño {anio}": df["escaladeresultadofinal"].groupby(cedulas).first(),
        f"Valoración Docencia {anio}": df["consolidadofuenteestudiantes"].groupby(cedulas).first().round(2),
    })


def construir_reporte(data_dir: Path) -> pd.DataFrame:
    """Arma la hoja Profesores_General completa: carga los 5 excels, cruza
    todo por cédula, calcula las columnas derivadas (edad, tiempo en
    categoría, productos, valoraciones) y evalúa el motor de reglas de
    Promoción y Escala Intermedia. Devuelve un DataFrame con las 31 columnas
    en el mismo orden que el reporte de muestra (más "Facultad", que no
    estaba en la muestra pero pidió agregarse)."""
    roster = cargar_roster(data_dir / NOMBRE_NOMINA)
    trayectoria = cargar_trayectoria_academica(data_dir / NOMBRE_TRAYECTORIA)

    productos = cargar_productos(data_dir / NOMBRE_PRODUCTOS)
    fecha_ingreso_categoria_productos = derivar_fecha_ingreso_categoria(productos)
    productos_resumen = calcular_productos_validos(productos, fecha_ingreso_categoria_productos)

    desempeno_2024 = cargar_valoracion(data_dir / NOMBRE_VALORACION_2024, "Cédula", 2024)
    desempeno_2025 = cargar_valoracion(data_dir / NOMBRE_VALORACION_2025, "ccevaluado", 2025)

    reporte = roster.set_index("Doc ID", drop=False)
    reporte = reporte.join(productos_resumen).join(desempeno_2024).join(desempeno_2025)

    # La Trayectoria Académica es la fuente confirmada para estos campos; se
    # usa con prioridad sobre lo que ya traía el roster de la Nómina (para
    # los ~14 profesores que no están en ese archivo, se conserva ese
    # respaldo en vez de dejarlos vacíos). "Facultad" y "Tiempo Directivo"
    # son enteramente nuevas: no tienen respaldo en la Nómina.
    for col in ("Categoría", "Categ/Escala Intermedia/Nivel", "Departamento adscrito"):
        reporte[col] = trayectoria[col].combine_first(reporte[col])
    reporte["Facultad"] = trayectoria["Facultad"]
    reporte["Fecha ingreso Categoría/Nivel"] = trayectoria["Fecha ingreso Categoría/Nivel"].combine_first(
        fecha_ingreso_categoria_productos
    )
    for col in ("Fecha ultima escala / nivel", "Actividad Preponderante",
                "Último nivel de formación", "Nivel de idioma", "Tiempo Directivo"):
        reporte[col] = trayectoria[col]

    for col in ("Productos Investigación", "Productos Docencia", "Total"):
        reporte[col] = reporte[col].fillna(0)
    reporte["Nivel de idioma"] = reporte["Nivel de idioma"].fillna(ETIQUETA_SIN_IDIOMA)
    for col in COLUMNAS_VALORACION:
        reporte[col] = reporte[col].fillna(ETIQUETA_SIN_VALORACION)

    hoy = pd.Timestamp.today().normalize()
    reporte["Edad"] = ((hoy - reporte["Fecha nacimiento"]).dt.days / 365).round(1)
    reporte["Tiempo Categoría / Nivel"] = (
        (hoy - reporte["Fecha ingreso Categoría/Nivel"]).dt.days / 365
    ).round(1)

    for col in COLUMNAS_SIN_FUENTE:
        reporte[col] = pd.NA

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
    """Ajusta el ancho de cada columna al contenido más largo que tenga,
    con un tope de 45 para que una celda con texto muy largo no deje la
    hoja inmanejable."""
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
    """Escribe el Excel final con las hojas Profesores_General y Resumen,
    ya formateadas."""
    with pd.ExcelWriter(destino, engine="openpyxl") as writer:
        reporte.to_excel(writer, sheet_name="Profesores_General", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        dar_formato(writer.sheets["Profesores_General"], reporte)
        dar_formato(writer.sheets["Resumen"], resumen)


def parse_args() -> argparse.Namespace:
    """Define y lee los argumentos de línea de comandos (--data-dir,
    --output)."""
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
    """Punto de entrada: valida la carpeta de datos, construye el reporte
    y lo guarda, imprimiendo un resumen de cuántos profesores salieron."""
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

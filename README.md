# Generador de Reporte Escalafón Profesoral

Script que arma el reporte mensual de Escalafón Profesoral a partir de los
excels fuente de la Oficina de Gestión de Trayectoria Profesoral, calculando
en Python lo que antes eran fórmulas de Excel (SUMIFS, VLOOKUP, IFs
anidados).

## Carpeta de datos

Los excels con datos de profesores (nómina, productos académicos,
valoraciones) **no están en este repositorio**: contienen información
personal y salarial, y quedan ignorados en `.gitignore`.

El script espera una carpeta con estos 5 archivos:

- `Nomina Corte Julio.xlsx`
- `PRODUCTOS TRAYECTORIA ACAD - 0309206.xlsx`
- `VALORACION GENERAL 2024.xlsx`
- `VALORACION GENERAL 2025.xlsx`
- `VALORACION DOCENCIA 2024-2025.xlsx`

Por defecto busca una carpeta hermana llamada `Proyecto_1_RegistroMensual`
(al lado de este script). Si esa carpeta se mueve, se renombra, o se quiere
usar el set de datos de otro mes, se indica con `--data-dir`.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
# Usa la carpeta de datos por defecto (Proyecto_1_RegistroMensual al lado del script)
python3 generar_reporte.py

# Apuntando a otra carpeta y/o archivo de salida
python3 generar_reporte.py --data-dir "/ruta/a/los/excels" --output "/ruta/Reporte.xlsx"
```

El resultado es un Excel con dos hojas:

- **Profesores_General**: una fila por profesor, con las mismas 30 columnas
  del reporte de muestra.
- **Resumen**: profesores, productos de investigación y de docencia por
  departamento y categoría (equivalente simplificado a la tabla dinámica
  "TD Prof Planta" de la muestra).

## Qué calcula y qué no

`generar_reporte.py` deriva de los 5 excels fuente: identificación,
unidad/departamento, categoría/escala, productos de investigación y
docencia, valoración de desempeño y de docencia, edad, tiempo en categoría,
y los resultados de los motores de reglas de **Promoción** y **Escala
Intermedia** (definidos en `reglas.py`, que reemplazan los IF anidados de
8-9 niveles del Excel original).

Columnas sin ninguna fuente en estos 5 archivos (`Cargo`, `Género`,
`Actividad Preponderante`, `Último nivel de formación`, `Nivel de idioma`,
`Fecha última escala/nivel`, `Portafolio Creado`, `Tiempo Directivo`) quedan
vacías a propósito, para llenarse manualmente.

Ver `Proyecto_1_RegistroMensual/AUDITORIA_COMPARACION.md` (no versionado)
para el detalle de la comparación contra el reporte de muestra y las
preguntas abiertas pendientes de resolver con la oficina.

## Archivos

- `generar_reporte.py` — carga los excels, calcula el reporte y lo exporta.
- `reglas.py` — tablas de reglas de Promoción y Escala Intermedia.
- `requirements.txt` — dependencias (pandas, openpyxl).

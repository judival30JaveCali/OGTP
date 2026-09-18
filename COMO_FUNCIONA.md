# Cómo funciona `generar_reporte.py`

Resumen del funcionamiento interno del script, para quien necesite
entenderlo o darle mantenimiento sin tener que leer todo el código.

## Idea general

El script lee 5 excels, cruza todo por número de cédula, calcula un puñado
de columnas (antes eran fórmulas de Excel) y arma un reporte nuevo con dos
hojas: `Profesores_General` (una fila por profesor) y `Resumen` (conteos por
departamento y categoría).

No modifica ni depende del reporte de muestra: se construye desde cero cada
vez que se corre.

## Los 5 archivos de entrada

| Archivo | Qué aporta |
|---|---|
| `Nomina Corte Julio.xlsx` | El roster base: quién es profesor, nombre, correo, fecha de nacimiento, unidad, cargo, y (de respaldo) departamento y categoría/escala. |
| `PRODUCTOS TRAYECTORIA ACAD...xlsx` | Los productos académicos (artículos, libros, etc.) de cada profesor, con su tipo (Investigación/Docencia) y su valor en puntos ("Equivalencia"). |
| `REPORTE PROFESORES TRAYECTORIA ACAD.xlsx` | Fuente confirmada por el vicerrectorado para Departamento/Facultad, Categoría/Escala, fechas de ingreso a la categoría y a la escala, Actividad Preponderante, Último nivel de formación y Nivel de idioma. |
| `Reporte Valoración desempeño 2024/2025 Evaluación Profesores 3.xlsx` | Base oficial de Valora: desempeño (`escaladeresultadofinal`) y valoración de Docencia (`consolidadofuenteestudiantes`), un archivo por año. Reemplazan a `VALORACION GENERAL 2024/2025.xlsx` y `VALORACION DOCENCIA 2024-2025.xlsx`, que quedaron sin usar. |

## Paso a paso

**1. Roster (`cargar_roster`)**
Lee la Nómina completa (incluye a todo el personal, no solo profesores) y
se queda solo con las filas donde `Categoria` no es `N/A` — ese es el
criterio para decir "esta persona es profesor". De ahí saca identificación,
unidad, cargo (`Desc Puesto`), y limpia el texto de categoría/escala (corrige
cosas como `Titular- Escala 2` sin espacio, `Instructor` sin escala -> Escala
2, o `Asistente`/`Asociado`/`Titular` sin escala -> Escala 1). También deriva
`Categoría` (el grupo amplio: Instructor/Asistente/Asociado/Titular)
quitándole el "- Escala N" a la categoría completa. Estos valores de
departamento y categoría/escala son solo el **respaldo**: se usan nada más
para los profesores que no aparecen en la Trayectoria Académica (paso 3).

**2. Productos (`cargar_productos`, `derivar_fecha_ingreso_categoria`,
`calcular_productos_validos`)**
Lee el archivo de productos y, por cada cédula, toma la primera fecha de
"ingreso a la categoría" que encuentre (se asume igual en todos sus
productos) — esto también es solo respaldo de la fecha de la Trayectoria
Académica (paso 3). Con la fecha de ingreso vigente calcula si cada
producto es "válido": solo cuenta si se publicó en el año en que el
profesor ya estaba en su categoría actual o después. Luego suma los puntos
("Equivalencia") de los productos válidos, separados por tipo, para obtener
`Productos Investigación`, `Productos Docencia` y `Total`.

**3. Trayectoria Académica (`cargar_trayectoria_academica`)**
Lee `REPORTE PROFESORES TRAYECTORIA ACAD.xlsx` y, por cédula, saca
Departamento (`Nombre Unidad`) y Facultad, Categoría/Escala (concatenando
las columnas `Categoría` y `Escala`), fecha de ingreso a la categoría y a la
escala, Actividad Preponderante, Último nivel de formación (`Máximo nivel
académico`) y Nivel de idioma (`Nivel`; si trae varios valores separados por
coma, se toma el más alto que sea un nivel CEFR válido A1-C2). No todos los
333 profesores de la Nómina están en este archivo (~14 no aparecen); para
esos, `construir_reporte` conserva el respaldo del roster/productos en vez
de dejarlos vacíos — Actividad Preponderante, formación e idioma sí quedan
vacíos porque no tienen ningún otro archivo de dónde salir.

**4. Valoraciones (`cargar_valoracion`)**
Cruce directo por cédula: cada profesor tiene un solo resultado por año,
tanto de desempeño (`escaladeresultadofinal`: `DESTACADO`, `MUY BUENO`,
etc.) como de Docencia (`consolidadofuenteestudiantes`, ya en escala 0-100),
ambos del mismo archivo anual de Valora.

**5. Columnas calculadas (`construir_reporte`)**
- `Edad` y `Tiempo Categoría / Nivel`: días entre la fecha correspondiente y
  hoy, divididos entre 365.
- `Tiempo Directivo`: se calcula a partir de "Tiempo en cargos directivos"
  de la Trayectoria Académica (texto tipo `"15 año(s) y 0 mes(es)"`),
  convertido a número decimal con `parsear_tiempo_directivo`. Vacío para
  quien nunca tuvo un cargo directivo.
- `Nivel de idioma` y `Valoración Desempeño/Docencia` vacíos se etiquetan
  como `"Sin diagnóstico"` y `"No aplica"` respectivamente en vez de
  dejarse en blanco (pedido explícito del vicerrectorado).

**6. Motor de reglas (`reglas.py`)**
`Promocion` y `Escala Intermedia` ya no son los IF anidados de 8-9 niveles
del Excel original: son listas de reglas (`REGLAS_PROMOCION`,
`REGLAS_ESCALA_INTERMEDIA`), cada una con sus condiciones (categoría,
tiempo mínimo, productos mínimos, formación, idioma, desempeño). El
script evalúa cada profesor contra la lista en orden y se queda con la
primera regla que aplique; si ninguna aplica, cae en el resultado por
defecto (`SIN POSIBILIDAD ASCENSO` / `SIN POSIBILIDAD CAMBIO ESCALA`, o
`MÁXIMA CATEGORÍA` si ya es Titular). Con Actividad Preponderante/formación/
idioma ahora casi siempre presentes, el motor sí marca elegibles (antes casi
nunca podía).

**7. Columnas sin fuente**
`Genero` y `Portafolio Creado` quedan vacías siempre: ningún archivo de
entrada trae esos datos (la Trayectoria Académica sí tiene columnas de
portafolio, pero el vicerrectorado pidió dejar esa parte pendiente por
ahora).

**8. Resumen (`construir_resumen`)**
Agrupa el reporte final por departamento y categoría, contando profesores y
sumando productos — el equivalente simplificado a una tabla dinámica.

**9. Exportar (`guardar_reporte`, `dar_formato`)**
Escribe las dos hojas con `pandas` + `openpyxl`, y les da formato: encabezado
en negrita, fila de títulos congelada, y formato de fecha sin hora en las
columnas de fecha.

## Cómo se corre

```bash
python3 generar_reporte.py                          # usa la carpeta de datos por defecto
python3 generar_reporte.py --data-dir "/ruta" --output "/ruta/Reporte.xlsx"
```

`--data-dir` apunta a la carpeta con los 5 excels (por defecto, la carpeta
hermana `Proyecto_1_RegistroMensual`). `--output` es opcional; si no se da,
el archivo se guarda dentro de `--data-dir` con la fecha de hoy en el
nombre.

## Limitaciones conocidas

- ~14 de los 333 profesores no están en `REPORTE PROFESORES TRAYECTORIA
  ACAD.xlsx`; para ellos, Actividad Preponderante/formación/Tiempo Directivo
  quedan vacíos y Nivel de idioma sale `"Sin diagnóstico"` (no hay otro
  archivo del que sacarlos), y Departamento/Categoría-Escala/fecha de
  ingreso caen al respaldo derivado de Nómina/Productos.
- `Actividad Preponderante` solo trae `Investigación` o `Docencia` en los
  datos actuales — nunca aparece `Planta de clínicos (no aplica)`, así que
  las reglas de `reglas.py` que exigen esa actividad puntual (no combinada
  con Docencia) no tienen cómo aplicar todavía.
- `Genero` y `Portafolio Creado` siguen sin fuente.
- `Departamento adscrito`/`Facultad` ahora usan el nombre completo de la
  Trayectoria Académica (p.ej. "Departamento de Gestión de Organizaciones"),
  que no coincide textualmente con la abreviación que traía el reporte de
  muestra ("Dpto Gestion de Organizaciones") aunque sea el mismo
  departamento — es un cambio de fuente esperado, no un error.

Detalle completo de estos puntos, con conteos, en
`Proyecto_1_RegistroMensual/AUDITORIA_COMPARACION.md` (no versionado).

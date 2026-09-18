# Cómo evalúa `reglas.py`

Diagramas del orden exacto en que `evaluar_promocion()` y
`evaluar_escala_intermedia()` recorren sus listas de reglas: de arriba hacia
abajo, gana la **primera** que aplique. Si ninguna aplica, cae en el
resultado por defecto.

## Promoción

```mermaid
flowchart TD
    A(["Profesor"]) --> G{"¿Desempeño 2024 y 2025<br/>en MUY BUENO o DESTACADO?"}
    G -- No --> Z1["SIN POSIBILIDAD ASCENSO"]
    G -- Sí --> R1{"1· Instructor, general<br/>tiempo en la categoría≥3,<br/>Productos de Investigación≥1,<br/>Productos de Docencia≥1<br/>formación e idioma nivel A"}
    R1 -- Sí --> RES1["POSIBLE ASCENSO ASIS"]
    R1 -- No --> R2{"2· Instructor, Planta de clínicos<br/>tiempo en la categoría≥3,<br/>Total de productos≥2<br/>formación e idioma nivel A"}
    R2 -- Sí --> RES1
    R2 -- No --> R3{"3· Asistente, Investigación<br/>tiempo en la categoría≥6,<br/>Productos de Investigación≥5,<br/>Productos de Docencia≥1<br/>formación e idioma nivel B"}
    R3 -- Sí --> RES2["POSIBLE ASCENSO ASOC"]
    R3 -- No --> R4{"4· Asistente, Docencia<br/>tiempo en la categoría≥6,<br/>Productos de Investigación≥2,<br/>Productos de Docencia≥2<br/>formación e idioma nivel B"}
    R4 -- Sí --> RES2
    R4 -- No --> R5{"5· Asistente, Planta de clínicos<br/>tiempo en la categoría≥6,<br/>Total de productos≥4<br/>formación e idioma nivel B"}
    R5 -- Sí --> RES2
    R5 -- No --> R6{"6· Asociado, Investigación<br/>tiempo en la categoría≥8,<br/>Productos de Investigación≥7,<br/>Productos de Docencia≥1<br/>formación e idioma nivel B"}
    R6 -- Sí --> RES3["POSIBLE ASCENSO TITU"]
    R6 -- No --> R7{"7· Asociado, Docencia<br/>tiempo en la categoría≥8,<br/>Productos de Investigación≥3,<br/>Productos de Docencia≥3<br/>formación e idioma nivel B"}
    R7 -- Sí --> RES3
    R7 -- No --> R8{"8· Asociado, Planta de clínicos<br/>tiempo en la categoría≥8,<br/>Total de productos≥6<br/>formación e idioma nivel B"}
    R8 -- Sí --> RES3
    R8 -- No --> R9{"¿Categoría = Titular?"}
    R9 -- Sí --> RES4["MÁXIMA CATEGORÍA"]
    R9 -- No --> Z1
```

## Escala Intermedia

```mermaid
flowchart TD
    A(["Profesor"]) --> G{"¿Valoración de Docencia 2024 y 2025<br/>ambas mayores a 90?"}
    G -- No --> Z["SIN POSIBILIDAD CAMBIO ESCALA"]
    G -- Sí --> R1{"1· Asistente - Escala 1, Investigación<br/>Total de productos≥3<br/>formación nivel C"}
    R1 -- Sí --> RES["POSIBLE CAMBIO ESCALA"]
    R1 -- No --> R2{"2· Asistente - Escala 1, Docencia o Planta de clínicos<br/>Total de productos≥2<br/>formación nivel C"}
    R2 -- Sí --> RES
    R2 -- No --> R3{"3· Asociado - Escala 1, Investigación<br/>Total de productos≥4<br/>formación nivel D"}
    R3 -- Sí --> RES
    R3 -- No --> R4{"4· Asociado - Escala 1, Docencia o Planta de clínicos<br/>Total de productos≥3<br/>formación nivel D"}
    R4 -- Sí --> RES
    R4 -- No --> R5{"5· Asociado - Escala 2, Investigación<br/>Total de productos≥6<br/>formación nivel D"}
    R5 -- Sí --> RES
    R5 -- No --> R6{"6· Asociado - Escala 2, Docencia o Planta de clínicos<br/>Total de productos≥5<br/>formación nivel D"}
    R6 -- Sí --> RES
    R6 -- No --> R7{"7· Titular - Escala 1, Investigación<br/>Total de productos≥8<br/>formación nivel D"}
    R7 -- Sí --> RES
    R7 -- No --> R8{"8· Titular - Escala 1, Docencia o Planta de clínicos<br/>Total de productos≥6<br/>formación nivel D"}
    R8 -- Sí --> RES
    R8 -- No --> Z
```

## Qué es cada "nivel" de formación/idioma

Para no repetir las mismas cuatro listas en cada casilla del diagrama, se
agrupan así (nombres tal cual en `reglas.py`):

| Nivel | Formación válida | Idioma válido |
|---|---|---|
| A (reglas 1-2, Instructor) | Maestría, Especialización Clínica Médica, Doctorado | B1, B2, C1, C2 |
| B (reglas 3-8, Asistente/Asociado → siguiente categoría) | Subespecialización Clínica Médica, Doctorado | B2, C1, C2 |
| C (reglas 1-2 de Escala, Asistente-Escala 1) | Maestría, Especialización Clínica Médica | — (no aplica idioma) |
| D (reglas 3-8 de Escala, Asociado/Titular) | Subespecialización Clínica Médica, Doctorado | — (no aplica idioma) |

Nota: los cortes "Investigación / Docencia / Planta de clínicos" corresponden
al campo `Actividad Preponderante` del profesor — hoy esa columna no tiene
fuente en ningún archivo (ver `COMO_FUNCIONA.md`), así que en la práctica
casi ninguna regla logra aplicar todavía.

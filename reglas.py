"""Reglas de negocio del Escalafón Profesoral.

Estas tablas reemplazan los dos IF anidados gigantes ("Promocion" y "Escala
Intermedia") de la hoja Profesores_General del reporte de muestra. Cada regla
es una condición completa; se evalúan en orden y gana la primera que aplique.
"""

from __future__ import annotations

from dataclasses import dataclass, field


DESEMPENO_VALIDO = {"MUY BUENO", "DESTACADO"}


@dataclass
class ReglaPromocion:
    categoria: str
    resultado: str
    actividad: str | None = None          # None = no se exige actividad puntual
    tiempo_min: float = 0
    prod_investigacion_min: float | None = None
    prod_docencia_min: float | None = None
    total_min: float | None = None
    formacion_validas: set[str] = field(default_factory=set)
    idiomas_validos: set[str] = field(default_factory=set)

    def aplica(self, fila: dict) -> bool:
        if fila.get("Categoría") != self.categoria:
            return False
        if self.actividad is not None and fila.get("Actividad Preponderante") != self.actividad:
            return False
        tiempo = fila.get("Tiempo Categoría / Nivel")
        if tiempo is None or tiempo < self.tiempo_min:
            return False
        if self.prod_investigacion_min is not None:
            if (fila.get("Productos Investigación") or 0) < self.prod_investigacion_min:
                return False
        if self.prod_docencia_min is not None:
            if (fila.get("Productos Docencia") or 0) < self.prod_docencia_min:
                return False
        if self.total_min is not None:
            if (fila.get("Total") or 0) < self.total_min:
                return False
        if self.formacion_validas and fila.get("Último nivel de formación") not in self.formacion_validas:
            return False
        if self.idiomas_validos and fila.get("Nivel de idioma") not in self.idiomas_validos:
            return False
        if fila.get("Valoración Desempeño 2024") not in DESEMPENO_VALIDO:
            return False
        if fila.get("Valoración Desempeño 2025") not in DESEMPENO_VALIDO:
            return False
        return True


@dataclass
class ReglaEscalaIntermedia:
    categoria_escala: str
    resultado: str
    actividades: set[str]
    total_min: float
    formacion_validas: set[str]
    umbral_docencia: float = 90

    def aplica(self, fila: dict) -> bool:
        if fila.get("Categ/Escala Intermedia/Nivel") != self.categoria_escala:
            return False
        if fila.get("Actividad Preponderante") not in self.actividades:
            return False
        if fila.get("Último nivel de formación") not in self.formacion_validas:
            return False
        if (fila.get("Total") or 0) < self.total_min:
            return False
        doc2024 = fila.get("Valoración Docencia 2024")
        doc2025 = fila.get("Valoración Docencia 2025")
        if not isinstance(doc2024, (int, float)) or doc2024 <= self.umbral_docencia:
            return False
        if not isinstance(doc2025, (int, float)) or doc2025 <= self.umbral_docencia:
            return False
        return True


FORMACION_INSTRUCTOR = {"Maestría", "Especialización Clínica Médica", "Doctorado"}
IDIOMA_INSTRUCTOR = {"B1", "B2", "C1", "C2"}
FORMACION_SUPERIOR = {"Subespecialización Clínica Médica", "Doctorado"}
IDIOMA_SUPERIOR = {"B2", "C1", "C2"}
PLANTA_CLINICOS = "Planta de clínicos (no aplica)"

# Orden importa: la primera regla que aplique define el resultado.
REGLAS_PROMOCION: list[ReglaPromocion] = [
    # Instructor -> Asistente
    ReglaPromocion(
        categoria="Instructor", resultado="POSIBLE ASCENSO ASIS", tiempo_min=3,
        prod_investigacion_min=1, prod_docencia_min=1,
        formacion_validas=FORMACION_INSTRUCTOR, idiomas_validos=IDIOMA_INSTRUCTOR,
    ),
    ReglaPromocion(
        categoria="Instructor", resultado="POSIBLE ASCENSO ASIS", tiempo_min=3,
        actividad=PLANTA_CLINICOS, total_min=2,
        formacion_validas=FORMACION_INSTRUCTOR, idiomas_validos=IDIOMA_INSTRUCTOR,
    ),
    # Asistente -> Asociado
    ReglaPromocion(
        categoria="Asistente", resultado="POSIBLE ASCENSO ASOC", tiempo_min=6,
        actividad="Investigación", prod_investigacion_min=5, prod_docencia_min=1,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
    ReglaPromocion(
        categoria="Asistente", resultado="POSIBLE ASCENSO ASOC", tiempo_min=6,
        actividad="Docencia", prod_investigacion_min=2, prod_docencia_min=2,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
    ReglaPromocion(
        categoria="Asistente", resultado="POSIBLE ASCENSO ASOC", tiempo_min=6,
        actividad=PLANTA_CLINICOS, total_min=4,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
    # Asociado -> Titular
    ReglaPromocion(
        categoria="Asociado", resultado="POSIBLE ASCENSO TITU", tiempo_min=8,
        actividad="Investigación", prod_investigacion_min=7, prod_docencia_min=1,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
    ReglaPromocion(
        categoria="Asociado", resultado="POSIBLE ASCENSO TITU", tiempo_min=8,
        actividad="Docencia", prod_investigacion_min=3, prod_docencia_min=3,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
    ReglaPromocion(
        categoria="Asociado", resultado="POSIBLE ASCENSO TITU", tiempo_min=8,
        actividad=PLANTA_CLINICOS, total_min=6,
        formacion_validas=FORMACION_SUPERIOR, idiomas_validos=IDIOMA_SUPERIOR,
    ),
]

REGLAS_ESCALA_INTERMEDIA: list[ReglaEscalaIntermedia] = [
    ReglaEscalaIntermedia(
        categoria_escala="Asistente - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Investigación"}, total_min=3,
        formacion_validas={"Maestría", "Especialización Clínica Médica"},
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Asistente - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Docencia", PLANTA_CLINICOS}, total_min=2,
        formacion_validas={"Maestría", "Especialización Clínica Médica"},
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Asociado - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Investigación"}, total_min=4, formacion_validas=FORMACION_SUPERIOR,
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Asociado - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Docencia", PLANTA_CLINICOS}, total_min=3, formacion_validas=FORMACION_SUPERIOR,
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Asociado - Escala 2", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Investigación"}, total_min=6, formacion_validas=FORMACION_SUPERIOR,
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Asociado - Escala 2", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Docencia", PLANTA_CLINICOS}, total_min=5, formacion_validas=FORMACION_SUPERIOR,
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Titular - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Investigación"}, total_min=8, formacion_validas=FORMACION_SUPERIOR,
    ),
    ReglaEscalaIntermedia(
        categoria_escala="Titular - Escala 1", resultado="POSIBLE CAMBIO ESCALA",
        actividades={"Docencia", PLANTA_CLINICOS}, total_min=6, formacion_validas=FORMACION_SUPERIOR,
    ),
]


def evaluar_promocion(fila: dict) -> str:
    for regla in REGLAS_PROMOCION:
        if regla.aplica(fila):
            return regla.resultado
    if fila.get("Categoría") == "Titular":
        return "MÁXIMA CATEGORÍA"
    return "SIN POSIBILIDAD ASCENSO"


def evaluar_escala_intermedia(fila: dict) -> str:
    for regla in REGLAS_ESCALA_INTERMEDIA:
        if regla.aplica(fila):
            return regla.resultado
    return "SIN POSIBILIDAD CAMBIO ESCALA"

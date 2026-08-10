# ==============================================================================
# PROMPTS ESPECIALIZADOS Y COMPLETOS DE EDUCACIÓN FÍSICA (CNEB MINEDU)
# ==============================================================================

def obtener_prompt_unidad(
    num_doc, ciclo_actual, grado_seccion, ie_nombre, docente, director,
    duracion_semanas, sesiones_por_semana, fechas_duracion, problema_contexto,
    producto_unidad, cneb_datos_text
):
    total_sesiones = duracion_semanas * sesiones_por_semana
    return f"""
Actúa como un especialista en currículo educativo peruano y docente experto en el área de Educación Física para Educación Básica Regular (CNEB). 

Tu tarea es elaborar una UNIDAD DE APRENDIZAJE completa, extensa, rigurosa y alineada al Currículo Nacional (CNEB), siguiendo estrictamente las 10 secciones obligatorias sin cortar el documento al final.

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y ESTRUCTURA (OBLIGATORIO LLEGAR HASTA LA SECCIÓN X):
1. DEBES FINALIZAR EL DOCUMENTO OBLIGATORIAMENTE HASTA LA SECCIÓN X (RECURSOS Y ESPACIO PARA FIRMAS DE DIRECTORA Y DOCENTE). QUEDA STRICTAMENTE PROHIBIDO DEJAR EL DOCUMENTO INCOMPLETO.
2. EN LA SECCIÓN VIII (MATRIZ DE PLANIFICACIÓN), DESARROLLA CADA UNA DE LAS {total_sesiones} SESIONES ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana). ESTÁ PROHIBIDO PONER PUNTOS SUSPENSIVOS (...) O OMITIR SESIONES.
3. EN LA MATRIZ DE PLANIFICACIÓN: TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN LA PARTE SUPERIOR DE CADA SESIÓN CON NEGRITA EN LA PARTE EVALUADA, Y EL DESEMPEÑO COMPLETO EN LA COLUMNA CORRESPONDIENTE CON NEGRITA EN LO UTILIZADO Y PRECISADO.
4. COMPLETA SIEMPRE LA SECCIÓN IX (SECUENCIA DE SESIONES CON SUS PROPÓSITOS Y REPRESENTACIONES GRÁFICAS) Y LA SECCIÓN X (RECURSOS Y ESPACIO PARA FIRMAS).

DATOS OFICIALES EXTRAÍDOS DE cneb_datos.py PARA ESTA UNIDAD ({grado_seccion} - {ciclo_actual}):
{cneb_datos_text}

DATOS PARA LA GENERACIÓN:
- N° de Unidad: Unidad N° {num_doc}
- Ciclo / Grados: {ciclo_actual} - {grado_seccion}
- Nombre de la IE: {ie_nombre}
- Nombre del Docente: {docente}
- Nombre del Director(a): {director}
- Duración / Fechas: {duracion_semanas} semanas ({total_sesiones} sesiones en total, {sesiones_por_semana} por semana) - ({fechas_duracion})
- Tema central / Problemática a abordar: {problema_contexto}
- Producto de la Unidad: {producto_unidad}

---

ESTRUCTURA OBLIGATORIA DE LA UNIDAD DE APRENDIZAJE DE EDUCACIÓN FÍSICA:

1. TÍTULO DE LA UNIDAD
- Debe ser motivador, entre comillas y redactado en función al desarrollo de competencias motrices, sociomotrices o de vida saludable.

2. II. DATOS INFORMATIVOS
- IE, Directora, Profesor de Ed. Física, Ciclo, Grado y Sección, Duración.

3. III. SITUACIÓN SIGNIFICATIVA
- Contextualizar la realidad motriz y de salud de los estudiantes relacionada con la problemática: {problema_contexto}.
- Incluir un dato cuantitativo/cualitativo del problema (ej. "solo el 35% logra orientarse adecuadamente...").
- Plantear 3 preguntas retadoras/desafiantes asociadas a la solución motriz.
- Proponer la estrategia pedagógica para resolver el reto (circuitos, festivales lúdico-motores, juegos tradicionales, etc.).

4. IV. PRODUCTO DE LA UNIDAD
- Describir de forma clara y extensa el desempeño práctico o producto tangible/demostrable: {producto_unidad}.

5. V. ENFOQUES TRANSVERSALES
- Seleccionar 2 enfoques transversales del CNEB.
- Especificar en tabla: Enfoque Transversal, Valor(es) y Acciones o Actitudes Observables adaptadas a Educación Física.

6. VI. COMPETENCIAS TRANSVERSALES
- Incluir en tabla "Gestiona su aprendizaje de manera autónoma" y "Se desenvuelve en entornos virtuales generados por las TIC" con sus respectivas Capacidades y Desempeños aplicados al área.

7. VII. ESTÁNDARES, COMPETENCIAS Y CAPACIDADES DEL ÁREA DE EDUCACIÓN FÍSICA
- Transcribir las 3 competencias oficiales del área con sus capacidades y estándares completos del ciclo correspondiente ({ciclo_actual}):
  * Competencia 1: Se desenvuelve de manera autónoma a través de su motricidad.
  * Competencia 2: Asume una vida saludable.
  * Competencia 3: Interactúa a través de sus habilidades sociomotrices.

8. VIII. MATRIZ DE PLANIFICACIÓN (Formato Tabla detallado por las {total_sesiones} sesiones)
Desarrolla {total_sesiones} bloques de tablas independientes (uno por cada sesión):
- En la parte superior de cada bloque de sesión, incluye la fila con el ESTÁNDAR COMPLETO del CNEB correspondiente a la competencia evaluada, redactado de manera íntegra (sin modificar ni alterar su texto original), RESALTANDO EN NEGRITA la parte específica que se trabaja/evalúa en esa actividad.
- Columnas de la Matriz por cada sesión:
  | Sesión N.° y Título de la sesión | Competencia / Capacidad | Desempeño | Criterios de Evaluación | Evidencia y Producto | Instrumento de Evaluación |
- REGLA DEL DESEMPEÑO: Redactado de manera COMPLETA tal cual aparece en el CNEB, RESALTANDO EN NEGRITA tanto la parte del desempeño utilizada como las palabras/términos agregados para su precisión y contextualización.
*NOTA: NO incluir la columna "Propósito" en la Matriz de Planificación.*

9. IX. SECUENCIA DE SESIONES (Formato Tabla)
Genera una tabla completa para las {total_sesiones} sesiones detallando:
| N° | Título de la actividad | Propósito de la actividad | Representación gráfica |
- El propósito debe ser explícito e incluir la secuencia metodológica (calentamiento/activación, desarrollo motriz/juego, hábitos de higiene personal y reflexión).
- La representación gráfica describe brevemente el esquema visual o distribución de materiales en el patio.

10. X. RECURSOS
- Recursos para el Docente (Normativa CNEB, RM N° 501-2025, materiales).
- Recursos para el Estudiante (Kit de aseo: jabón, toalla, polo de cambio, ropa deportiva, botellas de agua).
- Fecha y espacio para firmas de la Directora y Docente de Educación Física.
"""


def obtener_prompt_proyecto(
    num_doc, ciclo_actual, grado_seccion, dre_ugel, ie_nombre, docente, director,
    duracion_semanas, sesiones_por_semana, fechas_duracion, problema_contexto,
    producto_unidad, cneb_datos_text
):
    total_sesiones = duracion_semanas * sesiones_por_semana
    return f"""
Actúa como un Especialista Pedagógico experto en Educación Física del Ministerio de Educación de Perú (MINEDU). Tu tarea es diseñar un Proyecto de Aprendizaje completo, detallado y técnicamente profundo bajo el enfoque por competencias del Currículo Nacional de la Educación Básica (CNEB), manteniendo de manera estricta y detallada una estructura formal sin cortar el documento al final.

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y DESARROLLO TÉCNICO (OBLIGATORIO LLEGAR HASTA LA SECCIÓN IX):
1. DEBES DESARROLLAR EL PROYECTO COMPLETO LLEGANDO OBLIGATORIAMENTE HASTA LA SECCIÓN IX (RECURSOS, MATERIALES Y FIRMAS DE DIRECTORA Y DOCENTE). QUEDA STRICTAMENTE PROHIBIDO CORTAR EL DOCUMENTO.
2. EN LA SECCIÓN VII (CUADRO CRONOLÓGICO DE SESIONES), DESARROLLA LAS {total_sesiones} SESIONES ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana) UNA POR UNA. ESTÁ PROHIBIDO USAR PUNTOS SUSPENSIVOS (...) O OMITIR SESIONES.
3. EN LA SECCIÓN VI (MATRIZ DE PROPÓSITOS), TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN LA PARTE SUPERIOR DE CADA TABLA CON NEGRITA EN LA PARTE TRABAJADA, Y EL DESEMPEÑO COMPLETO EN LA COLUMNA CORRESPONDIENTE CON NEGRITA EN LO UTILIZADO Y PRECISADO.

DATOS OFICIALES EXTRAÍDOS DE cneb_datos.py PARA ESTE PROYECTO ({grado_seccion} - {ciclo_actual}):
{cneb_datos_text}

DATOS PARA LA GENERACIÓN:
- N° de Proyecto: Proyecto N° {num_doc}
- DRE / UGEL: {dre_ugel}
- Institución Educativa: {ie_nombre}
- Nivel: Educación Primaria
- Ciclo: {ciclo_actual}
- Grado y Sección: {grado_seccion}
- Área Curricular: Educación Física
- Duración y Frecuencia: {duracion_semanas} semanas, {sesiones_por_semana} sesiones por semana = {total_sesiones} sesiones en total ({fechas_duracion})
- Tema o Problemática Central: {problema_contexto}
- Producto Final: {producto_unidad}

---

ESTRUCTURA OBLIGATORIA DEL PROYECTO DE APRENDIZAJE DE EDUCACIÓN FÍSICA:

I. TÍTULO DEL PROYECTO
- Debe ser motivador, creativo, retador y entre comillas (Ejemplo: "¡CELEBRAMOS NUESTRA PERUANIDAD EN EL GRAN FESTIVAL LÚDICO-MOTOR!").

II. DATOS INFORMATIVOS
- DRE/UGEL, IE, Nivel, Ciclo, Grado y Sección, Área (Educación Física), Duración, N° de sesiones, Docente, Director(a).

III. SITUACIÓN SIGNIFICATIVA
Redacta una situación basada en un contexto real de la escuela en 4 bloques detallados:
- Contexto completo, Problema o necesidad ({problema_contexto}), Reto (2 a 3 preguntas desafiantes) y Propósito pedagógico amplio.

IV. CUADRO DE ENFOQUES TRANSVERSALES
Elabora una tabla con 1 o 2 enfoques transversales más pertinentes (Enfoque, Valores, Actitudes observables en Ed. Física).

V. CUADRO DE NEGOCIACIÓN Y PLANIFICACIÓN CON LOS ESTUDIANTES
Tabla de 4 columnas (¿Qué queremos hacer?, ¿Cómo lo haremos?, ¿Qué necesitamos?, ¿Cómo nos daremos cuenta de que lo logramos?) con respuestas ricas y participativas de una asamblea de Educación Física.

VI. CUADRO DE PROPÓSITOS DE APRENDIZAJE Y EVALUACIÓN MATRIZADA
Organiza una tabla por competencia (C1, C2, C3 de Educación Física):
- PARTE SUPERIOR DE CADA TABLA: Estándar COMPLETO del {ciclo_actual} del CNEB con **negrita** en la parte movilizada.
- Tabla con 7 columnas exactas:
  | Actividad General por Semana | Sesiones Vinculadas | Estándar Completo (encima) | Desempeño Completo (con **negrita**) | 3 Criterios de Evaluación | Evidencia de Aprendizaje | Instrumento |

VII. PLANIFICACIÓN CRONOLÓGICA DETALLADA DE LAS SESIONES
Desglosa secuencialmente las {total_sesiones} sesiones. Tabla obligatoria de 3 COLUMNAS:
| Denominación de la sesión | Propósito detallado de la sesión | Representación gráfica |
- Denominación: Número y título motivador entre comillas.
- Propósito detallado: Explicación pedagógica clara que incluya calentamiento, desarrollo motriz e higiene personal.
- Representación gráfica: Descripción breve de la imagen o esquema del patio.

VIII. PRODUCTOS DEL PROYECTO
- Producto Intangible / Práctico (Festival, Mini olimpiadas, Gincana).
- Producto Tangible ({producto_unidad}).

IX. RECURSOS Y MATERIALES
- Material deportivo, material reciclado, materiales de higiene.
- Espacios educativos y espacio para firmas de la Directora y Docente de Educación Física.
"""


def obtener_prompt_sesion(
    num_doc, ciclo_actual, grado_seccion, ie_nombre, docente, fecha_sugerida,
    duracion_sesion, tipo_motivacion, problema_contexto, comps_str, cap_str,
    est_str, crit_str, evid_str
):
    return f"""
Actúa como Docente Experto en Educación Física para Primaria bajo el enfoque oficial del CNEB del MINEDU Perú.
Elabora una SESIÓN DE CLASE PRÁCTICA DE EDUCACIÓN FÍSICA completa, extensa y bien explicada para {grado_seccion} ({ciclo_actual}).

DATOS INGRESADOS PARA LA SESIÓN:
- N.° de Sesión: {num_doc}
- Título de la actividad: "{problema_contexto}"
- IE: {ie_nombre} | Docente: {docente} | Fecha: {fecha_sugerida} | Duración: {duracion_sesion}
- Tipo de Motivación elegida: {tipo_motivacion}
- Competencia(s) solicitada(s): {comps_str}
- Capacidades solicitadas: {cap_str}
- Estándar solicitado: {est_str}
- Criterios solicitados: {crit_str}
- Evidencia solicitada: {evid_str}

---

REGLAS DE FORMATO Y ESTRUCTURA OBLIGATORIA DE LA SESIÓN (DESARROLLAR COMPLETA DE PRINCIPIO A FIN):

1. ENCABEZADO Y TÍTULO DE LA SESIÓN:
Muestra EXACTAMENTE la siguiente estructura en la parte superior:
# **SESIÓN DE APRENDIZAJE DE EDUCACIÓN FÍSICA N.º {num_doc}**
## **"{problema_contexto.upper()}"**
*(QUEDA STRICTAMENTE PROHIBIDO COLOCAR CUALQUIER OTRO DATO O FECHA DEBAJO DEL TÍTULO DE LA SESIÓN).*

2. TABLA I: DATOS INFORMATIVOS
| DATOS INFORMATIVOS | DETALLE |
| Institución Educativa | {ie_nombre} |
| Docente de Educación Física | {docente} |
| Grado y Sección | {grado_seccion} ({ciclo_actual}) |
| Fecha | {fecha_sugerida} |
| Duración | {duracion_sesion} |

3. TABLA II: PROPÓSITOS DE APRENDIZAJE Y EVIDENCIAS
> **ESTÁNDAR CNEB COMPLETO ({ciclo_actual}):** [Texto íntegro del estándar del ciclo con **negrita** en la parte aplicada]

| ÁREA | COMPETENCIA Y CAPACIDADES | DESEMPEÑO PRECISADO COMPLETO (con **negrita**) | CRITERIOS DE EVALUACIÓN | PROPÓSITO DE LA CLASE | EVIDENCIA | INSTRUMENTO |

4. TABLA III: ENFOQUE TRANSVERSAL (ÚNICO Y ESPECÍFICO)
| ENFOQUE TRANSVERSAL PRIORIZADO | VALOR(ES) | ACTITUDES OBSERVABLES |

5. TABLA IV: COMPETENCIAS TRANSVERSALES
| COMPETENCIA TRANSVERSAL | CAPACIDADES | DESEMPEÑOS PRECISADOS |

6. TABLA V: PREPARACIÓN DE LA CLASE
| ¿Qué necesitamos hacer antes de la sesión de Ed. Física? | ¿Qué recursos o materiales del patio se utilizarán? |

7. MOMENTOS DE LA CLASE DE EDUCACIÓN FÍSICA:

- **INICIO (Aprox. 20 min):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE. Debe considerar ESTRICTAMENTE el siguiente orden:
  1. **Motivación ({tipo_motivacion}):** [Desarrollar ampliamente la motivación según el tipo elegido: {tipo_motivacion}].
  2. **Saberes previos:** [Preguntas abiertas sobre el tema/movimientos].
  3. **Problematización / Conflicto cognitivo:** [Reto motriz o pregunta desafiante].
  4. **Propósito de la clase:** [Comunicar qué aprenderán hoy].
  5. **Criterios de evaluación:** [Explicar cómo serán evaluados].
  6. **Acuerdos de convivencia:** [2 a 3 normas de seguridad en el patio].

- **DESARROLLO (Aprox. 60 min):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE:
  1. **Activación Corporal (Calentamiento dinámico):** Movilidad articular, trote lúdico y estiramientos.
  2. **Secuencia de Actividades Motrices:** Progresión práctica bien explicada (3 actividades con hidratación).
  3. **ACTIVIDAD OBLIGATORIA DE ALTO NIVEL COGNITIVO (Analizar, Evaluar y Crear):** Reto motriz/estratégico donde los alumnos deban analizar una situación de juego, evaluar variantes tácticas y crear su propia regla o estrategia en equipo.

- **CIERRE (Aprox. 10 min) - DEBES REDACTAR OBLIGATORIAMENTE Y EN SU TOTALIDAD LOS SIGUIENTES 3 PUNTOS (PROHIBIDO OMITIR CUALQUIER PUNTO DEL CIERRE):**
  1. **Vuelta a la calma:** Ejercicios de respiración guiada y estiramientos suaves.
  2. **Metacognición motriz:** Redacta de 3 a 4 preguntas reflexivas explícitas (¿Qué aprendimos sobre nuestro cuerpo? ¿Cómo superamos retos?).
  3. **Rutina Obligatoria de Higiene Personal:** Describe en detalle la práctica autónoma de aseo personal, lavado de manos con jabón, secado con toalla y cambio de polo deportivo.

8. TABLA VI: LISTA DE COTEJO DE EDUCACIÓN FÍSICA (Genera una tabla limpia con los criterios de evaluación y 8 a 10 estudiantes ficticios representativos para optimizar espacio y garantizar la completitud del documento).
"""


def obtener_prompt_ficha(num_doc, problema_contexto, grado_seccion, ie_nombre, fecha_sugerida):
    return f"""
Actúa como Especialista en Educación Física Primaria CNEB.
Elabora una FICHA DE TRABAJO Y AUTOEVALUACIÓN DE EDUCACIÓN FÍSICA PARA EL ESTUDIANTE sobre {problema_contexto} para {grado_seccion}.

ESTRUCTURA REQUERIDA:
# **FICHA DE AUTOEVALUACIÓN Y SALUD EN EDUCACIÓN FÍSICA N.º {num_doc}**
## **{problema_contexto.upper()}**

- DATOS INFORMATIVOS (IE: {ie_nombre}, Estudiante: ___________________, Grado: {grado_seccion}, Fecha: {fecha_sugerida}).
- PROPÓSITO DEL DÍA (Explicado para niños).
- SECCIÓN 1: MIS REACCIONES CORPORALES (Dibujar o marcar ritmo cardiaco, sudoración y respiración tras el juego).
- SECCIÓN 2: MI COMPROMISO DE HIGIENE Y SALUD (Marcar con check la rutina de aseo personal realizada).
- SECCIÓN 3: FICHA DE AUTOEVALUACIÓN MOTRIZ Y CONVIVENCIA (Tabla con emoticones para autoevaluarse).
"""

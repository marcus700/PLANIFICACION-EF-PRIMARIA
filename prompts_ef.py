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
5. JUSTIFICACIÓN DE LA UNIDAD (Enfoque Implícito/Técnico): Redacta un párrafo formal y técnico que justifique la unidad. Explica que el área de Educación Física responde a [Mencionar el problema institucional] no de forma teórica o conceptual, sino a través del "Acondicionamiento Técnico-Motriz Operativo". Detalla cómo el desarrollo de las capacidades físicas (condicionales y coordinativas) dota al estudiante de las herramientas biológicas, metabólicas y biomecánicas necesarias para mitigar o adaptarse a dicha problemática desde la salud y el movimiento

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
- **Enfoque Dual (Estrategia Híbrida):** Redacta la situación significativa integrando la vinculación implícita y explícita.
- **Vinculación Implícita (Acción):** Describe el problema de contexto ({problema_contexto}) como una necesidad de mejora de capacidades motrices puras (coordinación, orientación, agilidad) que deben desarrollarse en el patio.
- **Vinculación Explícita (Reflexión):** Indica cómo, en momentos específicos de la unidad (inicio o cierre de sesión, diálogos reflexivos), se vinculará explícitamente la práctica motriz con la utilidad preventiva o de respuesta ante dicho contexto (ej. "¡Hoy movemos el cuerpo no solo para jugar, sino para saber cómo reaccionar rápido si la lluvia intensa nos obliga a desplazarnos con seguridad!").
- Incluye un dato cuantitativo o cualitativo sobre la problemática (ej. "se observa que solo el 30% logra orientarse adecuadamente en espacios dinámicos...").
- Plantea 3 preguntas retadoras/desafiantes, de las cuales al menos una debe ser de carácter puramente motriz y otra de carácter reflexivo-contextual.
- Propón la estrategia pedagógica (circuitos lúdico-motores, juegos de reglas, estaciones) para resolver el reto.

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

- **INICIO (Aprox. 20 min) - Enfoque Explícito:**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE, conectando directamente la actividad con el contexto.
  1. **Motivación ({tipo_motivacion}):** [Desarrolla la motivación vinculándola EXPLÍCITAMENTE con la problemática o contexto descrito en la unidad (ej. "¡Chicos, hoy iniciaremos el entrenamiento de superviviencia ante lluvias intensas simulando cruzar ríos!")].
  2. **Saberes previos:** [Preguntas sobre qué saben del movimiento técnico que veremos y su utilidad en la vida diaria].
  3. **Problematización / Conflicto cognitivo:** [Plantea el reto motriz explicando por qué es vital dominar esta técnica para prevenir riesgos en su entorno].
  4. **Propósito de la clase:** [Comunica qué aprenderán hoy y para qué les servirá en su contexto real].
  5. **Criterios de evaluación:** [Explica cómo evaluaremos su rigor técnico].
  6. **Acuerdos de convivencia:** [2 a 3 normas de seguridad para el patio enfocadas en el cuidado propio y del grupo].

- **DESARROLLO (Aprox. 60 min) - Vínculo Implícito Puro (Técnico):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE, centrado EXCLUSIVAMENTE en la biomecánica y el rigor técnico.
  1. **Activación Corporal (Calentamiento técnico):** [Describe movimientos articulares y desplazamientos específicos enfocados en preparar la musculatura que se usará en la parte principal].
  2. **Secuencia de Actividades Motrices (Rigor Técnico):** [Describe 3 actividades de alta exigencia biomecánica. Detalla la postura correcta, la ejecución técnica exacta (ej. "apoyo correcto de la planta del pie", "braceo alternado", "alineación corporal") y las correcciones pedagógicas precisas para lograr la maestría del movimiento. No menciones el contexto ambiental aquí, solo la técnica pura].
  3. **Hidratación:** [Pausa breve para recuperar].
  4. **ACTIVIDAD OBLIGATORIA DE ALTO NIVEL COGNITIVO (Evaluación Técnica):** [Reto donde los alumnos ejecutan la técnica y deben co-evaluar con un compañero el cumplimiento de los parámetros técnicos biomecánicos exactos (checklist de postura/ejecución)].

- **CIERRE (Aprox. 10 min) - Síntesis Dual y Metacognición:**
  Redactado OBLIGATORIAMENTE Y EN SU TOTALIDAD, uniendo ambos mundos.
  1. **Vuelta a la calma:** [Ejercicios de respiración y relajación muscular].
  2. **Metacognición de Síntesis Dual (OBLIGATORIO):** [Redacta explícitamente 3 preguntas que unan la técnica y el contexto: ¿Cómo te ayudó mantener una postura correcta (técnica) para desplazarte rápido y con seguridad en el juego (prevención)? ¿De qué te sirve controlar tu fuerza y equilibrio (dominio motriz) hoy en el patio para reaccionar ante un movimiento telúrico (crisis ambiental)? ¿Qué aprendiste hoy sobre tu cuerpo que te sirve para protegerte en la vida real?].
  3. **Rutina Obligatoria de Higiene Personal:** [Describe en detalle la práctica autónoma de aseo personal, lavado de manos con jabón, secado con toalla y cambio de polo deportivo].

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

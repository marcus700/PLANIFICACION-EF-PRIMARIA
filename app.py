import streamlit as st
from google import genai
from google.genai import types
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re

# Importar la Base de Datos CNEB directamente del archivo cneb_datos.py
try:
    from cneb_datos import CNEB_PRIMARIA, obtener_ciclo_primaria
except Exception as e:
    st.error(f"Error al cargar cneb_datos.py: {e}. Asegúrate de que el archivo se llame 'cneb_datos.py' sin punto al final.")

# Configuración visual de la plataforma
st.set_page_config(page_title="PlanificaEF", page_icon="🏃‍♂️", layout="centered")

st.title("🏃‍♂️ PlanificaEF")
st.subheader("Asistente Pedagógico de Educación Física (Primaria - CNEB)")
st.write("Herramienta inteligente para diseñar tus documentos curriculares al instante.")

# Enlace automático a la clave secreta guardada de forma segura
api_key = st.secrets.get("GEMINI_API_KEY", None)

if not api_key:
    st.error("⚠️ No se encontró la GEMINI_API_KEY en los secretos de Streamlit.")

# ==============================================================================
# FUNCIÓN AVANZADA: CONVERTIDOR PROFESIONAL DE MARKDOWN A TABLAS Y FORMATO WORD
# ==============================================================================
def set_cell_background(cell, fill_color):
    """Aplica color de fondo a una celda de tabla en Word."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def crear_archivo_word_profesional(texto_markdown):
    doc = Document()
    
    # Configurar márgenes de página (0.75 pulgadas)
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    lineas = texto_markdown.split('\n')
    i = 0
    
    while i < len(lineas):
        linea = lineas[i].strip()
        
        if not linea:
            i += 1
            continue

        # Encabezado Nivel 1 (#)
        if linea.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(linea.replace('# ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(15)
            run.font.bold = True
            run.font.color.rgb = RGBColor(27, 94, 32)
            i += 1
        # Encabezado Nivel 2 (##)
        elif linea.startswith('## '):
            p = doc.add_paragraph()
            run = p.add_run(linea.replace('## ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(46, 125, 50)
            i += 1
        # Encabezado Nivel 3 (###)
        elif linea.startswith('### '):
            p = doc.add_paragraph()
            run = p.add_run(linea.replace('### ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(10.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(27, 94, 32)
            i += 1

        # Tablas Markdown (| Col1 | Col2 |)
        elif linea.startswith('|'):
            filas_tabla = []
            while i < len(lineas) and lineas[i].strip().startswith('|'):
                l = lineas[i].strip()
                if not re.match(r'^\|[\s\:\-]+\|', l):
                    columnas = [c.strip() for c in l.split('|')[1:-1]]
                    if columnas:
                        filas_tabla.append(columnas)
                i += 1

            if filas_tabla:
                num_filas = len(filas_tabla)
                num_cols = max(len(r) for r in filas_tabla)
                table = doc.add_table(rows=num_filas, cols=num_cols)
                table.style = 'Table Grid'
                
                for r_idx, row_data in enumerate(filas_tabla):
                    for c_idx, cell_value in enumerate(row_data):
                        if c_idx < num_cols:
                            cell = table.cell(r_idx, c_idx)
                            p = cell.paragraphs[0]
                            p.text = ""
                            
                            # Formatear texto en negrita dentro de celdas
                            partes = re.split(r'(\*\*.*?\*\*)', cell_value)
                            for parte in partes:
                                if parte.startswith('**') and parte.endswith('**'):
                                    run = p.add_run(parte[2:-2])
                                    run.bold = True
                                else:
                                    run = p.add_run(parte)
                                run.font.name = 'Arial'
                                if r_idx == 0:
                                    run.font.bold = True
                                    run.font.color.rgb = RGBColor(255, 255, 255)
                                    run.font.size = Pt(9.5)
                                else:
                                    run.font.size = Pt(8.5)
                                    
                            # Fondo verde institucional para la primera fila (Encabezado)
                            if r_idx == 0:
                                set_cell_background(cell, "2E7D32")
                                
                doc.add_paragraph()
        else:
            p = doc.add_paragraph()
            partes = re.split(r'(\*\*.*?\*\*)', linea)
            for parte in partes:
                if parte.startswith('**') and parte.endswith('**'):
                    run = p.add_run(parte[2:-2])
                    run.bold = True
                else:
                    run = p.add_run(parte)
                run.font.name = 'Arial'
                run.font.size = Pt(10)
            i += 1

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Función inteligente que detecta dinámicamente los modelos habilitados en tu cuenta
def generar_respuesta_ia(client, system_instruction, prompt_usuario):
    modelos_disponibles = []
    try:
        for m in client.models.list():
            nombre = m.name.replace('models/', '')
            modelos_disponibles.append(nombre)
    except Exception:
        pass

    if not modelos_disponibles:
        modelos_disponibles = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash-latest']

    ultimo_error = None
    for modelo in modelos_disponibles:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=prompt_usuario,
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    temperature=0.7
                )
            )
            if response and response.text:
                return response.text
        except Exception as err:
            ultimo_error = err
            continue

    raise RuntimeError(f"Error al conectar con la API de Google: {ultimo_error}")

# Creación de las 3 Pestañas de Trabajo
tab1, tab2, tab3 = st.tabs([
    "📂 Crear Unidad de Aprendizaje", 
    "📝 Crear Sesión de Aprendizaje", 
    "📊 Crear Rúbrica de Evaluación"
])

# --- PESTAÑA 1: UNIDADES ---
with tab1:
    st.write("Estructura una unidad didáctica completa según el formato oficial MINEDU.")
    with st.form("form_unidad"):
        grado_u = st.selectbox("Grado de Primaria:", ["1° Grado", "2° Grado", "3° Grado", "4° Grado", "5° Grado", "6° Grado"], key="u1")
        duracion_u = st.selectbox("Duración de la Unidad:", ["4 Semanas", "5 Semanas", "6 Semanas", "8 Semanas"], key="u2")
        problema_u = st.text_area("Describe el problema del contexto o interés de los niños:", placeholder="Ej. Los estudiantes muestran dificultades para trabajar en equipo y respetar reglas en los juegos del recreo.", key="u3")
        boton_unidad = st.form_submit_button("📂 Generar Unidad en Word")

    if boton_unidad and problema_u:
        with st.spinner("Cargando matriz de cneb_datos.py y diseñando la Unidad..."):
            try:
                client = genai.Client(api_key=api_key)
                ciclo_u = obtener_ciclo_primaria(grado_u)
                
                # Extraer matriz de desempeñs CNEB para este grado desde cneb_datos.py
                matriz_desempenos_u = ""
                if CNEB_PRIMARIA:
                    des_u_list = []
                    for comp_name, comp_data in CNEB_PRIMARIA.items():
                        des_list = comp_data.get("desempenos", {}).get(grado_u, [])
                        if des_list:
                            des_u_list.append(f"COMPETENCIA: {comp_name}")
                            des_u_list.extend(des_list)
                    matriz_desempenos_u = "\n".join(des_u_list)

                instrucciones_u = f"""Actúa como un Especialista Curricular experto en Educación Física para Primaria bajo el enfoque del CNEB de Perú (MINEDU).
Diseña una UNIDAD DE APRENDIZAJE completa estructurada EN TABLAS MARKDOWN con la siguiente estructura oficial:

REGISTRO CURRICULAR Y BASE DE DATOS CNEB_DATOS.PY PARA {grado_u} ({ciclo_u}):
{matriz_desempenos_u}

1. DATOS INFORMATIVOS:
Genera una TABLA bien organizada con los campos completados con puntos [.....] para rellenar:
| Campo | Detalle |
| DRE / UGEL | DRE [.....] / UGEL [.....] |
| Institución Educativa | I.E. N° [.....] |
| Lugar / Localidad | [.....] |
| Ciclo | {ciclo_u} |
| Grado y Sección | {grado_u}, Secciones: [.....] |
| Docente del Área | [.....] |
| Director(a) | [.....] |
| Duración y Periodo | {duracion_u} (Del [Día/Mes] al [Día/Mes]) |

2. TÍTULO DE LA UNIDAD DE APRENDIZAJE (Significativo, retador e innovador).

3. SITUACIÓN SIGNIFICATIVA (Contexto del problema, Reto en pregunta y Producto de la unidad).

4. PROPÓSITOS DE APRENDIZAJE Y SECUENCIA DE SESIONES:
Genera una TABLA en formato Markdown de 7 COLUMNAS:
| ACTIVIDAD (SESIÓN) | DESCRIPCIÓN PEDAGÓGICA | COMPETENCIA / CAPACIDADES | ESTÁNDAR DE LA COMPETENCIA | DESEMPEÑO PRECISADO | CRITERIOS DE EVALUACIÓN | INSTRUMENTO DE EVALUACIÓN |

REGLAS ABSOLUTAS DE TRANSCRIPCIÓN LITERAL DEL CNEB:
- En 'ESTÁNDAR DE LA COMPETENCIA': Copia y transcribe la redacción LITERAL, EXACTA Y COMPLETA del Estándar CNEB oficial del {ciclo_u}. Queda PROHIBIDO refrasear o resumir. Únicamente **RESALTA EN NEGRITA (**texto**)** la frase original que se ejercita en esa sesión.
- En 'DESEMPEÑO PRECISADO': Usa los desempeños oficiales del CNEB de {grado_u} provistos arriba en la matriz CNEB. Conserva 100% las palabras originales del CNEB y **RESALTA EN NEGRITA (**texto en negrita**)** ÚNICAMENTE DOS PARTES:
  1. La frase tomada del desempeño original que se ejercita.
  2. Lo que le agregas al final para precisarlo con el tema de la sesión.
- En 'CRITERIOS DE EVALUACIÓN': Formula OBLIGATORIAMENTE EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN por cada sesión (Acción + Contenido + Condición) sin etiquetas explícitas '(Acción)'.
- En 'INSTRUMENTO DE EVALUACIÓN': Lista de cotejo / Rúbrica analítica.

5. ENFOQUES TRANSVERSALES PRIORIZADOS:
Genera una TABLA con las columnas:
| Enfoque Transversal Priorizado | Valores | Actitudes / Comportamientos Observables |

6. MATERIALES Y RECURSOS DIDÁCTICOS."""

                pedido_u = f"Crea una unidad para {grado_u} con duración de {duracion_u}. Contexto del problema: {problema_u}"
                
                resultado_u = generar_respuesta_ia(client, instrucciones_u, pedido_u)
                
                st.success("¡Unidad Curricular generada con éxito extrayendo cneb_datos.py!")
                st.markdown(resultado_u)
                
                archivo_word_u = crear_archivo_word_profesional(resultado_u)
                st.download_button(
                    label="📥 Descargar Unidad en Word (.docx)", 
                    data=archivo_word_u, 
                    file_name=f"Unidad_PlanificaEF_{grado_u.replace(' ', '_')}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error al generar la Unidad: {e}")

# --- PESTAÑA 2: SESIONES ---
with tab2:
    st.write("Genera el desarrollo de una sesión diaria paso a paso extrayendo datos reales del CNEB.")
    with st.form("form_sesion"):
        grado_s = st.selectbox("Grado de Primaria:", ["1° Grado", "2° Grado", "3° Grado", "4° Grado", "5° Grado", "6° Grado"], key="s1")
        competencia_s = st.selectbox("Competencia Principal:", ["Se desenvuelve de manera autónoma a través de su motricidad", "Asume una vida saludable", "Interactúa a través de sus habilidades sociomotrices"], key="s2")
        detalles_s = st.text_area("Tema de la clase o materiales disponibles:", placeholder="Ej. Coordinación óculo-manual lanzando y recibiendo pelotas de plástico.", key="s3")
        boton_sesion = st.form_submit_button("⚡ Generar Sesión en Word")

    if boton_sesion and detalles_s:
        with st.spinner("Extrayendo desempeño base de cneb_datos.py y diseñando sesión..."):
            try:
                client = genai.Client(api_key=api_key)
                ciclo_s = obtener_ciclo_primaria(grado_s)
                
                # Extraer Estándar y Desempeños reales del archivo cneb_datos.py
                estandar_base = ""
                desempenos_base = ""
                if CNEB_PRIMARIA and competencia_s in CNEB_PRIMARIA:
                    estandar_base = CNEB_PRIMARIA[competencia_s]["estandares"].get(ciclo_s, "")
                    desempenos_lista = CNEB_PRIMARIA[competencia_s]["desempenos"].get(grado_s, [])
                    desempenos_base = "\n".join(desempenos_lista)

                instrucciones = f"""Actúa como un docente experto de Educación Física de nivel Primaria en Perú, especialista en el enfoque por competencias del Currículo Nacional de la Educación Básica (CNEB) de MINEDU.

DATOS OFICIALES EXTRAÍDOS DIRECTAMENTE DE CNEB_DATOS.PY:
- Grado y Ciclo: {grado_s} ({ciclo_s})
- Competencia principal: {competencia_s}
- Tema / Propósito motriz: {detalles_s}
- ESTÁNDAR CNEB OFICIAL A TRANSCRIBIR DE MANERA COMPLETA Y LITERAL: "{estandar_base}"
- DESEMPEÑOS CNEB OFICIALES DISPONIBLES PARA {grado_s} Y PARA LA COMPETENCIA "{competencia_s}":
{desempenos_base}

ESTRUCTURA OBLIGATORIA A GENERAR EN FORMATO MARKDOWN:

# SESIÓN DE APRENDIZAJE N°.......
**Título:** [Crea un título motivador, lúdico e innovador sobre {detalles_s}]

## 1. DATOS INFORMATIVOS
Genera una TABLA de 2 columnas:
| Campo | Detalle |
| DRE / UGEL | DRE [.....] / UGEL [.....] |
| Institución Educativa | I.E. N° [.....] |
| Lugar / Localidad | [.....] |
| Docente | [.....] |
| Grado y Sección | {grado_s}, Secciones: [.....] |
| Área | Educación Física |
| Fecha y Duración | Fecha: [.....] | Duración: 90 minutos |

## 2. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE
Genera una TABLA con exactamente las siguientes 6 columnas:
| Competencia y Capacidades | Estándar CNEB | Desempeños Precisados | Criterios de Evaluación | Evidencia y Producto | Instrumento de Evaluación |

REGLAS ABSOLUTAS DE TRANSCRIPCIÓN Y PRECISIÓN DEL CNEB:
- **Columna 1:** Transcribe la competencia ({competencia_s}) y sus capacidades oficiales.
- **Columna 2:** Transcribe EXACTAMENTE Y DE MANERA COMPLETA el estándar oficial proporcionado arriba ("{estandar_base}"). PROHIBIDO refrasear o usar puntos suspensivos '...'. **Resalta en negrita (**texto**)** únicamente la frase del estándar que se aplica directamente hoy.
- **Columna 3 (DESEMPEÑO PRECISADO):** Copia el desempeño oficial del CNEB correspondiente de la lista de {grado_s} arriba provista. CONSERVA EL TEXTO BASE DEL CNEB Y **RESALTA EN NEGRITA (**texto en negrita**)** ÚNICAMENTE DOS PARTES:
  1. La frase o acción tomada del desempeño original del CNEB que se ejercita hoy.
  2. La adición específica que le agregas al final para precisarlo con el tema de la clase ({detalles_s}).
- **Columna 4:** Formula OBLIGATORIAMENTE EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN claros (Acción + Contenido + Condición) sin etiquetas explícitas '(Acción)'.
- **Columna 5:** Define la Evidencia de aprendizaje (producto o actuación medible).
- **Columna 6:** Lista de cotejo.

## 3. ENFOQUE TRANSVERSAL
Genera una TABLA con las columnas:
| Enfoque Transversal Priorizado | Valor | Actitud / Comportamiento Observable |

## 4. PREPARACIÓN DE LA SESIÓN
Genera una TABLA estrictamente de 2 columnas:
| ¿Qué necesitamos hacer antes de la sesión? | Recursos o Materiales a utilizar |

## 5. SECUENCIA DIDÁCTICA (MOMENTOS DE LA SESIÓN)
REGLA FUNDAMENTAL DE CNEB_DATOS.PY: Redacta TODAS las acciones de los momentos en **PRIMERA PERSONA** ("Recibo a mis estudiantes...", "Explico el juego...", "Organizo a las cuadrillas...") y en **TIEMPO PRESENTE**.

### A) INICIO (Aprox. 20% del tiempo - 18 min):
- **Motivación inicial:** Una historia corta, imagen o desafío relacionado con {detalles_s}.
- **Recojo de saberes previos:** Preguntas abiertas sobre el tema.
- **Problematización / Conflicto cognitivo:** Reto motriz o pregunta que despierte la curiosidad.
- **Propósito y organización:** Comunicar claramente qué van a aprender hoy en {grado_s}.
- **Acuerdos de convivencia:** 2 a 3 acuerdos para el campo o patio.
- **Activación Corporal (Calentamiento dinámico):** Juego motivador relacionado al tema, movilidad articular y TOMA DE PULSO INICIAL.

### B) DESARROLLO (Aprox. 60% del tiempo - 54 min) - Gestión y acompañamiento:
- Diseña una secuencia metodológica de lo simple a lo complejo (progresión motriz adecuada para {grado_s}).
- Incluye de 3 a 4 actividades prácticas explicadas con claridad para ejecutar en el patio (juegos tradicionales, circuitos, minitorneos o dinámicas de exploración).
- Incluye pausa de HIDRATACIÓN y REGLAS DE SEGURIDAD para evitar accidentes.
- Asegúrate de que las actividades promuevan la autonomía, el pensamiento estratégico y la interacción saludable.
- Describe la estrategia de retroalimentación (feedback) que brindo como docente durante la práctica.

### C) CIERRE (Aprox. 20% del tiempo - 18 min):
- **Actividad de Vuelta a la Calma:** Juegos de baja intensidad, estiramientos, ejercicios de respiración o relajación y TOMA DE PULSO FINAL.
- **Metacognición:** Preguntas de reflexión (¿Qué aprendimos hoy? ¿Cómo lo logramos? ¿En qué tuvimos dificultad? ¿Para qué nos sirve?).
- **Cuidado e Higiene Personal:** Hábitos de lavado de manos, cara, hidratación y orden del material recolectado.

## 6. ANEXO: INSTRUMENTO DE EVALUACIÓN
Diseña una TABLA de **Lista de Cotejo** con los 3 criterios de evaluación planteados al inicio y filas con espacio para los nombres de los estudiantes."""

                pedido = f"Diseña una sesión para {grado_s}. Competencia: {competencia_s}. Detalles del tema: {detalles_s}"
                
                resultado_s = generar_respuesta_ia(client, instrucciones, pedido)
                
                st.success("¡Sesión generada extrayendo desempeño exacto de cneb_datos.py!")
                st.markdown(resultado_s)
                
                archivo_word = crear_archivo_word_profesional(resultado_s)
                st.download_button(
                    label="📥 Descargar Sesión en Word (.docx)", 
                    data=archivo_word, 
                    file_name=f"Sesion_PlanificaEF_{grado_s.replace(' ', '_')}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error al generar la Sesión: {e}")

# --- PESTAÑA 3: RÚBRICAS ---
with tab3:
    st.write("Diseña instrumentos de evaluación con criterios claros.")
    with st.form("form_rubrica"):
        grado_r = st.selectbox("Grado de Primaria:", ["1° Grado", "2° Grado", "3° Grado", "4° Grado", "5° Grado", "6° Grado"], key="r1")
        competencia_r = st.selectbox("Competencia a Evaluar:", ["Se desenvuelve de manera autónoma a través de su motricidad", "Asume una vida saludable", "Interactúa a través de sus habilidades sociomotrices"], key="r2")
        criterio_r = st.text_input("Desempeño específico a evaluar:", placeholder="Ej. Control de la postura al saltar con un pie.")
        boton_rubrica = st.form_submit_button("📊 Generar Rúbrica en Word")

    if boton_rubrica and criterio_r:
        with st.spinner("Estructurando la rúbrica..."):
            try:
                client = genai.Client(api_key=api_key)
                instrucciones_r = f"""Actúa como un Evaluador Pedagógico experto en Educación Física para Primaria.
Diseña una rúbrica analítica estructurada con los niveles: En Inicio, En Proceso, Logrado y Logro Destacado para el desempeño solicitado, utilizando exactamente 3 criterios claros y observables alineados al CNEB sin etiquetar explícitamente '(Acción)' ni '(Contenido)'."""

                pedido_r = f"Crea una rúbrica para {grado_r}. Competencia: {competencia_r}. Desempeño: {criterio_r}"
                
                resultado_r = generar_respuesta_ia(client, instrucciones_r, pedido_r)
                
                st.success("¡Rúbrica generada con éxito!")
                st.markdown(resultado_r)
                
                archivo_word_r = crear_archivo_word_profesional(resultado_r)
                st.download_button(
                    label="📥 Descargar Rúbrica en Word (.docx)", 
                    data=archivo_word_r, 
                    file_name=f"Rubrica_PlanificaEF_{grado_r.replace(' ', '_')}.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Error al generar la Rúbrica: {e}")

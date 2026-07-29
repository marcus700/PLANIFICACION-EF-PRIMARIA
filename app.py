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
        with st.spinner("Diseñando la Unidad de Aprendizaje en tablas formateadas..."):
            try:
                client = genai.Client(api_key=api_key)
                instrucciones_u = (
                    "Actúa como un Especialista Curricular experto en Educación Física para Primaria bajo el enfoque del CNEB de Perú (MINEDU).\n"
                    "Diseña una UNIDAD DE APRENDIZAJE completa estructurada EN TABLAS MARKDOWN con la siguiente estructura oficial:\n\n"
                    "1. DATOS INFORMATIVOS:\n"
                    "Genera una TABLA bien organizada con los campos completados con puntos [.....] para rellenar:\n"
                    "| Campo | Detalle |\n"
                    "| DRE / UGEL | DRE [.....] / UGEL [.....] |\n"
                    "| Institución Educativa | I.E. N° [.....] |\n"
                    "| Lugar / Localidad | [.....] |\n"
                    "| Ciclo | Indicar automáticamente según el grado (III Ciclo para 1° y 2°; IV Ciclo para 3° y 4°; V Ciclo para 5° y 6°) |\n"
                    "| Grado y Sección | [Grado seleccionado], Secciones: [.....] |\n"
                    "| Docente del Área | [.....] |\n"
                    "| Director(a) | [.....] |\n"
                    "| Duración y Periodo | [Duración seleccionada] (Del [Día/Mes] al [Día/Mes]) |\n\n"
                    "2. TÍTULO DE LA UNIDAD DE APRENDIZAJE (Significativo, retador e innovador).\n\n"
                    "3. SITUACIÓN SIGNIFICATIVA (Contexto del problema, Reto en pregunta e Y Producto de la unidad).\n\n"
                    "4. PROPÓSITOS DE APRENDIZAJE Y SECUENCIA DE SESIONES:\n"
                    "Genera una TABLA en formato Markdown de 7 COLUMNAS:\n"
                    "| ACTIVIDAD (SESIÓN) | DESCRIPCIÓN PEDAGÓGICA | COMPETENCIA / CAPACIDADES | ESTÁNDAR DE LA COMPETENCIA | DESEMPEÑO PRECISADO | CRITERIOS DE EVALUACIÓN | INSTRUMENTO DE EVALUACIÓN |\n\n"
                    "REGLAS ESTRUCTURALES OBLIGATORIAS PARA ESTA TABLA:\n"
                    "- Cada fila corresponde a una Sesión de la unidad.\n"
                    "- En 'ESTÁNDAR DE LA COMPETENCIA': Transcribe el estándar oficial del CNEB y **RESALTA EN NEGRITA (**texto**)** la parte que se enfatiza en la sesión.\n"
                    "- En 'DESEMPEÑO PRECISADO': Transcribe el desempeño OFICIAL del CNEB del grado y **RESALTA EN NEGRITA (**texto precisado**)** la parte adaptada para el tema.\n"
                    "- En 'CRITERIOS DE EVALUACIÓN': Formula OBLIGATORIAMENTE ENTRE 2 Y 3 CRITERIOS DE EVALUACIÓN por cada sesión (Estructura: Acción + Contenido + Condición).\n"
                    "- En 'INSTRUMENTO DE EVALUACIÓN': Lista de cotejo / Rúbrica analítica.\n\n"
                    "5. ENFOQUES TRANSVERSALES PRIORIZADOS:\n"
                    "Genera una TABLA con las columnas:\n"
                    "| Enfoque Transversal Priorizado | Valores | Actitudes / Comportamientos Observables |\n\n"
                    "6. MATERIALES Y RECURSOS DIDÁCTICOS."
                )
                pedido_u = f"Crea una unidad para {grado_u} con duración de {duracion_u}. Contexto del problema: {problema_u}"
                
                resultado_u = generar_respuesta_ia(client, instrucciones_u, pedido_u)
                
                st.success("¡Unidad Curricular generada con éxito!")
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
    st.write("Genera el desarrollo de una sesión diaria paso a paso con tablas estructuradas.")
    with st.form("form_sesion"):
        grado_s = st.selectbox("Grado de Primaria:", ["1° Grado", "2° Grado", "3° Grado", "4° Grado", "5° Grado", "6° Grado"], key="s1")
        competencia_s = st.selectbox("Competencia Principal:", ["Se desenvuelve de manera autónoma a través de su motricidad", "Asume una vida saludable", "Interactúa a través de sus habilidades sociomotrices"], key="s2")
        detalles_s = st.text_area("Tema de la clase o materiales disponibles:", pla

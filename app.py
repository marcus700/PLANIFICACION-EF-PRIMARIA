import io
import re
import time
from PIL import Image
import docx
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
import streamlit as st

# ==============================================================================
# IMPORTACIÓN DESDE TU BASE DE DATOS EXTERNA (cneb_datos.py)
# ==============================================================================
from cneb_datos import CNEB_PRIMARIA, obtener_ciclo_primaria

def normalizar_grado_cneb(grado_str: str) -> str:
    """Mapea la opción seleccionada al formato de llave exacta en cneb_datos.py"""
    if "1" in grado_str: return "1° de Primaria"
    if "2" in grado_str: return "2° de Primaria"
    if "3" in grado_str: return "3° de Primaria"
    if "4" in grado_str: return "4° de Primaria"
    if "5" in grado_str: return "5° de Primaria"
    if "6" in grado_str: return "6° de Primaria"
    return "2° de Primaria"

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA Y CSS MEJORADO PARA ALTA VISIBILIDAD DE BOTONES
# ==============================================================================
st.set_page_config(
    page_title="PlanificaEF Primaria - Plataforma de Educación Física",
    page_icon="⚽",
    layout="wide"
)

st.markdown("""
<style>
    header, footer, [data-testid="stHeader"], [data-testid="stDecoration"], 
    [data-testid="stStatusWidget"], [data-testid="stViewerBadge"], 
    [data-testid="manage-app-button"], .stAppDeployButton, .viewerBadge_container__1613n {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
    }
    
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
    }
    
    .block-container {
        padding-top: 1.2rem !important;
        padding-bottom: 1rem !important;
    }
    
    .main-header {
        font-size: 2.2rem;
        color: #1E3A8A;
        text-align: center;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.05rem;
        color: #4B5563;
        text-align: center;
        margin-bottom: 1.5rem;
    }

    .stTextInput input, .stTextArea textarea, .stSelectbox [data-baseweb="select"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border: 1px solid #94A3B8 !important;
        border-radius: 8px !important;
    }

    /* ESTILOS DE ALTO CONTRASTE Y LEGIBILIDAD PARA BOTONES DE HERRAMIENTAS */
    div.st-key-btn_unidad > button {
        background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%) !important;
        background-color: #7C3AED !important;
        border: 2px solid #6D28D9 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 14px rgba(124, 58, 237, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    div.st-key-btn_unidad > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(124, 58, 237, 0.6) !important;
    }

    div.st-key-btn_proyecto > button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%) !important;
        background-color: #059669 !important;
        border: 2px solid #047857 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 14px rgba(5, 150, 105, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    div.st-key-btn_proyecto > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(5, 150, 105, 0.6) !important;
    }

    div.st-key-btn_sesion > button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        background-color: #2563EB !important;
        border: 2px solid #1D4ED8 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    div.st-key-btn_sesion > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.6) !important;
    }

    div.stButton > button:not([key="btn_unidad"]):not([key="btn_proyecto"]):not([key="btn_sesion"]) {
        background: linear-gradient(135deg, #1E40AF 0%, #1E3A8A 100%) !important;
        background-color: #1E40AF !important;
        border: 2px solid #1D4ED8 !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.5rem !important;
        box-shadow: 0 4px 16px rgba(30, 64, 175, 0.5) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:not([key="btn_unidad"]):not([key="btn_proyecto"]):not([key="btn_sesion"]):hover {
        background: linear-gradient(135deg, #2563EB 0%, #1E40AF 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.7) !important;
    }

    div.stButton > button,
    div.stButton > button *,
    div.stButton > button p,
    div.stButton > button span,
    div.stButton > button div {
        color: #FFFFFF !important;
        font-weight: 800 !important;
        font-size: 1.05rem !important;
        text-shadow: 0px 1px 3px rgba(0, 0, 0, 0.6) !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header">⚽ PlanificaEF - Plataforma de Educación Física</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Sistema de Planificación Curricular Especializado en Educación Física Primaria (CNEB - MINEDU)</div>', unsafe_allow_html=True)

# ==============================================================================
# CONTROL DE ACCESO
# ==============================================================================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔒 Acceso Restringido")
        pwd_input = st.text_input("Contraseña de acceso:", type="password", key="pwd_input")
        if st.button("Ingresar 🚀"):
            target_pwd = st.secrets.get("APP_PASSWORD", "docente2026ef")
            if pwd_input == target_pwd:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta.")
    return False

if not check_password():
    st.stop()

# MEMORIA PERSISTENTE
if 'resultado_md' not in st.session_state:
    st.session_state['resultado_md'] = None
if 'tipo_doc_generado' not in st.session_state:
    st.session_state['tipo_doc_generado'] = None
if 'fname_clean' not in st.session_state:
    st.session_state['fname_clean'] = None
if 'tipo_documento' not in st.session_state:
    st.session_state['tipo_documento'] = "Unidad de Aprendizaje"

# SIDEBAR CON MODELOS OFICIALES Y ACTIVOS DE GOOGLE STUDIO
st.sidebar.title("⚙️ Configuración EF")
if st.sidebar.button("🔒 Cerrar Sesión"):
    st.session_state["password_correct"] = False
    st.rerun()

st.sidebar.markdown("---")
if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 API Key activada.")
else:
    api_key = st.sidebar.text_input("🔑 Google AI Studio API Key:", type="password")

# MODELOS OFICIALES Y ACTIVOS
model_choice = st.sidebar.selectbox(
    "Modelo de Gemini:", 
    ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-2.5-pro", "gemini-3.5-flash", "gemini-3.6-flash"]
)

# ==============================================================================
# HERRAMIENTAS DE EDUCACIÓN FÍSICA
# ==============================================================================
st.markdown("### 📋 Selecciona el Documento de Educación Física a Elaborar:")

col_b1, col_b2, col_b3 = st.columns(3)

with col_b1:
    if st.button("📘 Unidad de Aprendizaje", key="btn_unidad", use_container_width=True):
        st.session_state['tipo_documento'] = "Unidad de Aprendizaje"
        st.rerun()

with col_b2:
    if st.button("🚀 Proyecto de Aprendizaje", key="btn_proyecto", use_container_width=True):
        st.session_state['tipo_documento'] = "Proyecto de Aprendizaje"
        st.rerun()

with col_b3:
    if st.button("🏃 Sesión de Aprendizaje de Ed. Física", key="btn_sesion", use_container_width=True):
        st.session_state['tipo_documento'] = "Sesión de Aprendizaje de Ed. Física"
        st.rerun()

tipo_documento = st.session_state['tipo_documento']

COLOR_MAP = {
    "Unidad de Aprendizaje": "#7C3AED",
    "Proyecto de Aprendizaje": "#059669",
    "Sesión de Aprendizaje de Ed. Física": "#2563EB"
}
banner_color = COLOR_MAP.get(tipo_documento, "#7C3AED")

st.markdown(f"""
<div style="background-color: {banner_color}; color: white; padding: 0.6rem 1rem; border-radius: 8px; font-weight: bold; font-size: 1.1rem; margin-top: 0.8rem; margin-bottom: 1.2rem; text-align: center; text-shadow: 0px 1px 3px rgba(0,0,0,0.4);">
    📍 Área Exclusiva: EDUCACIÓN FÍSICA | Herramienta: {tipo_documento.upper()}
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# CONVERTIDOR DE MARKDOWN A WORD CON TABLAS EN TONOS PASTELES
# ==============================================================================
def add_formatted_text(paragraph, text):
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = paragraph.add_run(part[2:-2])
            run.font.bold = True
        else:
            paragraph.add_run(part)

def markdown_to_docx(md_text, ie_nombre="I.E. N° 22314", es_horizontal=False):
    doc = docx.Document()
    PASTEL_COLORS = ['D9E1F2', 'E2EFDA', 'FFF2CC', 'E8D8F8', 'E0F2FE', 'FCE4D6']
    table_count = 0
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        if es_horizontal:
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Inches(11.69)
            section.page_height = Inches(8.27)
        else:
            section.orientation = WD_ORIENT.PORTRAIT
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)

    lines = md_text.split('\n')
    in_table = False
    table_data = []

    def render_table(t_data, color_hex):
        rows = len(t_data)
        cols = max(len(r) for r in t_data) if rows > 0 else 0
        if rows > 0 and cols > 0:
            t = doc.add_table(rows=rows, cols=cols)
            t.style = 'Table Grid'
            for r_idx, row_cells in enumerate(t_data):
                for c_idx, cell_value in enumerate(row_cells):
                    if c_idx < cols:
                        cell = t.cell(r_idx, c_idx)
                        p_cell = cell.paragraphs[0]
                        p_cell.text = ""
                        add_formatted_text(p_cell, cell_value)
                        
                        if r_idx == 0:
                            shading_elm = OxmlElement('w:shd')
                            shading_elm.set(qn('w:val'), 'clear')
                            shading_elm.set(qn('w:color'), 'auto')
                            shading_elm.set(qn('w:fill'), color_hex)
                            cell._tc.get_or_add_tcPr().append(shading_elm)
                            for paragraph in cell.paragraphs:
                                for run in paragraph.runs:
                                    run.font.color.rgb = RGBColor(30, 58, 138)
                                    run.font.bold = True

    for line in lines:
        line_str = line.strip()
        line_str = re.sub(r'<br\s*/?>', ' ', line_str)
        
        if line_str.startswith('|') and line_str.endswith('|'):
            in_table = True
            if re.match(r'^\|[\s\:\-\|]+\|$', line_str):
                continue
            cells = [c.strip() for c in line_str.split('|')[1:-1]]
            table_data.append(cells)
            continue
        elif in_table:
            if table_data:
                table_count += 1
                header_color = PASTEL_COLORS[(table_count - 1) % len(PASTEL_COLORS)]
                render_table(table_data, header_color)
            in_table = False
            table_data = []

        heading_match = re.match(r'^(#{1,6})\s*(.*)$', line_str)
        if heading_match:
            hashes = heading_match.group(1)
            title_text = heading_match.group(2).strip()
            level = len(hashes)
            
            p = doc.add_paragraph()
            if level in [1, 2]:
                run = p.add_run(title_text.replace('**', ''))
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 58, 138)
            elif level in [3, 4]:
                run = p.add_run(title_text.replace('**', ''))
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 58, 138)
            else:
                add_formatted_text(p, title_text)
            continue

        if line_str.startswith('• ') or line_str.startswith('- '):
            p = doc.add_paragraph(style='List Bullet')
            clean_bullet = line_str[2:].strip()
            add_formatted_text(p, clean_bullet)
        elif line_str != "":
            p = doc.add_paragraph()
            add_formatted_text(p, line_str)

    if in_table and table_data:
        table_count += 1
        header_color = PASTEL_COLORS[(table_count - 1) % len(PASTEL_COLORS)]
        render_table(table_data, header_color)
        in_table = False
        table_data = []
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==============================================================================
# FORMULARIO DE DATOS
# ==============================================================================
st.subheader(f"📝 Configuración de Datos para Educación Física: {tipo_documento}")

c1, c2, c3 = st.columns(3)
with c1:
    dre_ugel = st.text_input("DRE / UGEL:", "Ica / Ica")
    ie_nombre = st.text_input("Institución Educativa:", "N.° 22314 'Vicenta Aquije de Huamán'")
with c2:
    director = st.text_input("Directora:", "Prof. Luisa Ruth Aronés Herrera")
    docente = st.text_input("Docente de Educación Física:", "Mario A. García Torres")
with c3:
    grado_seccion = st.selectbox("Grado y Sección:", ["1er Grado A", "2do Grado A", "3er Grado A", "4to Grado A", "5to Grado A", "6to Grado A"], index=1)
    
    # Detección del Ciclo leyendo desde cneb_datos.py
    grado_normalizado_cneb = normalizar_grado_cneb(grado_seccion)
    ciclo_actual = obtener_ciclo_primaria(grado_normalizado_cneb)
    st.info(f"Ciclo CNEB Detectado: **{ciclo_actual}**")

# VARIABLES ESPECÍFICAS PARA CADA HERRAMIENTA
if tipo_documento == "Sesión de Aprendizaje de Ed. Física":
    f1, f2, f3 = st.columns(3)
    with f1:
        num_doc = st.text_input("N.° de Sesión:", "01")
    with f2:
        fecha_sugerida = st.text_input("Fecha:", "22 de junio de 2026")
    with f3:
        duracion_sesion = st.selectbox("Duración de la Clase:", ["45 minutos", "90 minutos", "135 minutos"], index=1)
    
    st.markdown("##### 📌 Configuración Pedagógica de la Sesión (Opcional: Dejar en blanco para generación automática del CNEB)")
    titulo_sesion_input = st.text_input("Título de la Actividad / Sesión de Clase (Opcional):", value="", placeholder="Ej. Leemos señales para desplazarnos y reconocer direcciones en el patio")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        comps_seleccionadas = st.multiselect(
            "Competencia(s) a Trabajar (1, 2 o 3):",
            [
                "Se desenvuelve de manera autónoma a través de su motricidad",
                "Asume una vida saludable",
                "Interactúa a través de sus habilidades sociomotrices"
            ],
            default=["Se desenvuelve de manera autónoma a través de su motricidad"]
        )
        capacidades_custom = st.text_input("Capacidades Específicas (Opcional - Blanco para automático):", value="", placeholder="Ej. Comprende su cuerpo. / Se expresa corporalmente.")
        estandar_custom = st.text_area("Estándar de la Competencia (Opcional - Blanco para automático):", value="", height=70, placeholder="Texto del estándar...")
    with col_s2:
        tipo_motivacion = st.selectbox(
            "Tipo de Motivación para el Inicio de la Clase:",
            ["A través de una actividad física", "A través de una imagen", "A través de una historia"],
            index=0
        )
        criterios_custom = st.text_area("Criterios de Evaluación (Opcional - Blanco para automático):", value="", height=70, placeholder="Ej. 1. Ejecuta desplazamientos orientados en el patio. 2. Identifica nociones de derecha e izquierda.")
        evidencia_custom = st.text_input("Evidencia de Aprendizaje (Opcional - Blanco para automático):", value="", placeholder="Ej. Ejecución de desplazamientos coordinados hacia señales leídas.")

    fechas_duracion = fecha_sugerida
    duracion_semanas = 1
    sesiones_por_semana = 1
    producto_unidad = ""
    problema_contexto = titulo_sesion_input.strip() if titulo_sesion_input.strip() else "Desarrollo de nociones espaciales, coordinación motriz y convivencia en juegos de Educación Física."

else:  # Unidad o Proyecto EF
    f1, f2, f3, f4, f5 = st.columns(5)
    with f1:
        num_doc = st.text_input("N.° de Unidad / Proyecto:", "04")
    with f2:
        fechas_duracion = st.text_input("Fechas / Periodo:", "Del 22 de junio al 17 de julio de 2026")
        fecha_sugerida = fechas_duracion  # Garantiza que ambas variables existan siempre
    with f3:
        duracion_semanas = st.slider("Número de Semanas:", min_value=2, max_value=8, value=4)
    with f4:
        sesiones_por_semana = st.selectbox("Sesiones por Semana:", [1, 2, 3], index=1)
    with f5:
        producto_unidad = st.text_input("Producto Final Tangible:", "Festival Lúdico-Motor Peruanito")
        duracion_sesion = "90 minutos"

    problema_contexto = st.text_area(
        "📋 Describe el Tema, Problema de Contexto o Necesidad Motriz/Saludable de los Estudiantes:",
        height=120,
        value="Dificultades de coordinación motriz, orientación espacial en el patio al desplazarse en grupo, poco conocimiento de juegos tradicionales peruanos y falta de hábitos de higiene personal (lavado de manos, cambio de polo) al finalizar la actividad física."
    )

# ==============================================================================
# PROMPTS ESPECIALIZADOS QUE LEEN DE cneb_datos.py
# ==============================================================================

def generar_prompt_unidad_ef_10_secciones():
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "")
        des_list = comp_info["desempenos"].get(grado_normalizado_cneb, [])
        cneb_datos_text += f"\n\nCOMPETENCIA: {comp_nombre}\nESTÁNDAR OFICIAL ({ciclo_actual}):\n{est_txt}\nDESEMPEÑOS OFICIALES ({grado_normalizado_cneb}):\n" + "\n".join(des_list)

    total_sesiones_unidad = duracion_semanas * sesiones_por_semana

    return f"""
Actúa como un especialista en currículo educativo peruano y docente experto en el área de Educación Física para Educación Básica Regular (CNEB). 

Tu tarea es elaborar una UNIDAD DE APRENDIZAJE completa, extensa, rigurosa y alineada al Currículo Nacional (CNEB), siguiendo estrictamente las 10 secciones obligatorias sin cortar el documento al final.

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y SÍNTESIS DE CELDAS (OBLIGATORIO LLEGAR HASTA LA SECCIÓN X):
1. DEBES FINALIZAR EL DOCUMENTO OBLIGATORIAMENTE HASTA LA SECCIÓN X (RECURSOS Y FIRMAS DE DIRECTORA Y DOCENTE). QUEDA STRICTAMENTE PROHIBIDO DEJAR EL DOCUMENTO INCOMPLETO.
2. SÍNTESIS EN TABLAS: MANTÉN EL TEXTO DENTRO DE LAS CELDAS DE LAS TABLAS DE FORMA SINTÉTICA Y CONCISA (1 A 2 LÍNEAS POR CELDA) PARA NUNCA EXCEDER EL LÍMITE DE MEMORIA Y LOGRAR DESARROLLAR LAS 10 SECCIONES ENTERAS.
3. EN LA SECCIÓN VIII (MATRIZ DE PLANIFICACIÓN), DESARROLLA CADA UNA DE LAS {total_sesiones_unidad} SESIONES ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana). ESTÁ PROHIBIDO PONER PUNTOS SUSPENSIVOS (...) O SALTARSE SESIONES.
4. EN LA MATRIZ DE PLANIFICACIÓN: TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN LA PARTE SUPERIOR DE CADA SESIÓN CON NEGRITA EN LA PARTE EVALUADA, Y EL DESEMPEÑO COMPLETO EN LA COLUMNA CORRESPONDIENTE CON NEGRITA EN LO UTILIZADO Y PRECISADO.
5. COMPLETA SIEMPRE LA SECCIÓN IX (SECUENCIA DE SESIONES) Y LA SECCIÓN X (RECURSOS Y ESPACIO PARA FIRMAS).

DATOS OFICIALES EXTRAÍDOS DE cneb_datos.py PARA ESTA UNIDAD ({grado_seccion} - {ciclo_actual}):
{cneb_datos_text}

DATOS PARA LA GENERACIÓN:
- N° de Unidad: Unidad N° {num_doc}
- Ciclo / Grados: {ciclo_actual} - {grado_seccion}
- Nombre de la IE: {ie_nombre}
- Nombre del Docente: {docente}
- Nombre del Director(a): {director}
- Duración / Fechas: {duracion_semanas} semanas ({total_sesiones_unidad} sesiones en total, {sesiones_por_semana} por semana) - ({fechas_duracion})
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
- Describir un desempeño práctico o un producto tangible/demostrable claro: {producto_unidad}.

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

8. VIII. MATRIZ DE PLANIFICACIÓN (Formato Tabla detallado por las {total_sesiones_unidad} sesiones)
Desarrolla {total_sesiones_unidad} bloques de tablas independientes (uno por cada sesión):
- En la parte superior de cada bloque de sesión, incluye la fila con el ESTÁNDAR COMPLETO del CNEB correspondiente a la competencia evaluada, redactado de manera íntegra (sin modificar ni alterar su texto original), RESALTANDO EN NEGRITA la parte específica que se trabaja/evalúa en esa actividad.
- Columnas de la Matriz por cada sesión (Celdas sintéticas de 1 a 2 líneas):
  | Sesión N.° y Título de la sesión | Competencia / Capacidad | Desempeño | Criterios de Evaluación | Evidencia y Producto | Instrumento de Evaluación |
- REGLA DEL DESEMPEÑO: Redactado de manera COMPLETA tal cual aparece en el CNEB, RESALTANDO EN NEGRITA tanto la parte del desempeño utilizada como las palabras/términos agregados para su precisión y contextualización.
*NOTA: NO incluir la columna "Propósito" en la Matriz de Planificación.*

9. IX. SECUENCIA DE SESIONES (Formato Tabla)
Genera una tabla completa para las {total_sesiones_unidad} sesiones detallando (Propósito directo de 2 a 3 líneas):
| N° | Título de la actividad | Propósito de la actividad | Representación gráfica |

10. X. RECURSOS
- Recursos para el Docente (Normativa CNEB, RM N° 501-2025, materiales).
- Recursos para el Estudiante (Kit de aseo: jabón, toalla, polo de cambio, ropa deportiva, botellas de agua).
- Fecha y espacio para firmas de la Directora y Docente de Educación Física.
"""

def generar_prompt_proyecto_ef():
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "")
        des_list = comp_info["desempenos"].get(grado_normalizado_cneb, [])
        cneb_datos_text += f"\n\nCOMPETENCIA: {comp_nombre}\nESTÁNDAR OFICIAL ({ciclo_actual}):\n{est_txt}\nDESEMPEÑOS OFICIALES ({grado_normalizado_cneb}):\n" + "\n".join(des_list)

    total_sesiones_proyecto = duracion_semanas * sesiones_por_semana

    return f"""
Actúa como un Especialista Pedagógico experto en Educación Física del Ministerio de Educación de Perú (MINEDU). Tu tarea es diseñar un Proyecto de Aprendizaje completo bajo el enfoque por competencias del Currículo Nacional de la Educación Básica (CNEB), manteniendo de manera estricta y detallada una estructura formal sin cortar el documento al final.

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y SÍNTESIS EN TABLAS (OBLIGATORIO LLEGAR HASTA LA SECCIÓN IX):
1. DEBES DESARROLLAR EL PROYECTO COMPLETO LLEGANDO OBLIGATORIAMENTE HASTA LA SECCIÓN IX (RECURSOS Y MATERIALES Y FIRMAS DE DIRECTORA Y DOCENTE). QUEDA STRICTAMENTE PROHIBIDO CORTAR EL DOCUMENTO.
2. MANTÉN LAS RESPUESTAS Y TEXTOS DENTRO DE LAS CELDAS DE LAS TABLAS DE FORMA SINTÉTICA Y CONCISA (1 A 2 LÍNEAS POR CELDA) PARA GARANTIZAR QUE EL DOCUMENTO SE GENERE COMPLETO.
3. EN LA SECCIÓN VII (CUADRO CRONOLÓGICO DE SESIONES), DESARROLLA LAS {total_sesiones_proyecto} SESIONES ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana) UNA POR UNA. ESTÁ PROHIBIDO USAR PUNTOS SUSPENSIVOS (...) O OMITIR SESIONES.
4. EN LA SECCIÓN VI (MATRIZ DE PROPÓSITOS), TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN LA PARTE SUPERIOR DE CADA TABLA CON NEGRITA EN LA PARTE TRABAJADA, Y EL DESEMPEÑO COMPLETO EN LA COLUMNA CORRESPONDIENTE CON NEGRITA EN LO UTILIZADO Y PRECISADO.

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
- Duración y Frecuencia: {duracion_semanas} semanas, {sesiones_por_semana} sesiones por semana = {total_sesiones_proyecto} sesiones en total ({fechas_duracion})
- Tema o Problemática Central: {problema_contexto}
- Producto Final: {producto_unidad}

---

ESTRUCTURA OBLIGATORIA DEL PROYECTO DE APRENDIZAJE DE EDUCACIÓN FÍSICA:

I. TÍTULO DEL PROYECTO
- Debe ser motivador, creativo, retador y entre comillas (Ejemplo: "¡CELEBRAMOS NUESTRA PERUANIDAD EN EL GRAN FESTIVAL LÚDICO-MOTOR!").

II. DATOS INFORMATIVOS
- DRE/UGEL, IE, Nivel, Ciclo, Grado y Sección, Área (Educación Física), Duración, N° de sesiones, Docente, Director(a).

III. SITUACIÓN SIGNIFICATIVA
Redacta una situación basada en un contexto real de la escuela en 4 bloques sintéticos:
- Contexto, Problema o necesidad ({problema_contexto}), Reto (2 a 3 preguntas) y Propósito.

IV. CUADRO DE ENFOQUES TRANSVERSALES
Elabora una tabla con 1 o 2 enfoques transversales más pertinentes (Enfoque, Valores, Actitudes observables).

V. CUADRO DE NEGOCIACIÓN Y PLANIFICACIÓN CON LOS ESTUDIANTES
Tabla sintética de 4 columnas (¿Qué queremos hacer?, ¿Cómo lo haremos?, ¿Qué necesitamos?, ¿Cómo nos daremos cuenta de que lo logramos?) con respuestas realistas de asamblea.

VI. CUADRO DE PROPÓSITOS DE APRENDIZAJE Y EVALUACIÓN MATRIZADA
Organiza una tabla por competencia (C1, C2, C3 de Educación Física):
- PARTE SUPERIOR DE CADA TABLA: Estándar COMPLETO del {ciclo_actual} del CNEB con **negrita** en la parte movilizada.
- Tabla con 7 columnas (Celdas sintéticas y precisas):
  | Actividad General por Semana | Sesiones Vinculadas | Estándar Completo (encima) | Desempeño Completo (con **negrita**) | 3 Criterios de Evaluación | Evidencia de Aprendizaje | Instrumento |

VII. PLANIFICACIÓN CRONOLÓGICA DETALLADA DE LAS SESIONES
Desglosa secuencialmente las {total_sesiones_proyecto} sesiones. Tabla obligatoria de 3 COLUMNAS:
| Denominación de la sesión | Propósito detallado de la sesión | Representación gráfica |
- Denominación: Número y título motivador entre comillas.
- Propósito detallado: Síntesis clara que incluya calentamiento, juego práctico e higiene personal.
- Representación gráfica: Descripción breve de la imagen o esquema del patio.

VIII. PRODUCTOS DEL PROYECTO
- Producto Intangible / Práctico (Festival, Mini olimpiadas, Gincana).
- Producto Tangible ({producto_unidad}).

IX. RECURSOS Y MATERIALES
- Material deportivo, material reciclado, materiales de higiene.
- Espacios educativos y espacio para firmas de la Directora y Docente de Educación Física.
"""

def generar_prompt_sesion_ef():
    comps_str = ", ".join(comps_seleccionadas) if comps_seleccionadas else "Seleccionar automáticamente según el tema del CNEB"
    cap_str = capacidades_custom.strip() if capacidades_custom.strip() else "Generar automáticamente según la(s) competencia(s) elegida(s)"
    est_str = estandar_custom.strip() if estandar_custom.strip() else "Transcribir el Estándar COMPLETO oficial del ciclo del CNEB con negrita en la parte movilizada"
    crit_str = criterios_custom.strip() if criterios_custom.strip() else "Formular automáticamente mínimo 3 criterios claros con la estructura Acción + Contenido + Condición"
    evid_str = evidencia_custom.strip() if evidencia_custom.strip() else "Generar automáticamente la evidencia motriz o demostración práctica adecuada"

    return f"""
Actúa como Docente Experto en Educación Física para Primaria bajo el enfoque oficial del CNEB del MINEDU Perú.
Elabora una SESIÓN DE CLASE PRÁCTICA DE EDUCACIÓN FÍSICA completa para {grado_seccion} ({ciclo_actual}).

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
  1. **Motivación ({tipo_motivacion}):** [Desarrollar la motivación según el tipo elegido: {tipo_motivacion}].
  2. **Saberes previos:** [Preguntas abiertas sobre el tema/movimientos].
  3. **Problematización / Conflicto cognitivo:** [Reto motriz o pregunta desafiante].
  4. **Propósito de la clase:** [Comunicar qué aprenderán hoy].
  5. **Criterios de evaluación:** [Explicar cómo serán evaluados].
  6. **Acuerdos de convivencia:** [2 a 3 normas de seguridad en el patio].

- **DESARROLLO (Aprox. 60 min):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE:
  1. **Activación Corporal (Calentamiento dinámico):** Movilidad articular, trote lúdico y estiramientos.
  2. **Secuencia de Actividades Motrices:** Progresión práctica (3 actividades con hidratación).
  3. **ACTIVIDAD OBLIGATORIA DE ALTO NIVEL COGNITIVO (Analizar, Evaluar y Crear):** Reto motriz/estratégico donde los alumnos deban analizar una situación de juego, evaluar variantes tácticas y crear su propia regla o estrategia en equipo.

- **CIERRE (Aprox. 10 min) - DEBES REDACTAR OBLIGATORIAMENTE Y EN SU TOTALIDAD LOS SIGUIENTES 3 PUNTOS (PROHIBIDO OMITIR EL CIERRE):**
  1. **Vuelta a la calma:** Ejercicios de respiración guiada y estiramientos suaves.
  2. **Metacognición motriz:** Redacta de 3 a 4 preguntas reflexivas explícitas (¿Qué aprendimos sobre nuestro cuerpo? ¿Cómo superamos retos?).
  3. **Rutina Obligatoria de Higiene Personal:** Describe en detalle la práctica autónoma de aseo personal, lavado de manos con jabón, secado con toalla y cambio de polo deportivo.

8. TABLA VI: LISTA DE COTEJO DE EDUCACIÓN FÍSICA (Genera una tabla limpia con los criterios de evaluación y 8 a 10 estudiantes ficticios representativos para optimizar espacio y garantizar la completitud del documento).
"""

# ==============================================================================
# EJECUCIÓN CON SISTEMA DUAL ROBUSTO ANTI-404 Y COMPLETITUD
# ==============================================================================
st.markdown("---")

if st.button(f"✨ Generar {tipo_documento}"):
    if not api_key:
        st.error("⚠️ Ingresa tu API Key de Google AI Studio en la barra lateral izquierda.")
    elif not problema_contexto:
        st.warning("⚠️ Completa el campo del Tema o Problemática de Educación Física.")
    else:
        try:
            client = genai.Client(api_key=api_key)
            
            if tipo_documento == "Unidad de Aprendizaje":
                prompt_maestro = generar_prompt_unidad_ef_10_secciones()
            elif tipo_documento == "Proyecto de Aprendizaje":
                prompt_maestro = generar_prompt_proyecto_ef()
            else:
                prompt_maestro = generar_prompt_sesion_ef()

            sys_inst = """Eres un Especialista Curricular del MINEDU Perú dedicado exclusivamente al área de Educación Física. Generas documentos pedagógicos completos en Markdown alineados estrictamente al CNEB.
REGLA ABSOLUTA DE COMPLETITUD: Queda STRICTAMENTE PROHIBIDO recortar, abreviar o dejar incompletos los documentos al final.
Para evitar que el documento se corte al final, debes ser SINTÉTICO, CONCISO Y DIRECTO dentro de las celdas de las tablas, garantizando que el documento se redacte ENTERO de principio a fin, concluyendo obligatoriamente en la última sección con el espacio para firmas correspondientes."""

            with st.spinner(f"⚽ Google Gemini está redactando tu {tipo_documento} para {grado_seccion}..."):
                config = types.GenerateContentConfig(
                    system_instruction=sys_inst,
                    temperature=0.10,          # Temperatura baja para máxima fidelidad y concisión técnica
                    max_output_tokens=8192
                )
                
                # LISTA DE MODELOS ESTABLES CON RESPALDO AUTOMÁTICO EN CASO DE 404
                modelos_a_probar = [
                    model_choice,
                    "gemini-2.0-flash",
                    "gemini-1.5-flash",
                    "gemini-1.5-pro"
                ]
                
                response = None
                ultimo_err = None
                
                for mod in modelos_a_probar:
                    try:
                        response = client.models.generate_content(
                            model=mod,
                            contents=prompt_maestro,
                            config=config
                        )
                        if response and response.text:
                            break
                    except Exception as err:
                        ultimo_err = err
                        continue
                
                if not response or not response.text:
                    raise ultimo_err
                
                st.session_state['resultado_md'] = response.text
                st.session_state['tipo_doc_generado'] = tipo_documento
                st.session_state['fname_clean'] = f"{tipo_documento.replace(' ', '_')}_EF_N{num_doc}_{grado_seccion.replace(' ', '_')}.docx"
                
                st.success(f"✅ ¡{tipo_documento} de Educación Física generado con éxito!")

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                st.warning("⏳ Límite de velocidad alcanzado. Por favor, espera 60 segundos y vuelve a intentarlo.")
            else:
                st.error(f"❌ Ocurrió un error con la API de Google AI Studio: {err_str}")

# ==============================================================================
# DESPLIEGUE DE RESULTADOS Y DESCARGA EN WORD
# ==============================================================================
if st.session_state['resultado_md'] is not None:
    st.markdown("---")
    
    tab_preview, tab_download = st.tabs(["📄 Vista Previa (Permanente)", "📥 Descargar en Word (.docx)"])
    
    with tab_preview:
        st.markdown(st.session_state['resultado_md'])
        
    with tab_download:
        es_horizontal_doc = st.session_state['tipo_doc_generado'] in ["Unidad de Aprendizaje", "Proyecto de Aprendizaje"]
        
        buffer_doc = markdown_to_docx(
            st.session_state['resultado_md'], 
            ie_nombre=ie_nombre,
            es_horizontal=es_horizontal_doc
        )
        
        st.download_button(
            label=f"💾 Descargar {st.session_state['tipo_doc_generado']} en Word (.docx)",
            data=buffer_doc,
            file_name=st.session_state['fname_clean'],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        st.info("💡 **Nota:** El documento Word incluye la insignia editable y las tablas en tonos pasteles.")

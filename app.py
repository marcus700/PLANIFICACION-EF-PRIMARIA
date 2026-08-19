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
try:
    from cneb_datos import CNEB_PRIMARIA, obtener_ciclo_primaria
except ImportError:
    st.error("❌ Error: No se encuentra el archivo 'cneb_datos.py'. Asegúrate de tenerlo en la misma carpeta.")
    st.stop()

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

# SIDEBAR CON MODELOS ESTABLES DE GOOGLE STUDIO
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

# OPCIONES DE MODELOS OFICIALES Y ESTABLES
model_choice = st.sidebar.selectbox(
    "Modelo de Gemini:", 
    ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-pro", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"] # Lista actualizada con modelos comunes
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
    if st.button("🏃 Sesión de Aprendizaje", key="btn_sesion", use_container_width=True):
        st.session_state['tipo_documento'] = "Sesión de Aprendizaje"
        st.rerun()

tipo_documento = st.session_state['tipo_documento']

COLOR_MAP = {
    "Unidad de Aprendizaje": "#7C3AED",
    "Proyecto de Aprendizaje": "#059669",
    "Sesión de Aprendizaje": "#2563EB"
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

def markdown_to_docx(md_text, es_horizontal=False):
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
        if rows == 0: return
        cols = max(len(r) for r in t_data)
        if cols == 0: return
        
        t = doc.add_table(rows=rows, cols=cols)
        t.style = 'Table Grid'
        
        # Fijar ancho de columnas para tablas grandes
        if es_horizontal and cols > 4:
            t.autofit = False
            widths = [1.8, 1.2, 2.5, 2.2, 1.3, 1.0]
            for i, width in enumerate(widths):
                if i < cols:
                    for cell in t.columns[i].cells:
                        cell.width = Inches(width)

        for r_idx, row_cells in enumerate(t_data):
            for c_idx, cell_value in enumerate(row_cells):
                if c_idx < cols:
                    cell = t.cell(r_idx, c_idx)
                    p_cell = cell.paragraphs[0]
                    p_cell.text = ""
                    add_formatted_text(p_cell, cell_value.strip())
                    
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
        line_str = re.sub(r'<br\s*/?>', '\n', line_str)
        
        if line_str.startswith('|') and line_str.endswith('|'):
            if not in_table:
                in_table = True
                table_data = []
            if re.match(r'^\|[\s\:\-\|]+\|$', line_str):
                continue
            cells = [c.strip() for c in line_str.split('|')[1:-1]]
            table_data.append(cells)
            continue
        
        if in_table:
            table_count += 1
            header_color = PASTEL_COLORS[(table_count - 1) % len(PASTEL_COLORS)]
            render_table(table_data, header_color)
            in_table = False
            table_data = []

        heading_match = re.match(r'^(#{1,6})\s*(.*)$', line_str)
        if heading_match:
            hashes = heading_match.group(1)
            title_text = heading_match.group(2).strip().replace('**', '')
            level = len(hashes)
            
            p = doc.add_paragraph()
            run = p.add_run(title_text)
            if level <= 2:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 58, 138)
            elif level <= 4:
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = RGBColor(30, 58, 138)
            else:
                 run.font.bold = True
            continue

        if line_str.startswith(('• ', '- ', '* ')):
            p = doc.add_paragraph(style='List Bullet')
            add_formatted_text(p, line_str[2:].strip())
        elif line_str:
            p = doc.add_paragraph()
            add_formatted_text(p, line_str)

    if in_table and table_data:
        table_count += 1
        header_color = PASTEL_COLORS[(table_count - 1) % len(PASTEL_COLORS)]
        render_table(table_data, header_color)
            
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ==============================================================================
# FORMULARIO DE DATOS
# ==============================================================================
st.subheader(f"📝 Configuración de Datos para: {tipo_documento}")

c1, c2, c3 = st.columns(3)
with c1:
    dre_ugel = st.text_input("DRE / UGEL:", "Ica / Ica")
    ie_nombre = st.text_input("Institución Educativa:", "N.° 22314 'Vicenta Aquije de Huamán'")
with c2:
    director = st.text_input("Directora:", "Prof. Luisa Ruth Aronés Herrera")
    docente = st.text_input("Docente de Educación Física:", "Mario A. García Torres")
with c3:
    if tipo_documento in ["Unidad de Aprendizaje", "Proyecto de Aprendizaje"]:
        ciclo_seleccionado = st.selectbox("Ciclo y Grados a cargo:", ["III Ciclo (1° y 2°)", "IV Ciclo (3° y 4°)", "V Ciclo (5° y 6°)"], index=1)
        ciclo_actual = ciclo_seleccionado.split(" (")[0]
        if "III" in ciclo_actual: grados_ciclo = ["1° de Primaria", "2° de Primaria"]
        elif "IV" in ciclo_actual: grados_ciclo = ["3° de Primaria", "4° de Primaria"]
        else: grados_ciclo = ["5° de Primaria", "6° de Primaria"]
        st.info(f"Ciclo: **{ciclo_actual}** (Grados: {', '.join(grados_ciclo)})")
        grado_seccion = ciclo_seleccionado
        grado_normalizado_cneb = grados_ciclo[0]
    else: # Para Sesión de Aprendizaje
        grado_seccion = st.selectbox("Grado y Sección:", ["1° Grado A", "2° Grado A", "3° Grado A", "4° Grado A", "5° Grado A", "6° Grado A"], index=1)
        grado_normalizado_cneb = f"{grado_seccion.split('°')[0]}° de Primaria"
        ciclo_actual = obtener_ciclo_primaria(grado_normalizado_cneb)
        st.info(f"Ciclo CNEB Detectado: **{ciclo_actual}**")
        grados_ciclo = [] # No aplica para sesión

# VARIABLES ESPECÍFICAS PARA CADA HERRAMIENTA
if tipo_documento == "Sesión de Aprendizaje":
    f1, f2, f3 = st.columns(3)
    with f1:
        num_doc = st.text_input("N.° de Sesión:", "01")
    with f2:
        fecha_sugerida = st.text_input("Fecha:", "22 de junio de 2026")
    with f3:
        duracion_sesion = st.selectbox("Duración de la Clase:", ["45 minutos", "90 minutos", "135 minutos"], index=1)
    
    st.markdown("##### 📌 Configuración Pedagógica de la Sesión (Opcional)")
    titulo_sesion_input = st.text_input("Título de la Actividad / Sesión de Clase (Opcional):", value="", placeholder="Ej. Leemos señales para desplazarnos y reconocer direcciones")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        comps_seleccionadas = st.multiselect(
            "Competencia(s) a Trabajar:",
            list(CNEB_PRIMARIA.keys()),
            default=[list(CNEB_PRIMARIA.keys())[0]]
        )
        capacidades_custom = st.text_input("Capacidades Específicas (Opcional):", placeholder="Ej. Comprende su cuerpo.")
        estandar_custom = st.text_area("Estándar (Opcional):", height=70, placeholder="Texto del estándar...")
    with col_s2:
        tipo_motivacion = st.selectbox(
            "Tipo de Motivación:",
            ["A través de una actividad física", "A través de una imagen", "A través de una historia"],
            index=0
        )
        criterios_custom = st.text_area("Criterios de Evaluación (Opcional):", height=70, placeholder="1. Criterio uno. 2. Criterio dos.")
        evidencia_custom = st.text_input("Evidencia de Aprendizaje (Opcional):", placeholder="Ej. Ejecución de desplazamientos coordinados.")

    fechas_duracion = fecha_sugerida
    duracion_semanas = 1
    sesiones_por_semana = 1
    producto_unidad = ""
    problema_contexto = titulo_sesion_input.strip() if titulo_sesion_input.strip() else "Desarrollo de nociones espaciales y coordinación motriz."

else:  # Unidad o Proyecto EF
    f1, f2, f3, f4, f5 = st.columns(5)
    with f1:
        num_doc = st.text_input("N.° de Unidad / Proyecto:", "04")
    with f2:
        fechas_duracion = st.text_input("Fechas / Periodo:", "Del 22 de junio al 17 de julio de 2026")
    with f3:
        duracion_semanas = st.slider("Número de Semanas:", min_value=2, max_value=8, value=4)
    with f4:
        sesiones_por_semana = st.selectbox("Sesiones por Semana:", [1, 2, 3], index=1)
    with f5:
        producto_unidad = st.text_input("Producto Final Tangible:", "Festival Lúdico-Motor")

    problema_contexto = st.text_area(
        "📋 Describe el Tema, Problema de Contexto o Necesidad Motriz/Saludable:",
        height=120,
        value="Dificultades de coordinación motriz, orientación espacial, poco conocimiento de juegos tradicionales y falta de hábitos de higiene personal."
    )
    fecha_sugerida = ""
    duracion_sesion = "90 minutos"

# ==============================================================================
# PROMPTS ESPECIALIZADOS
# ==============================================================================

def generar_prompt_unidad_ef():
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "No encontrado")
        des_ciclo_list = []
        for g_ciclo in grados_ciclo:
            des_lista_grado = comp_info["desempenos"].get(g_ciclo, [])
            des_ciclo_list.extend([f"[{g_ciclo}] {d}" for d in des_lista_grado])
        cneb_datos_text += f"\n\n**COMPETENCIA: {comp_nombre}**\n**ESTÁNDAR OFICIAL ({ciclo_actual}):**\n{est_txt}\n**DESEMPEÑOS OFICIALES ({' y '.join(grados_ciclo)}):**\n" + "\n".join(des_ciclo_list)

    total_sesiones = duracion_semanas * sesiones_por_semana

    return f"""
Actúa como un especialista en currículo peruano y docente de Educación Física (CNEB). Elabora una UNIDAD DE APRENDIZAJE completa y rigurosa para el {ciclo_actual}, atendiendo simultáneamente a los grados {', '.join(grados_ciclo)}.

🚨 REGLAS CRÍTICAS DE ESTRUCTURA MULTIGRADO Y COMPLETITUD:
1.  FINALIZA EL DOCUMENTO HASTA LA SECCIÓN X (RECURSOS Y FIRMAS). Está prohibido dejar el documento incompleto.
2.  En la SECCIÓN VIII (MATRIZ DE PLANIFICACIÓN), desarrolla una ÚNICA tabla cronológica para las {total_sesiones} sesiones.
3.  En cada sesión de la matriz, selecciona y transcribe el DESEMPEÑO oficial del grado ({grados_ciclo[0]} o {grados_ciclo[1]}) que sea más pertinente para la complejidad de la actividad, asegurando un balance entre ambos grados a lo largo de la unidad.
4.  En la misma matriz, asegúrate de que, en el conjunto de las {total_sesiones} sesiones, se trabajen las 3 COMPETENCIAS del área de forma equilibrada.
5.  Para cada sesión, formula EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN claros y medibles (Acción + Contenido + Condición) pero redactados como una oración fluida, sin las etiquetas.
6.  Sé sintético en las celdas para garantizar la generación completa del documento.

**DATOS OFICIALES EXTRAÍDOS DE cneb_datos.py PARA ESTA UNIDAD ({ciclo_actual}):**
{cneb_datos_text}

**DATOS PARA LA GENERACIÓN:**
-   **N° de Unidad:** {num_doc}
-   **Ciclo / Grados:** {ciclo_actual} ({', '.join(grados_ciclo)})
-   **IE:** {ie_nombre}
-   **Docente:** {docente}
-   **Director(a):** {director}
-   **Duración:** {duracion_semanas} semanas ({total_sesiones} sesiones) - ({fechas_duracion})
-   **Problemática:** {problema_contexto}
-   **Producto:** {producto_unidad}

---
**ESTRUCTURA OBLIGATORIA DE LA UNIDAD DE APRENDIZAJE:**

**I. TÍTULO DE LA UNIDAD**
(Debe ser motivador y relacionado con el desarrollo de competencias)

**II. DATOS INFORMATIVOS**
(Tabla con IE, Director, Docente, Ciclo, Grados, Duración)

**III. SITUACIÓN SIGNIFICATIVA (Enfoque Dual - Estrategia Híbrida)**
-   **Vinculación Implícita (Acción):** Describe la problemática {problema_contexto} como una necesidad de mejora motriz pura (coordinación, orientación, agilidad).
-   **Vinculación Explícita (Reflexión):** Explica cómo en momentos de diálogo se conectará la práctica motriz con la utilidad preventiva ante un contexto real (ej. desastres naturales, salud, etc.).
-   Incluye un dato observable del problema y 3 preguntas retadoras.
-   Propón la estrategia pedagógica (ej. circuitos lúdico-motores).

**IV. PRODUCTO DE LA UNIDAD**
(Descripción clara del producto tangible: {producto_unidad})

**V. ENFOQUES TRANSVERSALES**
(Tabla con 2 enfoques, sus valores y actitudes observables en Ed. Física)

**VI. COMPETENCIAS TRANSVERSALES**
(Tabla con "Gestiona su aprendizaje" y "Se desenvuelve en entornos virtuales", con capacidades y desempeños aplicados)

**VII. ESTÁNDARES, COMPETENCIAS Y CAPACIDADES DEL ÁREA**
(Transcribe las 3 competencias oficiales con sus capacidades y estándares completos del {ciclo_actual})

**VIII. MATRIZ DE PLANIFICACIÓN (Formato Único Cronológico y Multigrado)**
> **ESTÁNDAR CNEB COMPLETO ({ciclo_actual}):** [Inserta aquí el estándar completo de la competencia principal que vertebra la unidad]
| Sesión N.° y Título | Competencia / Capacidades | Desempeño CNEB Completo (con **negrita**) | EXACTAMENTE 3 Criterios de Evaluación | Evidencia y Producto | Instrumento |

**IX. SECUENCIA DE SESIONES**
(Tabla con N°, Título, Propósito y Representación gráfica para las {total_sesiones} sesiones)

**X. RECURSOS**
(Recursos para docente y estudiante, fecha y espacio para firmas)
"""

def generar_prompt_proyecto_ef():
    # Similar a la unidad, pero con la estructura del proyecto.
    # Por brevedad, se mantiene una estructura similar al original ya que no fue modificada.
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "")
        des_list = []
        for g_ciclo in grados_ciclo:
             des_list.extend(comp_info["desempenos"].get(g_ciclo, []))
        cneb_datos_text += f"\n\n**COMPETENCIA: {comp_nombre}**\n**ESTÁNDAR ({ciclo_actual}):**\n{est_txt}\n**DESEMPEÑOS ({' y '.join(grados_ciclo)}):**\n" + "\n".join(des_list)
    
    total_sesiones = duracion_semanas * sesiones_por_semana
    return f"""
Actúa como Especialista Pedagógico del MINEDU en Educación Física. Diseña un PROYECTO DE APRENDIZAJE completo para el {ciclo_actual} ({', '.join(grados_ciclo)}).

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y SÍNTESIS:
1.  DESARROLLA EL PROYECTO HASTA LA SECCIÓN IX (RECURSOS Y FIRMAS).
2.  Sé CONCISO en las celdas de las tablas para asegurar la generación completa.
3.  En la SECCIÓN VII, desarrolla las {total_sesiones} sesiones sin omitir ninguna.
4.  En la SECCIÓN VI (MATRIZ DE PROPÓSITOS), organiza por competencia y usa desempeños de ambos grados del ciclo de forma pertinente.

**DATOS OFICIALES PARA ESTE PROYECTO ({ciclo_actual}):**
{cneb_datos_text}

**DATOS PARA LA GENERACIÓN:**
-   **N° de Proyecto:** {num_doc}
-   **DRE / UGEL:** {dre_ugel}
-   **IE:** {ie_nombre}
-   **Ciclo / Grados:** {ciclo_actual} ({', '.join(grados_ciclo)})
-   **Duración:** {duracion_semanas} semanas ({total_sesiones} sesiones) - {fechas_duracion}
-   **Problemática:** {problema_contexto}
-   **Producto Final:** {producto_unidad}
---
**ESTRUCTURA OBLIGATORIA DEL PROYECTO DE APRENDIZAJE:**

**I. TÍTULO DEL PROYECTO**
(Creativo, retador y motivador)

**II. DATOS INFORMATIVOS**
(Tabla con DRE/UGEL, IE, Nivel, Ciclo, Grados, Área, Duración, Docente, Director)

**III. SITUACIÓN SIGNIFICATIVA**
(Contexto, Problema, Reto y Propósito)

**IV. CUADRO DE ENFOQUES TRANSVERSALES**
(Tabla con 1 o 2 enfoques, valores y actitudes)

**V. CUADRO DE NEGOCIACIÓN Y PLANIFICACIÓN CON LOS ESTUDIANTES**
(Tabla con ¿Qué haremos?, ¿Cómo?, ¿Qué necesitamos?, ¿Cómo sabremos que lo logramos?)

**VI. CUADRO DE PROPÓSITOS DE APRENDIZAJE Y EVALUACIÓN MATRIZADA**
(Organizado por competencia, con estándar, desempeños de ambos grados, 3 criterios, evidencia e instrumento)

**VII. PLANIFICACIÓN CRONOLÓGICA DETALLADA DE LAS SESIONES**
(Tabla con Denominación, Propósito detallado y Representación gráfica para las {total_sesiones} sesiones)

**VIII. PRODUCTOS DEL PROYECTO**
(Producto Intangible y Producto Tangible: {producto_unidad})

**IX. RECURSOS Y MATERIALES**
(Material deportivo, reciclado, higiene, espacios y firmas)
"""


def generar_prompt_sesion_ef():
    comps_str = ", ".join(comps_seleccionadas) if comps_seleccionadas else "Seleccionar automáticamente"
    cap_str = capacidades_custom.strip() if capacidades_custom.strip() else "Generar automáticamente"
    est_str = estandar_custom.strip() if estandar_custom.strip() else f"Transcribir el Estándar COMPLETO del {ciclo_actual}"
    crit_str = criterios_custom.strip() if criterios_custom.strip() else "Formular automáticamente 3 criterios"
    evid_str = evidencia_custom.strip() if evidencia_custom.strip() else "Generar automáticamente"

    return f"""
Actúa como Docente Experto en Educación Física (CNEB). Elabora una SESIÓN DE CLASE PRÁCTICA completa para {grado_seccion} ({ciclo_actual}).

**DATOS PARA LA SESIÓN:**
-   **N.° Sesión:** {num_doc}
-   **Título:** "{problema_contexto}"
-   **IE:** {ie_nombre} | **Docente:** {docente} | **Fecha:** {fecha_sugerida} | **Duración:** {duracion_sesion}
-   **Motivación:** {tipo_motivacion}
-   **Competencia(s):** {comps_str}
-   **Capacidades:** {cap_str}
-   **Estándar:** {est_str}
-   **Criterios:** {crit_str}
-   **Evidencia:** {evid_str}

---
**REGLAS DE FORMATO Y ESTRUCTURA OBLIGATORIA:**

**1. ENCABEZADO Y TÍTULO:**
# **SESIÓN DE APRENDIZAJE DE EDUCACIÓN FÍSICA N.º {num_doc}**
## **"{problema_contexto.upper()}"**

**2. TABLA I: DATOS INFORMATIVOS**
| DATOS INFORMATIVOS | DETALLE |
| Institución Educativa | {ie_nombre} |
| Docente | {docente} |
| Grado y Sección | {grado_seccion} ({ciclo_actual}) |
| Fecha | {fecha_sugerida} |
| Duración | {duracion_sesion} |

**3. TABLA II: PROPÓSITOS DE APRENDIZAJE Y EVIDENCIAS**
> **ESTÁNDAR CNEB COMPLETO ({ciclo_actual}):** [Texto íntegro del estándar con **negrita** en la parte aplicada]
| ÁREA | COMPETENCIA Y CAPACIDADES | DESEMPEÑO PRECISADO (con **negrita**) | EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN | PROPÓSITO DE LA CLASE | EVIDENCIA | INSTRUMENTO |

**4. TABLA III: ENFOQUE TRANSVERSAL**
| ENFOQUE TRANSVERSAL | VALOR(ES) | ACTITUDES OBSERVABLES |

**5. TABLA IV: COMPETENCIAS TRANSVERSALES**
| COMPETENCIA TRANSVERSAL | CAPACIDADES | DESEMPEÑOS PRECISADOS |

**6. TABLA V: PREPARACIÓN DE LA CLASE**
| ¿Qué necesitamos hacer antes de la sesión? | ¿Qué recursos o materiales se utilizarán? |

**7. MOMENTOS DE LA CLASE (ENFOQUE DUAL COMPARTIMENTADO):**

-   **INICIO (Aprox. 20 min) - Enfoque Explícito:**
    (Redactado en primera persona plural, conectando la actividad con el contexto real)
    1.  **Motivación ({tipo_motivacion}):** [Vincula EXPLÍCITAMENTE la actividad con la problemática contextual].
    2.  **Saberes previos:** [Preguntas sobre el movimiento y su utilidad].
    3.  **Problematización:** [Reto motriz explicando su importancia para la vida real].
    4.  **Propósito y Criterios:** [Comunicar qué aprenderán y cómo serán evaluados].
    5.  **Acuerdos de convivencia:** [Normas de seguridad en el patio].

-   **DESARROLLO (Aprox. 60 min) - Vínculo Implícito Puro (Técnico):**
    (Redactado en primera persona plural, centrado en la biomecánica y el rigor técnico del movimiento)
    1.  **Activación Corporal:** [Calentamiento técnico específico].
    2.  **Secuencia de Actividades Motrices:** [3 actividades con descripción biomecánica precisa (postura, ejecución), sin mencionar el contexto].
    3.  **Actividad de Alto Nivel Cognitivo:** [Reto donde los alumnos co-evalúan la técnica de un compañero].

-   **CIERRE (Aprox. 10 min) - Síntesis Dual y Metacognición:**
    (Redacta obligatoriamente los 3 puntos, uniendo el mundo técnico con el contextual)
    1.  **Vuelta a la calma:** [Ejercicios de respiración y relajación].
    2.  **Metacognición de Síntesis Dual:** [3 preguntas explícitas que conecten la técnica aprendida con su utilidad en la vida real. Ej: ¿Cómo te ayuda la postura correcta que practicamos hoy a reaccionar mejor ante un sismo?].
    3.  **Rutina de Higiene Personal:** [Descripción detallada de la práctica de aseo (lavado de manos, cambio de polo)].

**8. TABLA VI: LISTA DE COTEJO DE EDUCACIÓN FÍSICA**
(Tabla con los criterios de evaluación y 8-10 estudiantes ficticios)
"""

# ==============================================================================
# EJECUCIÓN DEL MODELO
# ==============================================================================
st.markdown("---")

if st.button(f"✨ Generar {tipo_documento}"):
    if not api_key:
        st.error("⚠️ Ingresa tu API Key de Google AI Studio en la barra lateral.")
    elif not problema_contexto:
        st.warning("⚠️ Completa el campo del Tema o Problemática.")
    else:
        try:
            genai.configure(api_key=api_key)
            
            if tipo_documento == "Unidad de Aprendizaje":
                prompt_maestro = generar_prompt_unidad_ef()
                es_horizontal_word = True
            elif tipo_documento == "Proyecto de Aprendizaje":
                prompt_maestro = generar_prompt_proyecto_ef()
                es_horizontal_word = True
            else: # Sesión de Aprendizaje
                prompt_maestro = generar_prompt_sesion_ef()
                es_horizontal_word = False

            sys_inst = """Eres un Especialista Curricular del MINEDU Perú, experto en Educación Física. Generas documentos pedagógicos completos en Markdown, alineados estrictamente al CNEB.
REGLA ABSOLUTA DE COMPLETITUD: Queda STRICTAMENTE PROHIBIDO recortar, abreviar o dejar incompletos los documentos. Debes ser sintético y directo en las celdas de las tablas para garantizar que redactas el documento ENTERO, de principio a fin, concluyendo siempre en la última sección con las firmas."""

            with st.spinner(f"⚽ Google Gemini está redactando tu {tipo_documento}..."):
                model = genai.GenerativeModel(
                    model_name=model_choice,
                    system_instruction=sys_inst,
                    generation_config={"temperature": 0.15, "max_output_tokens": 8192}
                )
                
                # Sistema de reintento simple
                response = None
                try:
                    response = model.generate_content(prompt_maestro)
                except Exception as e:
                    st.warning(f"El modelo '{model_choice}' falló. Reintentando con 'gemini-1.5-flash'...")
                    time.sleep(5) # Pausa antes de reintentar
                    model = genai.GenerativeModel('gemini-1.5-flash', system_instruction=sys_inst)
                    response = model.generate_content(prompt_maestro)

                if not response or not hasattr(response, 'text'):
                     raise ValueError("La respuesta de la API no contiene texto válido.")

                st.session_state['resultado_md'] = response.text
                st.session_state['tipo_doc_generado'] = tipo_documento
                st.session_state['fname_clean'] = f"{tipo_documento.replace(' ', '_')}_{ciclo_actual.replace(' ','')}_N{num_doc}.docx"
                st.session_state['es_horizontal_doc'] = es_horizontal_word

                st.success(f"✅ ¡{tipo_documento} generado con éxito!")

        except Exception as e:
            err_str = str(e)
            if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                st.error("⏳ Límite de velocidad alcanzado. Por favor, espera 60 segundos y vuelve a intentarlo.")
            elif "503" in err_str or "high demand" in err_str:
                 st.error("🔥 El modelo está experimentando alta demanda. Por favor, intenta de nuevo en unos momentos o selecciona otro modelo en la barra lateral.")
            else:
                st.error(f"❌ Ocurrió un error con la API de Google AI Studio: {err_str}")

# ==============================================================================
# DESPLIEGUE DE RESULTADOS Y DESCARGA EN WORD
# ==============================================================================
if st.session_state.get('resultado_md'):
    st.markdown("---")
    
    tab_preview, tab_download = st.tabs(["📄 Vista Previa", "📥 Descargar en Word (.docx)"])
    
    with tab_preview:
        st.markdown(st.session_state['resultado_md'])
        
    with tab_download:
        try:
            buffer_doc = markdown_to_docx(
                st.session_state['resultado_md'], 
                es_horizontal=st.session_state.get('es_horizontal_doc', False)
            )
            
            st.download_button(
                label=f"💾 Descargar {st.session_state['tipo_doc_generado']} en Word",
                data=buffer_doc,
                file_name=st.session_state['fname_clean'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            st.info("💡 **Nota:** El documento Word se generará con tablas en tonos pasteles y formato optimizado.")
        except Exception as e:
            st.error(f"Ocurrió un error al generar el archivo Word: {e}")

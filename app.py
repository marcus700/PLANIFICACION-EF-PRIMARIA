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
# BASE DE DATOS EXACTA CNEB EDUCACIÓN FÍSICA PRIMARIA (MINEDU PERÚ)
# ==============================================================================
CNEB_EF_PRIMARIA = {
    "Se desenvuelve de manera autónoma a través de su motricidad": {
        "estandares": {
            "III Ciclo": "Se desenvuelve de manera autónoma a través de su motricidad cuando comprende cómo usar su cuerpo en las diferentes acciones que realiza utilizando su lado dominante y realiza movimientos coordinados que le ayudan a sentirse seguro en la práctica de actividades físicas. Se orienta espacialmente en relación a sí mismo y a otros puntos de referencia. Se expresa corporalmente con sus pares de diferentes utilizando el ritmo, gestos y movimientos como recursos para comunicar.",
            "IV Ciclo": "Se desenvuelve de manera autónoma a través de su motricidad cuando comprende cómo usar su cuerpo explorando la alternancia de sus lados corporales de acuerdo a su utilidad y ajustando la posición del cuerpo en el espacio y en el tiempo en diferentes etapas de las acciones motrices, con una actitud positiva y una voluntad de experimentar situaciones diversas. Experimenta nuevas posibilidades expresivas de su cuerpo y las utiliza para relacionarse y comunicar ideas, emociones, sentimientos, pensamientos.",
            "V Ciclo": "Se desenvuelve de manera autónoma a través de su motricidad cuando acepta sus posibilidades y limitaciones según su desarrollo e imagen corporal. Realiza secuencias de movimientos coordinados aplicando la alternancia de sus lados corporales de acuerdo a su utilidad. Produce con sus pares secuencias de movimientos corporales, expresivos o rítmicos en relación a una intención."
        },
        "desempenos": {
            "1er Grado A": [
                "Es autónomo al explorar las posibilidades de su cuerpo en diferentes acciones para mejorar sus movimientos (saltar, correr, lanzar) al mantener y/o recuperar el equilibrio en el espacio y con los objetos.",
                "Se orienta a través de sus nociones espacio-temporales (arriba - abajo, dentro - fuera, cerca - lejos) en relación a sí mismo.",
                "Descubre nuevos movimientos y gestos para representar objetos, personajes y estados de ánimo y ritmos sencillos.",
                "Se expresa motrizmente para comunicar sus emociones y representa en el juego acciones cotidianas de su familia y comunidad."
            ],
            "2do Grado A": [
                "Explora de manera autónoma sus posibilidades de movimiento al realizar con seguridad y confianza habilidades motrices básicas realizando movimientos coordinados.",
                "Se orienta en el espacio y tiempo en relación a sí mismo y a otros puntos de referencia, reconociendo su lado derecho e izquierdo y sus posibilidades de equilibrio en acciones lúdicas.",
                "Resuelve situaciones motrices al utilizar su lenguaje corporal (gesto, contacto visual, actitud corporal) para comunicarse mejor.",
                "Utiliza su cuerpo y el movimiento para expresar ideas y emociones en la práctica de actividades lúdicas con diferentes tipos de ritmos."
            ],
            "3er Grado A": [
                "Reconoce la izquierda y derecha en relación a objetos y en sus pares para mejorar sus posibilidades de movimiento en acciones lúdicas.",
                "Se orienta en un espacio y tiempo determinado en relación a sí mismo, objetos y compañeros, coordinando sus movimientos y regulando su equilibrio.",
                "Resuelve situaciones motrices al utilizar su lenguaje corporal, verbal y sonoro para comunicar actitudes y estados de ánimo.",
                "Vivencia el ritmo y se apropia de secuencias rítmicas corporales en situaciones de juego."
            ],
            "4to Grado A": [
                "Regula la posición del cuerpo en situaciones de equilibrio con modificación del espacio teniendo como referencia la trayectoria de objetos y sus desplazamientos.",
                "Alterna sus lados corporales de acuerdo a su utilidad/necesidad y se orienta en el espacio y tiempo en actividades lúdicas y predeportivas.",
                "Utiliza su cuerpo (posturas, gestos y mímica) y diferentes movimientos para expresar formas, ideas, emociones y pensamientos.",
                "Utiliza su lenguaje corporal para expresar su forma particular de moverse creando secuencias sencillas de movimiento."
            ],
            "5to Grado A": [
                "Anticipa las acciones motrices a realizar en un espacio y tiempo para mejorar las posibilidades de respuesta en la acción aplicando alternancia corporal.",
                "Pone en práctica las habilidades motrices específicas (carrera, salto y lanzamientos) para dar respuesta a situaciones motrices lúdicas y predeportivas.",
                "Crea movimientos y desplazamientos rítmicos teniendo como base la música de su región.",
                "Valora en sí mismo y en sus pares nuevas formas de movimiento y gestos corporales aceptando la diversidad de expresión."
            ],
            "6to Grado A": [
                "Anticipa las acciones motrices a realizar en un espacio y tiempo para mejorar la respuesta aplicando alternancia de lados corporales según su preferencia.",
                "Afianza las habilidades motrices específicas (carrera, salto y lanzamientos) a través de la regulación de su cuerpo en contextos predeportivos.",
                "Aplica su lenguaje corporal para expresar su forma particular de moverse al asumir y adjudicar diferentes roles.",
                "Crea con sus pares una secuencia de movimientos corporales, expresivos o rítmicos de manera programada y estructurada."
            ]
        }
    },
    "Asume una vida saludable": {
        "estandares": {
            "III Ciclo": "Asume una vida saludable cuando diferencia los alimentos saludables de su dieta familiar, los momentos adecuados para ingerirlos y las posturas que lo ayudan al buen desempeño en la práctica de actividad física y de la vida cotidiana, reconociendo la importancia del autocuidado. Participa regularmente en la práctica de actividades lúdicas identificando su ritmo cardiaco, respiración y sudoración; utiliza prácticas de activación corporal y psicológica antes de la actividad lúdica.",
            "IV Ciclo": "Asume una vida saludable cuando diferencia los alimentos de su dieta familiar y de su región que son saludables de los que no lo son. Previene riesgos relacionados con la postura e higiene conociendo aquellas que favorecen y no favorecen su salud e identifica su fuerza, resistencia y velocidad en la práctica de actividades lúdicas. Adapta su esfuerzo en la práctica de actividad física de acuerdo a las características de la actividad y a sus posibilidades, aplicando conocimientos relacionados con el ritmo cardiaco, la respiración y la sudoración. Realiza prácticas de activación corporal y psicológica, e incorpora el autocuidado relacionado con los ritmos de actividad y descanso para mejorar el funcionamiento de su organismo.",
            "V Ciclo": "Asume una vida saludable cuando utiliza instrumentos que miden la aptitud física y estado nutricional e interpreta la información de los resultados obtenidos para mejorar su calidad de vida. Replantea sus hábitos higiénicos y alimenticios tomando en cuenta los cambios físicos propios de la edad, evita la realización de ejercicios y posturas contraindicadas para la salud en la práctica de actividad física. Incorpora prácticas saludables para su organismo consumiendo alimentos adecuados a las características personales y evitando el consumo de drogas. Propone ejercicios de activación y relajación antes, durante y después de la práctica y participa en actividad física de distinta intensidad regulando su esfuerzo."
        },
        "desempenos": {
            "1er Grado A": [
                "Reconoce los alimentos de su dieta familiar y las posturas que son beneficiosas para su salud en la vida cotidiana y lúdica.",
                "Identifica en sí mismo y en otros la diferencia entre inspiración y espiración, en reposo y movimiento, regulando su esfuerzo.",
                "Realiza con autonomía prácticas de cuidado personal al asearse, vestirse y adoptar posturas adecuadas en el juego.",
                "Busca satisfacer sus necesidades corporales cuando tiene sed y resuelve dificultades producidas por el cansancio e inactividad."
            ],
            "2do Grado A": [
                "Comprende la importancia de la activación corporal (calentamiento) y psicológica antes de la actividad lúdica identificando signos como ritmo cardiaco, respiración y sudoración.",
                "Reflexiona sobre los alimentos saludables de su dieta familiar/regional, la hidratación y las posturas adecuadas en la práctica física.",
                "Incorpora prácticas de cuidado personal al asearse, vestirse y adoptar posturas adecuadas sin afectar su desempeño.",
                "Reconoce la importancia del autocuidado regulando su esfuerzo en la práctica de actividades lúdicas."
            ],
            "3er Grado A": [
                "Explica la importancia de la activación corporal (calentamiento) y psicológica que le ayuda a estar predispuesto a la actividad.",
                "Diferencia los alimentos de su dieta familiar y regional que son saludables de los que no lo son para la actividad física.",
                "Aplica los conocimientos de los beneficios de la actividad física relacionados con el ritmo cardiaco, respiración y sudoración adaptando su esfuerzo.",
                "Incorpora el autocuidado relacionado con los ritmos de actividad-descanso para mejorar el funcionamiento de su organismo."
            ],
            "4to Grado A": [
                "Selecciona actividades para la activación corporal y psicológica e identifica variaciones en la frecuencia cardiaca y respiratoria según el nivel de esfuerzo.",
                "Selecciona e incorpora en su dieta alimentos nutritivos y energéticos de su región que contribuyen a la práctica física.",
                "Incorpora el autocuidado relacionado con los ritmos de actividad-descanso, hidratación y protección solar.",
                "Adopta posturas adecuadas para prevenir problemas musculares y óseos incorporando ritmos de descanso."
            ],
            "5to Grado A": [
                "Identifica las condiciones que favorecen la aptitud física (pruebas físicas e IMC) para mejorar su calidad de vida.",
                "Comprende los cambios físicos propios de la edad y su repercusión en la higiene en la práctica física y reflexión alimenticia.",
                "Identifica posturas y ejercicios contraindicados para la salud en la práctica de actividad física.",
                "Aplica los beneficios relacionados con la salud al realizar actividades de activación corporal, psicológica y de recuperación."
            ],
            "6to Grado A": [
                "Conoce diferentes métodos de evaluación para determinar la aptitud física y selecciona los que mejor se adecúen a sus posibilidades.",
                "Comprende la importancia de la actividad física incorporándola a su vida cotidiana e identifica cambios físicos e higiene.",
                "Evita la realización de posturas y ejercicios contraindicados o cualquier práctica que perjudique su salud.",
                "Previene hábitos perjudiciales para su organismo como comida chatarra, sedentarismo y desórdenes alimenticios."
            ]
        }
    },
    "Interactúa a través de sus habilidades sociomotrices": {
        "estandares": {
            "III Ciclo": "Interactúa a través de sus habilidades sociomotrices al aceptar al otro como compañero de juego y busca el consenso sobre la manera de jugar para lograr el bienestar común y muestra una actitud de respeto evitando juegos violentos y humillantes; expresa su posición ante un conflicto con intención de resolverlo y escucha la posición de sus compañeros en los diferentes tipos de juegos. Resuelve situaciones motrices a través de estrategias colectivas y participa en la construcción de reglas de juego adaptadas a la situación y al entorno, para lograr un objetivo común en la práctica de actividades lúdicas.",
            "IV Ciclo": "Interactúa a través de sus habilidades sociomotrices al tomar acuerdos sobre la manera de jugar y los posibles cambios o conflictos que se den y propone adaptaciones o modificaciones para favorecer la inclusión de compañeros en actividades lúdicas, aceptando al oponente como compañero de juego. Adapta la estrategia de juego anticipando las intenciones de sus compañeros y oponentes para cumplir con los objetivos planteados. Propone reglas y las modifica de acuerdo a las necesidades del contexto y los intereses del grupo en la práctica de actividades físicas.",
            "V Ciclo": "Interactúa a través de sus habilidades sociomotrices proactivamente con un sentido de cooperación teniendo en cuenta las adaptaciones o modificaciones propuestas por el grupo en diferentes actividades físicas. Hace uso de estrategias de cooperación y oposición seleccionando los diferentes elementos técnicos y tácticos que se pueden dar en la práctica de actividades lúdicas y predeportivas, para resolver la situación de juego que le dé un mejor resultado y que responda a las variaciones que se presentan en el entorno."
        },
        "desempenos": {
            "1er Grado A": [
                "Asume roles y funciones de manera individual y dentro de un grupo interactuando espontáneamente en actividades lúdicas.",
                "Participa en juegos cooperativos y de oposición en parejas y pequeños grupos, aceptando al oponente como compañero de juego.",
                "Propone soluciones a situaciones motrices poniéndose de acuerdo con sus pares y respetando las reglas de juego."
            ],
            "2do Grado A": [
                "Participa en juegos cooperativos y de oposición tomando consensos sobre la manera de jugar y respetando al oponente.",
                "Muestra una actitud de respeto en la práctica lúdica evitando juegos bruscos, amenazas o apodos e incluyendo a todos.",
                "Resuelve de manera compartida situaciones en juegos tradicionales/autóctonos y adecúa reglas para la inclusión."
            ],
            "3er Grado A": [
                "Propone cambios en las condiciones de juego para posibilitar la inclusión de sus pares, promoviendo el respeto y participación.",
                "Participa en juegos cooperativos y de oposición aceptando al oponente como compañero de juego y tomando consensos.",
                "Asocia el resultado favorable en el juego a la necesidad de generar estrategias colectivas conociendo el rol de cada integrante."
            ],
            "4to Grado A": [
                "Propone normas y reglas en las actividades lúdicas modificándolas según las necesidades para favorecer la inclusión.",
                "Propone juegos populares/tradicionales con adaptaciones consensuadas por el grupo respetando al oponente.",
                "Adapta la estrategia de juego cuando prevé las intenciones de sus compañeros y oponentes para cumplir el objetivo."
            ],
            "5to Grado A": [
                "Emplea la resolución reflexiva y el diálogo para solucionar conflictos surgidos durante la práctica de actividades lúdicas y predeportivas.",
                "Realiza actividades lúdicas interactuando con compañeros y oponentes con respeto a las diferencias personales y cambio de roles.",
                "Propone junto a sus pares soluciones estratégicas oportunas al practicar juegos tradicionales, autóctonos y predeportivos."
            ],
            "6to Grado A": [
                "Participa en actividades en la naturaleza, eventos predeportivos y juegos populares tomando decisiones en favor del grupo con sentido solidario.",
                "Modifica juegos y actividades para que se adecúen a las posibilidades del grupo y a la lógica del juego deportivo.",
                "Discrimina y pone en práctica estrategias en actividades lúdicas, predeportivas y deportivas adecuando normas y soluciones tácticas."
            ]
        }
    }
}

def obtener_ciclo_ef(grado: str) -> str:
    if "1er" in grado or "2do" in grado:
        return "III Ciclo"
    elif "3er" in grado or "4to" in grado:
        return "IV Ciclo"
    return "V Ciclo"

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
    ["gemini-3.6-flash", "gemini-3.5-flash", "gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
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
    ciclo_actual = obtener_ciclo_ef(grado_seccion)
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
# PROMPTS ESPECIALIZADOS EN EDUCACIÓN FÍSICA
# ==============================================================================

def generar_prompt_unidad_ef_10_secciones():
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_EF_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "")
        des_list = comp_info["desempenos"].get(grado_seccion, [])
        cneb_datos_text += f"\n\nCOMPETENCIA: {comp_nombre}\nESTÁNDAR OFICIAL ({ciclo_actual}):\n{est_txt}\nDESEMPEÑOS OFICIALES ({grado_seccion}):\n" + "\n".join(des_list)

    total_sesiones_unidad = duracion_semanas * sesiones_por_semana

    return f"""
Actúa como un especialista en currículo educativo peruano y docente experto en el área de Educación Física para Educación Básica Regular (CNEB). 

Tu tarea es elaborar una UNIDAD DE APRENDIZAJE completa, extensa, rigurosa y alineada al Currículo Nacional (CNEB), siguiendo estrictamente las 10 secciones obligatorias sin cortar ni interrumpir el documento al final.

🚨 REGLAS CRÍTICAS DE COMPLETITUD Y ESTRUCTURA (OBLIGATORIO LLEGAR HASTA LA SECCIÓN X):
1. DEBES FINALIZAR EL DOCUMENTO OBLIGATORIAMENTE HASTA LA SECCIÓN X (RECURSOS Y FIRMAS). QUEDA PROHIBIDO CORTAR O DEJAR INCOMPLETA LA UNIDAD AL FINAL.
2. EN LA SECCIÓN VIII (MATRIZ DE PLANIFICACIÓN), DESARROLLA CADA UNA DE LAS {total_sesiones_unidad} SESIONES DE FORMA CLARA Y CONCISA ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana). ESTÁ PROHIBIDO PONER PUNTOS SUSPENSIVOS (...), RESÚMENES O OMITIR SESIONES.
3. EN LA MATRIZ DE PLANIFICACIÓN: TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN LA PARTE SUPERIOR DE CADA SESIÓN CON NEGRITA EN LA PARTE EVALUADA, Y EL DESEMPEÑO COMPLETO EN LA COLUMNA CORRESPONDIENTE CON NEGRITA EN LO UTILIZADO Y PRECISADO.
4. COMPLETA SIEMPRE LA SECCIÓN IX (SECUENCIA DE SESIONES CON SUS PROPÓSITOS Y REPRESENTACIONES GRÁFICAS) Y LA SECCIÓN X (RECURSOS Y ESPACIO PARA FIRMAS DE DIRECTORA Y DOCENTE).

DATOS OFICIALES EXTRAÍDOS DEL CNEB DE EDUCACIÓN FÍSICA PARA UTILIZAR EN ESTA UNIDAD ({grado_seccion} - {ciclo_actual}):
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
- Columnas de la Matriz por cada sesión:
  | Sesión N.° y Título de la sesión | Competencia / Capacidad | Desempeño | Criterios de Evaluación | Evidencia y Producto | Instrumento de Evaluación |
- REGLA DEL DESEMPEÑO: Redactado de manera COMPLETA tal cual aparece en el CNEB, RESALTANDO EN NEGRITA tanto la parte del desempeño utilizada como las palabras/términos agregados para su precisión y contextualización.
*NOTA: NO incluir la columna "Propósito" en la Matriz de Planificación.*

9. IX. SECUENCIA DE SESIONES (Formato Tabla)
Genera una tabla completa para las {total_sesiones_unidad} sesiones detallando:
| N° | Título de la actividad | Propósito de la actividad | Representación gráfica |
- El propósito debe ser explícito e incluir la secuencia metodológica (calentamiento/activación, desarrollo motriz/juego, hábitos de higiene personal y reflexión).
- La representación gráfica describe brevemente el esquema visual o distribución de materiales en el patio.

10. X. RECURSOS
- Recursos para el Docente (Normativa CNEB, RM N° 501-2025, materiales).
- Recursos para el Estudiante (Kit de aseo: jabón, toalla, polo de cambio, ropa deportiva, botellas de agua).
"""

def generar_prompt_proyecto_ef():
    cneb_datos_text = ""
    for comp_nombre, comp_info in CNEB_EF_PRIMARIA.items():
        est_txt = comp_info["estandares"].get(ciclo_actual, "")
        des_list = comp_info["desempenos"].get(grado_seccion, [])
        cneb_datos_text += f"\n\nCOMPETENCIA: {comp_nombre}\nESTÁNDAR OFICIAL ({ciclo_actual}):\n{est_txt}\nDESEMPEÑOS OFICIALES ({grado_seccion}):\n" + "\n".join(des_list)

    total_sesiones_proyecto = duracion_semanas * sesiones_por_semana

    return f"""
Actúa como un Especialista Pedagógico experto en Educación Física del Ministerio de Educación de Perú (MINEDU). Tu tarea es diseñar un Proyecto de Aprendizaje completo bajo el enfoque por competencias del Currículo Nacional de la Educación Básica (CNEB), manteniendo de manera estricta y detallada una estructura formal.

DATOS OFICIALES EXTRAÍDOS DEL CNEB DE EDUCACIÓN FÍSICA PARA UTILIZAR EN ESTE PROYECTO ({grado_seccion} - {ciclo_actual}):
{cneb_datos_text}

Para este nuevo proyecto, los datos de entrada son:
- Nivel y Grado: Educación Primaria, {grado_seccion} - {ciclo_actual}
- DRE / UGEL: {dre_ugel}
- Institución Educativa: {ie_nombre}
- Director(a): {director}
- Docente de Ed. Física: {docente}
- Duración y Frecuencia: {duracion_semanas} semanas, {sesiones_por_semana} sesiones por semana = {total_sesiones_proyecto} sesiones en total ({fechas_duracion})
- Tema o Problemática Central: {problema_contexto}
- Producto Final: {producto_unidad}

Genera el proyecto de manera exhaustiva respetando fielmente las siguientes secciones:

I. DATOS GENERALES:
Muestra la tabla de Datos Informativos con: DRE / UGEL ({dre_ugel}), I.E. ({ie_nombre}), Director ({director}), Docente ({docente}), Grado y Sección ({grado_seccion} - {ciclo_actual}), Duración ({duracion_semanas} semanas, {total_sesiones_proyecto} sesiones en total, {fechas_duracion}).

II. TÍTULO DEL PROYECTO: Redacta un título motivador entre comillas que evidencie el producto y el propósito (Ejemplo: "¡{producto_unidad.upper()} PARA PROMOVER LA VIDA SALUDABLE!").

III. SITUACIÓN SIGNIFICATIVA: Redacta una situación basada en un contexto real de la escuela, describiendo la problemática ({problema_contexto}), las consecuencias y planteando obligatoriamente de 2 a 3 retos en forma de preguntas para los estudiantes.

IV. CUADRO DE ENFOQUES TRANSVERSALES: Una tabla con 3 columnas (Enfoque Transversal, Valores, Actitudes observables o Acciones concretas) adaptados a la problemática.

V. CUADRO DE NEGOCIACIÓN / PLANIFICACIÓN CON LOS ESTUDIANTES: Una tabla de 4 columnas (¿Qué queremos hacer?, ¿Cómo lo haremos?, ¿Qué necesitamos?, ¿Cómo nos daremos cuenta de que lo logramos?) simulando las respuestas participativas de los niños en la sesión 1.

VI. CUADRO DE PROPÓSITOS DE APRENDIZAJE Y EVALUACIÓN MATRIZADA:
Organiza la matriz dividida por cada Competencia del área que intervenga (Se desenvuelve de manera autónoma..., Asume una vida saludable, Interactúa a través de sus habilidades sociomotrices). Para cada competencia, debes estructurar una tabla con las siguientes 7 columnas exactas:
  1. Actividad General por Semana: Coloca el título del bloque general o eje de la semana.
  2. Sesiones Vinculadas: Indica el número y nombre de la sesión (ejemplo: Sesión 1: "Nombre").
  3. Estándar de Aprendizaje Completo (CNEB): Copia de forma LITERAL y completa el estándar del {ciclo_actual} según el CNEB, sin recortar nada. Resalta en NEGRITA únicamente el fragmento específico que se movilizará en esa sesión.
  4. Desempeño Completo del Grado (CNEB) con Precisión: Copia de forma LITERAL y completa el desempeño oficial de {grado_seccion} del CNEB. Resalta en NEGRITA lo que se está utilizando de la norma y añade al final (también en NEGRITA) la precisión o el contexto específico de la sesión que tú le estás agregando como docente.
  5. 3 Criterios de Evaluación por Sesión: Redacta exactamente tres criterios claros, medibles y específicos por cada sesión lineal.
  6. Evidencia de Aprendizaje: Define el producto parcial o actuación tangible que dejará el alumno en esa sesión.
  7. Instrumento de Evaluación: Indica la herramienta de calificación formativa (Lista de cotejo, rúbrica, escala de valoración, etc.).

VII. PLANIFICACIÓN CRONOLÓGICA DETALLADA DE LAS SESIONES:
Desglosa secuencialmente las {total_sesiones_proyecto} sesiones indicadas ({duracion_semanas} semanas, {sesiones_por_semana} sesión(es) por semana). Presenta OBLIGATORIAMENTE una tabla con exactamente las siguientes 3 COLUMNAS:
| Denominación de la sesión | Propósito detallado de la sesión | Representación gráfica |
- Denominación de la sesión: Número y título motivador de la sesión entre comillas (ejemplo: Sesión 1: "Descubrimos trayectorias en el patio").
- Propósito detallado de la sesión: Explicación pedagógica explícita que incluya la secuencia metodológica (activación corporal/calentamiento, desarrollo motriz/juego práctico y rutina de higiene personal).
- Representación gráfica: Descripción breve de la imagen, esquema visual o distribución de materiales en el patio que representa la sesión.

VIII. PRODUCTOS DEL PROYECTO
- Producto Intangible / Práctico: (Ej. Festival deportivo, Mini olimpiadas, Gincana, Circuito motriz demostrativo).
- Producto Tangible: {producto_unidad} (Ej. Cartelera de compromisos de salud, mapa del circuito de juegos, etc.).

IX. RECURSOS Y MATERIALES
Detallar exhaustivamente:
- Material deportivo del patio (conos, aros, pelotas, cuerdas, silbato).
- Material reciclado / alternativo.
- Material de señalización y kit de aseo (jabón, toalla, polo de repuesto).
- Espacios educativos (patio, losa deportiva, campo).
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
*(Nota: Si algún dato anterior dice "generar automáticamente", completa los campos de forma rigurosa utilizando la base de datos oficial del CNEB de Educación Física).*

---

REGLAS DE FORMATO Y ESTRUCTURA OBLIGATORIA DE LA SESIÓN:

1. ENCABEZADO Y TÍTULO DE LA SESIÓN:
Muestra EXACTAMENTE la siguiente estructura en la parte superior:
# **SESIÓN DE APRENDIZAJE DE EDUCACIÓN FÍSICA N.º {num_doc}**
## **"{problema_contexto.upper()}"**
*(QUEDA STRICTAMENTE PROHIBIDO COLOCAR CUALQUIER OTRO DATO, FECHA O SUBTÍTULO DEBAJO DEL TÍTULO DE LA SESIÓN).*

2. TABLA I: DATOS INFORMATIVOS
| DATOS INFORMATIVOS | DETALLE |
| Institución Educativa | {ie_nombre} |
| Docente de Educación Física | {docente} |
| Grado y Sección | {grado_seccion} ({ciclo_actual}) |
| Fecha | {fecha_sugerida} |
| Duración | {duracion_sesion} |

3. TABLA II: PROPÓSITOS DE APRENDIZAJE Y EVIDENCIAS
REGLA DEL ESTÁNDAR: Coloca en la PARTE SUPERIOR / ENCABEZADO DE ESTA TABLA (como bloque inmediatamente superior a la matriz) el ESTÁNDAR COMPLETO CNEB del {ciclo_actual} para la competencia evaluada, redactado de manera íntegra (sin modificar ni recortar su texto original), RESALTANDO EN NEGRITA únicamente la parte específica que se trabaja/evalúa en esta sesión.

Estructura del bloque y tabla II:
> **ESTÁNDAR CNEB COMPLETO ({ciclo_actual}):** [Texto íntegro del estándar del ciclo con **negrita** en la parte aplicada]

| ÁREA | COMPETENCIA Y CAPACIDADES | DESEMPEÑO PRECISADO COMPLETO (con **negrita**) | CRITERIOS DE EVALUACIÓN | PROPÓSITO DE LA CLASE | EVIDENCIA | INSTRUMENTO |
- **Competencias:** Incluye la(s) competencia(s) solicitada(s) de Educación Física.
- **Desempeño:** Transcribe el Desempeño COMPLETO del CNEB para {grado_seccion}, resaltando en **negrita** la parte utilizada y los términos precisados agregados.

4. TABLA III: ENFOQUE TRANSVERSAL (ÚNICO Y ESPECÍFICO)
Coloca UN SOLO Enfoque Transversal (el más coherente e ideal para la actividad específica):
| ENFOQUE TRANSVERSAL PRIORIZADO | VALOR(ES) | ACTITUDES OBSERVABLES |

5. TABLA IV: COMPETENCIAS TRANSVERSALES
Coloca la tabla con las Competencias Transversales que se emplean en la sesión ("Gestiona su aprendizaje de manera autónoma" y/o "Se desenvuelve en entornos virtuales"):
| COMPETENCIA TRANSVERSAL | CAPACIDADES | DESEMPEÑOS PRECISADOS |

6. TABLA V: PREPARACIÓN DE LA CLASE
| ¿Qué necesitamos hacer antes de la sesión de Ed. Física? | ¿Qué recursos o materiales del patio se utilizarán? |

7. MOMENTOS DE LA CLASE DE EDUCACIÓN FÍSICA:

- **INICIO (Aprox. 20 min):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE. Debe considerar ESTRICTAMENTE el siguiente orden:
  1. **Motivación ({tipo_motivacion}):** [Desarrollar la motivación STRICTAMENTE según el tipo elegido: {tipo_motivacion}. Si es "A través de una imagen", describe la imagen motivadora y las preguntas reflexivas; si es "A través de una actividad física", describe el juego motivador inicial; si es "A través de una historia", relata la historia corta o cuento regional motriz].
  2. **Saberes previos:** [Preguntas abiertas sobre el tema o movimientos]
  3. **Problematización / Conflicto cognitivo:** [Reto motriz o pregunta desafiante sobre el juego/cuerpo]
  4. **Propósito de la clase:** [Comunicar con claridad qué aprenderán hoy]
  5. **Criterios de evaluación:** [Explicar de forma sencilla cómo serán evaluados]
  6. **Acuerdos de convivencia:** [Establecer 2 a 3 normas de respeto y seguridad en el patio]

- **DESARROLLO (Aprox. 60 min):**
  Redactado en PRIMERA PERSONA DEL PLURAL Y TIEMPO PRESENTE. Incluye la secuencia pedagógica:
  1. **Activación Corporal (Calentamiento dinámico):** Movilidad articular, trote lúdico con ritmos/cambios de dirección y estiramientos dinámicos en el patio.
  2. **Secuencia de Actividades Motrices:** Progresión de lo simple a lo complejo (3 a 4 actividades prácticas de exploración, juegos cooperativos o deportivos con pausas de hidratación).
  3. **ACTIVIDAD OBLIGATORIA DE ALTO NIVEL COGNITIVO (Analizar, Evaluar y Crear):** Incluye un reto motriz/estratégico específico donde los estudiantes deban **analizar** una situación de juego, **evaluar** soluciones o variantes tácticas en equipo y **crear** su propia regla, secuencia o estrategia motriz colectiva.

- **CIERRE (Aprox. 10 min) - DEBES REDACTAR OBLIGATORIAMENTE Y EN SU TOTALIDAD LOS SIGUIENTES 3 PUNTOS (PROHIBIDO OMITIR O CORTAR EL CIERRE):**
  1. **Vuelta a la calma:** Ejercicios de respiración guiada, relajación muscular y estiramientos suaves en el patio.
  2. **Metacognición motriz:** Redacta de 3 a 4 preguntas reflexivas pedagógicas explícitas (ej. ¿Qué aprendimos sobre nuestro cuerpo hoy? ¿Cómo superamos las dificultades en el juego? ¿Para qué nos sirve lo aprendido?).
  3. **Rutina Obligatoria de Higiene Personal:** Describe en detalle la práctica autónoma de aseo personal, lavado de manos con agua y jabón, secado con toalla y cambio de polo deportivo al concluir la clase.

8. TABLA VI: LISTA DE COTEJO DE EDUCACIÓN FÍSICA (Tabla con criterios de evaluación y 03 estudiantes ficticios).
"""

# ==============================================================================
# EJECUCIÓN CON SISTEMA DUAL ROBUSTO ANTI-404
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

            sys_inst = "Eres un Especialista Curricular del MINEDU Perú dedicado exclusivamente al área de Educación Física. Generas documentos completos en Markdown alineados strictly al CNEB."

            with st.spinner(f"⚽ Google Gemini está redactando tu {tipo_documento} para {grado_seccion}..."):
                config = types.GenerateContentConfig(
                    system_instruction=sys_inst,
                    temperature=0.15,
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

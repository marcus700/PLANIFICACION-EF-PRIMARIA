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

    /* ==========================================================================
       ESTILOS DE ALTO CONTRASTE Y LEGIBILIDAD PARA BOTONES DE HERRAMIENTAS
       ========================================================================== */
    
    /* 1. BOTÓN UNIDAD CNEB EF (PÚRPURA VIBRANTE) */
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

    /* 2. BOTÓN PROYECTO LÚDICO (VERDE ESMERALDA VIBRANTE) */
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

    /* 3. BOTÓN SESIÓN DE CLASE (AZUL ELÉCTRICO VIBRANTE) */
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

    /* 4. BOTÓN FICHA DE TRABAJO (NARANJA ÁMBAR VIBRANTE) */
    div.st-key-btn_ficha > button {
        background: linear-gradient(135deg, #D97706 0%, #B45309 100%) !important;
        background-color: #D97706 !important;
        border: 2px solid #B45309 !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 14px rgba(217, 119, 6, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    div.st-key-btn_ficha > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 18px rgba(217, 119, 6, 0.6) !important;
    }

    /* 5. BOTÓN PRINCIPAL DE GENERAR DOCUMENTO (✨ GENERAR) */
    div.stButton > button:not([key="btn_unidad"]):not([key="btn_proyecto"]):not([key="btn_sesion"]):not([key="btn_ficha"]) {
        background: linear-gradient(135deg, #1E40AF 0%, #1E3A8A 100%) !important;
        background-color: #1E40AF !important;
        border: 2px solid #1D4ED8 !important;
        border-radius: 12px !important;
        padding: 0.75rem 1.5rem !important;
        box-shadow: 0 4px 16px rgba(30, 64, 175, 0.5) !important;
        transition: all 0.3s ease !important;
    }
    div.stButton > button:not([key="btn_unidad"]):not([key="btn_proyecto"]):not([key="btn_sesion"]):not([key="btn_ficha"]):hover {
        background: linear-gradient(135deg, #2563EB 0%, #1E40AF 100%) !important;
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.7) !important;
    }

    /* FORZAR TEXTO BLANCO INTENSO Y NEGRILLA EN TODOS LOS BOTONES */
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
            target_pwd = st.secrets.get("APP_PASSWORD", "docente2026")
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
    st.session_state['tipo_documento'] = "Unidad de Aprendizaje (CNEB EF 10 Secciones)"

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

col_b1, col_b2, col_b3, col_b4 = st.columns(4)

with col_b1:
    if st.button("📘 Unidad CNEB EF (10 Secciones)", key="btn_unidad", use_container_width=True):
        st.session_state['tipo_documento'] = "Unidad de Aprendizaje (CNEB EF 10 Secciones)"
        st.rerun()

with col_b2:
    if st.button("🚀 Proyecto Lúdico / Deportivo EF", key="btn_proyecto", use_container_width=True):
        st.session_state['tipo_documento'] = "Proyecto Lúdico / Deportivo EF"
        st.rerun()

with col_b3:
    if st.button("🏃 Sesión de Clase de Ed. Física", key="btn_sesion", use_container_width=True):
        st.session_state['tipo_documento'] = "Sesión de Clase de Ed. Física"
        st.rerun()

with col_b4:
    if st.button("📝 Ficha de Autoevaluación EF", key="btn_ficha", use_container_width=True):
        st.session_state['tipo_documento'] = "Ficha de Trabajo / Autoevaluación EF"
        st.rerun()

tipo_documento = st.session_state['tipo_documento']

COLOR_MAP = {
    "Unidad de Aprendizaje (CNEB EF 10 Secciones)": "#7C3AED",
    "Proyecto Lúdico / Deportivo EF": "#059669",
    "Sesión de Clase de Ed. Física": "#2563EB",
    "Ficha de Trabajo / Autoevaluación EF": "#D97706"
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

if tipo_documento in ["Sesión de Clase de Ed. Física", "Ficha de Trabajo / Autoevaluación EF"]:
    f1, f2, f3 = st.columns(3)
    with f1:
        num_doc = st.text_input("N.° de Sesión / Ficha:", "01")
    with f2:
        fecha_sugerida = st.text_input("Fecha:", "22 de junio de 2026")
    with f3:
        duracion_sesion = st.selectbox("Duración de la Clase:", ["45 minutos", "90 minutos", "135 minutos"], index=1)
    
    fechas_duracion = fecha_sugerida
    duracion_semanas = 1
    producto_unidad = ""

else:  # Unidad o Proyecto EF
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        num_doc = st.text_input("N.° de Unidad / Proyecto:", "04")
    with f2:
        fechas_duracion = st.text_input("Fechas / Periodo:", "Del 22 de junio al 17 de julio de 2026")
    with f3:
        duracion_semanas = st.slider("Número de Semanas:", min_value=2, max_value=8, value=4)
    with f4:
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

    return f"""
Actúa como un especialista en currículo educativo peruano y docente experto en el área de Educación Física para Educación Básica Regular (CNEB). 

Tu tarea es elaborar una UNIDAD DE APRENDIZAJE completa, extensa, rigurosa y alineada al Currículo Nacional (CNEB), siguiendo estrictamente las 10 secciones obligatorias y las reglas anti-resumen.

🚨 REGLAS CRÍTICAS ANTI-RESUMEN (CUMPLIMIENTO OBLIGATORIO):
1. NO RESUMAS, NO ABREVIES Y NO OMITAS NINGUNA SECCIÓN NI NINGUNA SESIÓN.
2. EN LA SECCIÓN VIII (MATRIZ DE PLANIFICACIÓN) DEBES DESARROLLAR OBLIGATORIAMENTE LAS {duracion_semanas} SESIONES COMPLETAS. ESTÁ PROHIBIDO PONER PUNTOS SUSPENSIVOS (...), RESÚMENES O FRASES COMO "se repite para las siguientes sesiones".
3. TRANSCRIBE EL ESTÁNDAR COMPLETO DEL CNEB EN CADA BLOQUE DE ACTIVIDAD DE LA MATRIZ SIN RECORTAR TEXTO, RESALTANDO EN NEGRITA LA PARTE EVALUADA EN ESA SESIÓN.
4. TRANSCRIBE EL DESEMPEÑO COMPLETO DEL CNEB EN LA COLUMNA DE DESEMPEÑO SIN RECORTAR TEXTO, RESALTANDO EN NEGRITA LA PARTE UTILIZADA Y LOS TÉRMINOS PRECISADOS AGREGADOS PARA LA ACTIVIDAD.

DATOS OFICIALES EXTRAÍDOS DEL CNEB DE EDUCACIÓN FÍSICA PARA UTILIZAR EN ESTA UNIDAD ({grado_seccion} - {ciclo_actual}):
{cneb_datos_text}

DATOS PARA LA GENERACIÓN:
- N° de Unidad: Unidad N° {num_doc}
- Ciclo / Grados: {ciclo_actual} - {grado_seccion}
- Nombre de la IE: {ie_nombre}
- Nombre del Docente: {docente}
- Nombre del Director(a): {director}
- Duración / Fechas: {duracion_semanas} semanas ({fechas_duracion})
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

8. VIII. MATRIZ DE PLANIFICACIÓN (Formato Tabla detallado por las {duracion_semanas} sesiones)
Desarrolla {duracion_semanas} bloques de tablas independientes (uno por cada sesión/semana):
- En la parte superior de cada bloque de sesión, incluye la fila con el ESTÁNDAR COMPLETO del CNEB correspondiente a la competencia evaluada, redactado de manera íntegra (sin modificar ni alterar su texto original), RESALTANDO EN NEGRITA la parte específica que se trabaja/evalúa en esa actividad.
- Columnas de la Matriz por cada sesión:
  | Sesión N.° y Título de la sesión | Competencia / Capacidad | Desempeño | Criterios de Evaluación | Evidencia y Producto | Instrumento de Evaluación |
- REGLA DEL DESEMPEÑO: Redactado de manera COMPLETA tal cual aparece en el CNEB, RESALTANDO EN NEGRITA tanto la parte del desempeño utilizada como las palabras/términos agregados para su precisión y contextualización.
*NOTA: NO incluir la columna "Propósito" en la Matriz de Planificación.*

9. IX. SECUENCIA DE SESIONES (Formato Tabla)
Genera una tabla completa para las {duracion_semanas} sesiones detallando:
| N° | Título de la actividad | Propósito de la actividad | Representación gráfica |
- El propósito debe ser detallado e incluir la secuencia metodológica (calentamiento/activación, desarrollo motriz/juego, hábitos de higiene personal y reflexión).
- La representación gráfica describe brevemente el esquema visual o distribución de materiales en el patio.

10. X. RECURSOS
- Recursos para el Docente (Normativa CNEB, RM N° 501-2025, materiales).
- Recursos para el Estudiante (Kit de aseo: jabón, toalla, polo de cambio, ropa deportiva, botellas de agua).
- Fecha y espacio para firmas (Directora y Docente de Educación Física).
"""

def generar_prompt_proyecto_ef():
    return f"""
Actúa como Especialista en Educación Física del MINEDU Perú. Elabora un PROYECTO LÚDICO / DEPORTIVO COMPLETO de {duracion_semanas} semanas.
Área exclusiva: Educación Física. Tema/Problema: {problema_contexto}.

ESTRUCTURA DE SALIDA (MARKDOWN LIMPIO EN TABLAS):
# **PROYECTO LÚDICO - DEPORTIVO DE EDUCACIÓN FÍSICA N.º {num_doc}**
## **"{producto_unidad.upper()}"**

1. DATOS INFORMATIVOS (Tabla con DRE/UGEL, IE: {ie_nombre}, Director: {director}, Docente Ed. Física: {docente}, Grado: {grado_seccion}, Fechas: {fechas_duracion}).
2. SITUACIÓN SIGNIFICATIVA (3 párrafos enfocados en la motricidad, salud, juego en equipo y preguntas retadoras).
3. PLANIFICACIÓN DEL PROYECTO CON LOS ESTUDIANTES (Tabla: ¿Qué haremos?, ¿Qué sabemos?, ¿Cómo nos organizaremos?, ¿Qué materiales del patio necesitamos?).
4. MATRIZ DE PROPÓSITOS DE APRENDIZAJE DE EDUCACIÓN FÍSICA POR SEMANA (Semana 1 a {duracion_semanas}):
   - Aborda las 3 competencias de Ed. Física (Se desenvuelve..., Asume una vida saludable, Interactúa...).
   - Incluye el Estándar COMPLETO del CNEB con **negrita** en la parte aplicada y el Desempeño COMPLETO con **negrita** en la precisión.
5. SECUENCIA DE SESIONES PRÁCTICAS SEMANA A SEMANA ({duracion_semanas} semanas):
   - Tabla con actividades de calentamiento, desarrollo de juegos/deportes e higiene personal posjuego.
6. PRODUCTO FINAL TANGIBLE/DEMOSTRABLE ({producto_unidad}).
7. LISTA DE MATERIALES DEL PATIO Y KIT DE ASEO.
8. REFLEXIONES Y METAGOGNICIÓN DOCENTE.
"""

def generar_prompt_sesion_ef():
    return f"""
Actúa como Docente Experto en Educación Física para Primaria (CNEB MINEDU).
Elabora una SESIÓN DE CLASE PRÁCTICA DE EDUCACIÓN FÍSICA completa para {grado_seccion}.
Tema: {problema_contexto}. Duración: {duracion_sesion}.

ESTRUCTURA DE SALIDA REQUERIDA:
# **SESIÓN DE APRENDIZAJE DE EDUCACIÓN FÍSICA N.º {num_doc}**
## **"{problema_contexto.upper()}"**

• TABLA I: DATOS INFORMATIVOS (IE: {ie_nombre}, Docente: {docente}, Grado: {grado_seccion}, Fecha: {fecha_sugerida}, Duración: {duracion_sesion}).
• TABLA II: PROPÓSITOS DE APRENDIZAJE Y EVIDENCIAS
| ÁREA | COMPETENCIA Y CAPACIDADES | ESTÁNDAR CNEB COMPLETO (con **negrita**) | DESEMPEÑO PRECISADO COMPLETO (con **negrita**) | CRITERIOS DE EVALUACIÓN | PROPÓSITO DE LA CLASE | EVIDENCIA | INSTRUMENTO |
• TABLA III: ENFOQUES TRANSVERSALES Y COMPETENCIA TRANSVERSAL.
• TABLA IV: PREPARACIÓN DE LA CLASE (Materiales del patio: conos, aros, pelotas, silbato, kit de aseo).

• MOMENTOS DE LA CLASE DE EDUCACIÓN FÍSICA:
- **INICIO (20 min):** Activación corporal / Calentamiento dinámico con música/juegos, movilidad articular, saberes previos, delimitación del espacio y acuerdos de seguridad.
- **DESARROLLO (60 min):** Secuencia motriz de lo simple a lo complejo. 3 a 4 actividades prácticas descritas en PRIMERA PERSONA DEL PLURAL TIEMPO PRESENTE (exploración motriz, juegos cooperativos/deportivos, pausas de hidratación).
- **CIERRE (10 min):** Vuelta a la calma (respiración guiada, estiramientos), metacognición motriz y **Rutina Obligatoria de Higiene Personal (lavado de manos, secado con toalla y cambio de polo)**.

• TABLA V: LISTA DE COTEJO DE EDUCACIÓN FÍSICA (con 30 alumnos ficticios).
"""

def generar_prompt_ficha_ef():
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
            
            if tipo_documento == "Unidad de Aprendizaje (CNEB EF 10 Secciones)":
                prompt_maestro = generar_prompt_unidad_ef_10_secciones()
            elif tipo_documento == "Proyecto Lúdico / Deportivo EF":
                prompt_maestro = generar_prompt_proyecto_ef()
            elif tipo_documento == "Sesión de Clase de Ed. Física":
                prompt_maestro = generar_prompt_sesion_ef()
            else:
                prompt_maestro = generar_prompt_ficha_ef()

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
        es_horizontal_doc = st.session_state['tipo_doc_generado'] in ["Unidad de Aprendizaje (CNEB EF 10 Secciones)", "Proyecto Lúdico / Deportivo EF"]
        
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

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
import datetime

# Importar la Base de Datos CNEB directamente del archivo cneb_datos.py
try:
    from cneb_datos import CNEB_PRIMARIA, obtener_ciclo_primaria
except Exception as e:
    st.error(f"Error al cargar cneb_datos.py: {e}. Asegúrate de que el archivo se llame 'cneb_datos.py' sin punto al final.")

# Función auxiliar para conectar los nombres del menú con las claves de cneb_datos.py
def mapear_grado_cneb(grado_str):
    mapa = {
        "1° Grado": "1° de Primaria",
        "2° Grado": "2° de Primaria",
        "3° Grado": "3° de Primaria",
        "4° Grado": "4° de Primaria",
        "5° Grado": "5° de Primaria",
        "6° Grado": "6° de Primaria"
    }
    return mapa.get(grado_str, grado_str)

# Función para calcular el PIN del mes actual de forma automática
def obtener_pin_mes_actual():
    pines_mensuales = {
        1:  "EF26-ENE#9482",
        2:  "CNEB-FEB$7391",
        3:  "MOTOR-MAR*5820",
        4:  "ACTIV-ABR#3164",
        5:  "FIT26-MAY$8295",
        6:  "LUDO-JUN*6417",
        7:  "MINEDU-JUL#9531",
        8:  "ATHLET-AGO$4826",
        9:  "SPORT-SET*1935",
        10: "MOTRIZ-OCT#7264",
        11: "EF26-NOV$8419",
        12: "SALUD-DIC*3058"
    }
    mes_actual = datetime.datetime.now().month
    return pines_mensuales.get(mes_actual, "MINEDU-JUL#9531")

# Configuración visual de la plataforma
st.set_page_config(page_title="PlanificaEF", page_icon="🏃‍♂️", layout="centered")

# ==============================================================================
# ESTILOS CSS PERSONALIZADOS: PLATAFORMA COLORIDA Y TABLAS PASTEL
# ==============================================================================
st.markdown("""
    <style>
    /* Ocultar barra superior y menú Streamlit / GitHub */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stAppHeader {display: none;}
    .stDeployButton {display: none;}
    div[data-testid="stToolbar"] {visibility: hidden; height: 0%; position: fixed;}
    div[data-testid="stDecoration"] {visibility: hidden; height: 0%; position: fixed;}
    
    /* Fondo y fuentes globales */
    .stApp {
        background-color: #F4F9F4;
    }
    
    /* Botones vibrantes con bordes redondeados */
    .stButton>button {
        background: linear-gradient(135deg, #4CAF50 0%, #2E7D32 100%) !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 10px !important;
        border: none !important;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.15) !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0px 6px 15px rgba(0, 0, 0, 0.2) !important;
    }
    
    /* Pestañas coloridas */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        white-space: pre-wrap;
        background-color: #E8F5E9;
        border-radius: 8px 8px 0px 0px;
        color: #1B5E20;
        font-weight: bold;
        border: 1px solid #A5D6A7;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2E7D32 !important;
        color: white !important;
    }
    
    /* Estilos para Tablas en pantalla con tonos Pastel y bordes verdes */
    table {
        border-collapse: collapse !important;
        width: 100% !important;
        border: 2px solid #81C784 !important;
        border-radius: 8px !important;
        overflow: hidden !important;
        margin: 15px 0 !important;
    }
    th {
        background-color: #C8E6C9 !important;
        color: #1B5E20 !important;
        border: 1px solid #A5D6A7 !important;
        font-weight: bold !important;
        padding: 10px !important;
        text-align: center !important;
    }
    td {
        border: 1px solid #C8E6C9 !important;
        padding: 8px 12px !important;
        font-size: 0.9em !important;
    }
    tr:nth-child(even) {
        background-color: #F1F8E9 !important;
    }
    tr:nth-child(odd) {
        background-color: #FFFFFF !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# CONFIGURACIÓN DE PAGO, SEGURIDAD Y PINES AUTOMÁTICOS
# ==============================================================================
NUMERO_WHATSAPP = "51900000000"  # 👈 REEMPLAZA CON TU NÚMERO DE WHATSAPP CON 51 (ej. 51987654321)
NUMERO_YAPE_PLIN = "900 000 000" # 👈 REEMPLAZA CON TU NÚMERO DE YAPE / PLIN

# PIN seguro que cambia solo automáticamente cada mes + tu PIN Maestro permanente
PIN_DEL_MES = obtener_pin_mes_actual()
PIN_MAESTRO_ADMIN = "MAESTRO-ADMIN"
PINES_ACTIVOS = [PIN_DEL_MES, PIN_MAESTRO_ADMIN]

# LÍMITES GRATUITOS
MAX_UNIDADES_GRATIS = 1
MAX_SESIONES_GRATIS = 1

# Registro global de DNIs/Teléfonos que ya gastaron su prueba gratis en el servidor
if "dnis_bloqueados_servidor" not in st.session_state:
    st.session_state["dnis_bloqueados_servidor"] = set()

# Estado de la sesión del usuario actual
if "identificado" not in st.session_state:
    st.session_state["identificado"] = False
if "dni_usuario" not in st.session_state:
    st.session_state["dni_usuario"] = ""
if "unidades_generadas" not in st.session_state:
    st.session_state["unidades_generadas"] = 0
if "sesiones_generadas" not in st.session_state:
    st.session_state["sesiones_generadas"] = 0
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

# ==============================================================================
# SISTEMA DE SEGURIDAD Y PANTALLA DE BLOQUEO POR PAGO
# ==============================================================================
def mostrar_bloqueo_pago(motivo=""):
    st.markdown("---")
    st.error("🔒 **Acceso Restringido - Límite Gratuito Alcanzado**")
    st.warning("Estimado docente, el DNI/Celular " + str(st.session_state['dni_usuario']) + " o su dispositivo ya consumió la prueba gratuita (1 Unidad y 1 Sesión). Para seguir generando documentos sin límites, suscríbete por **S/ 15.00 soles al mes**.")
    
    col_pago, col_login = st.columns(2)
    
    with col_pago:
        st.markdown("### 📲 ¿Cómo suscribirte?")
        st.write("1. Realiza el Yape o Plin de **S/ 15.00** al número: **" + str(NUMERO_YAPE_PLIN) + "**")
        st.write("2. Envía la captura del pago por WhatsApp.")
        st.write("3. Te enviaremos tu **PIN de Acceso Mensual** al instante.")
        
        mensaje_wa = "Hola, soy el docente con DNI " + str(st.session_state['dni_usuario']) + ". Alcancé mi prueba gratuita en PlanificaEF y deseo suscribirme por S/ 15 soles al mes. Adjunto mi pago."
        link_wa = "https://wa.me/" + str(NUMERO_WHATSAPP) + "?text=" + re.sub(r' ', '%20', mensaje_wa)
        
        st.markdown(f'''
            <a href="{link_wa}" target="_blank">
                <button style="background-color:#25D366; color:white; border:none; padding:12px 20px; border-radius:8px; font-weight:bold; cursor:pointer; width:100%;">
                    📲 Solicitar PIN por WhatsApp
                </button>
            </a>
        ''', unsafe_allow_html=True)

    with col_login:
        st.markdown("### 🔑 Desbloquear con PIN")
        pin_ingresado = st.text_input("Ingresa tu PIN mensual de suscriptor:", type="password", key=f"input_pin_{motivo}")
        
        if st.button("🔓 Activar Acceso Ilimitado", key=f"btn_pin_{motivo}", use_container_width=True):
            if pin_ingresado in PINES_ACTIVOS:
                st.session_state["autenticado"] = True
                st.success("¡PIN Correcto! Acceso ilimitado activado.")
                st.rerun()
            else:
                st.error("❌ PIN incorrecto o vencido. Solicita tu PIN mensual por WhatsApp.")

# PANTALLA DE IDENTIFICACIÓN OBLIGATORIA
if not st.session_state["autenticado"] and not st.session_state["identificado"]:
    st.title("🏃‍♂️ PlanificaEF - Registro de Prueba")
    st.subheader("Asistente Pedagógico de Educación Física (Primaria - CNEB)")
    st.info("💡 **Prueba Gratuita:** Ingresa tu DNI o N° de Celular para activar 1 Unidad y 1 Sesión de regalo.")
    
    with st.form("form_identificacion"):
        dni_input = st.text_input("Ingresa tu DNI o N° de Celular:", placeholder="Ej. 71234567 o 987654321", key="input_ident")
        btn_ident = st.form_submit_button("🚀 Iniciar Prueba Gratuita")
        
    if btn_ident and dni_input:
        dni_limpio = dni_input.strip()
        st.session_state["dni_usuario"] = dni_limpio
        
        if dni_limpio in st.session_state["dnis_bloqueados_servidor"]:
            st.session_state["unidades_generadas"] = MAX_UNIDADES_GRATIS
            st.session_state["sesiones_generadas"] = MAX_SESIONES_GRATIS
            
        st.session_state["identificado"] = True
        st.rerun()
        
    st.markdown("---")
    st.markdown("### 🔑 ¿Ya eres suscriptor pagado?")
    st.write("Si ya tienes tu PIN mensual de S/ 15 soles, ingrésalo directamente aquí:")
    pin_directo = st.text_input("PIN de Suscriptor:", type="password", key="pin_directo_front")
    if st.button("🔓 Entrar con PIN"):
        if pin_directo in PINES_ACTIVOS:
            st.session_state["autenticado"] = True
            st.session_state["identificado"] = True
            st.rerun()
        else:
            st.error("❌ PIN Incorrecto.")
            
    st.stop()

# ENCABEZADO DE ESTADO
st.title("🏃‍♂️ PlanificaEF")

if st.session_state["autenticado"]:
    st.success("✅ **Suscripción Activa:** Generaciones ilimitadas activadas para este mes.")
else:
    u_usadas = st.session_state["unidades_generadas"]
    s_usadas = st.session_state["sesiones_generadas"]
    st.info("👤 Docente: **" + str(st.session_state['dni_usuario']) + "** | 💡 **Modo Prueba:** Unidades creadas: **" + str(u_usadas) + "/" + str(MAX_UNIDADES_GRATIS) + "** | Sesiones creadas: **" + str(s_usadas) + "/" + str(MAX_SESIONES_GRATIS) + "**")

# Enlace automático a la clave secreta guardada de forma segura
api_key = st.secrets.get("GEMINI_API_KEY", None)

if not api_key:
    st.error("⚠️ No se encontró la GEMINI_API_KEY en los secretos de Streamlit.")

# ==============================================================================
# FUNCIONES AUXILIARES: LIMPIEZA Y CONVERTIDOR PROFESIONAL DE MARKDOWN A WORD
# ==============================================================================
def limpiar_texto(texto):
    if not texto:
        return ""
    texto_limpio = texto.replace('||', '|\n|')
    return texto_limpio

def set_cell_background(cell, fill_color):
    """Aplica color de fondo a una celda de tabla en Word."""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def set_cell_borders(cell, color="81C784"):
    """Aplica bordes de color verde menta a la celda en Word."""
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        element = tcBorders.find(qn('w:{}'.format(edge)))
        if element is None:
            element = OxmlElement('w:{}'.format(edge))
            tcBorders.append(element)
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '6')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), color)

def crear_archivo_word_profesional(texto_markdown):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    lineas = texto_markdown.split('\n')
    i = 0
    
    # Paleta de Colores Pastel para encabezados de columnas
    colores_pastel = ["D4EDDA", "D1ECF1", "FFF3CD", "E8DAEF", "F8D7DA", "D5F5E3", "FDEBD0"]

    while i < len(lineas):
        linea = lineas[i].strip()
        if not linea:
            i += 1
            continue

        if linea.startswith('# '):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(linea.replace('# ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(15)
            run.font.bold = True
            run.font.color.rgb = RGBColor(27, 94, 32)
            i += 1
        elif linea.startswith('## '):
            p = doc.add_paragraph()
            run = p.add_run(linea.replace('## ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            run.font.bold = True
            run.font.color.rgb = RGBColor(46, 125, 50)
            i += 1
        elif linea.startswith('### '):
            p = doc.add_paragraph()
            run = p.add_run(linea.replace('### ', ''))
            run.font.name = 'Arial'
            run.font.size = Pt(10.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(27, 94, 32)
            i += 1
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
                            valor_celda_limpio = cell_value.replace('<br>', '\n').replace('<br/>', '\n').replace('<BR>', '\n')
                            partes = re.split(r'(\*\*.*?\*\*)', valor_celda_limpio)
                            for parte in partes:
                                if parte.startswith('**') and parte.endswith('**'):
                                    run = p.add_run(parte[2:-2])
                                    run.bold = True
                                else:
                                    run = p.add_run(parte)
                                run.font.name = 'Arial'
                                if r_idx == 0:
                                    run.font.bold = True
                                    run.font.color.rgb = RGBColor(27, 94, 32) # Texto verde oscuro sobre fondo pastel
                                    run.font.size = Pt(9.5)
                                else:
                                    run.font.size = Pt(8.5)
                            
                            # Aplicar bordes verde menta
                            set_cell_borders(cell, color="81C784")
                            
                            # Aplicar tono pastel a encabezados de columna
                            if r_idx == 0:
                                color_p = colores_pastel[c_idx % len(colores_pastel)]
                                set_cell_background(cell, color_p)
                            elif r_idx % 2 == 1:
                                set_cell_background(cell, "F1F8E9") # Cebra pastel muy suave
                d

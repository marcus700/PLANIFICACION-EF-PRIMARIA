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

# FUNCIONES AUXILIARES
def limpiar_texto(texto):
    if not texto:
        return ""
    texto_limpio = texto.replace('||', '|\n|')
    return texto_limpio

def set_cell_background(cell, fill_color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

def crear_archivo_word_profesional(texto_markdown):
    doc = Document()
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
                                    run.font.color.rgb = RGBColor(255, 255, 255)
                                    run.font.size = Pt(9.5)
                                else:
                                    run.font.size = Pt(8.5)
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
        if not st.session_state["autenticado"] and st.session_state["unidades_generadas"] >= MAX_UNIDADES_GRATIS:
            st.session_state["dnis_bloqueados_servidor"].add(st.session_state["dni_usuario"])
            mostrar_bloqueo_pago("Unidad de Aprendizaje")
        else:
            with st.spinner("Extrayendo matriz curricular oficial de cneb_datos.py..."):
                try:
                    client = genai.Client(api_key=api_key)
                    grado_cneb_u = mapear_grado_cneb(grado_u)
                    ciclo_u = obtener_ciclo_primaria(grado_cneb_u)
                    
                    matriz_estandares_u = ""
                    matriz_desempenos_u = ""
                    if CNEB_PRIMARIA:
                        est_u_list = []
                        des_u_list = []
                        for comp_name, comp_data in CNEB_PRIMARIA.items():
                            est_val = comp_data.get("estandares", {}).get(ciclo_u, "")
                            des_list = comp_data.get("desempenos", {}).get(grado_cneb_u, [])
                            if est_val:
                                est_u_list.append("• COMPETENCIA: " + str(comp_name) + "\nESTÁNDAR OFICIAL: \"" + str(est_val) + "\"")
                            if des_list:
                                des_u_list.append("• COMPETENCIA: " + str(comp_name) + "\nDESEMPEÑOS CON NUMERACIÓN OFICIAL CNEB PARA " + str(grado_u) + ":\n" + "\n".join(des_list))
                        matriz_estandares_u = "\n\n".join(est_u_list)
                        matriz_desempenos_u = "\n\n".join(des_u_list)

                    instrucciones_u = (
                        "Actúa como un Especialista Curricular experto en Educación Física para Primaria bajo el enfoque del CNEB de Perú (MINEDU).\n"
                        "Diseña una UNIDAD DE APRENDIZAJE completa estructurada EN TABLAS MARKDOWN CON LÍNEAS SEPARADORAS (|---|---|).\n\n"
                        "MATRIZ OFICIAL EXACTA Y PALABRA POR PALABRA EXTRAÍDA DE CNEB_DATOS.PY PARA " + str(grado_u) + " (" + str(ciclo_u) + "):\n\n"
                        "ESTÁNDARES OFICIALES DEL CICLO (" + str(ciclo_u) + "):\n" + str(matriz_estandares_u) + "\n\n"
                        "DESEMPEÑOS CON NUMERACIÓN OFICIAL DEL GRADO (" + str(grado_u) + "):\n" + str(matriz_desempenos_u) + "\n\n"
                        "ESTRUCTURA DE LA UNIDAD DE APRENDIZAJE:\n\n"
                        "1. DATOS INFORMATIVOS:\n"
                        "| Campo | Detalle |\n"
                        "| :--- | :--- |\n"
                        "| **DRE / UGEL** | DRE [.....] / UGEL [.....] |\n"
                        "| **Institución Educativa** | I.E. N° [.....] |\n"
                        "| **Lugar / Localidad** | [.....] |\n"
                        "| **Ciclo** | " + str(ciclo_u) + " |\n"
                        "| **Grado y Sección** | " + str(grado_u) + ", Secciones: [.....] |\n"
                        "| **Docente del Área** | [.....] |\n"
                        "| **Director(a)** | [.....] |\n"
                        "| **Duración y Periodo** | " + str(duracion_u) + " (Del [Día/Mes] al [Día/Mes]) |\n\n"
                        "2. TÍTULO DE LA UNIDAD DE APRENDIZAJE (Significativo, retador e innovador).\n\n"
                        "3. SITUACIÓN SIGNIFICATIVA (Contexto del problema, Reto en pregunta y Producto de la unidad).\n\n"
                        "4. PROPÓSITOS DE APRENDIZAJE Y SECUENCIA DE SESIONES:\n"
                        "| ACTIVIDAD (SESIÓN) | DESCRIPCIÓN PEDAGÓGICA | COMPETENCIA / CAPACIDADES | ESTÁNDAR DE LA COMPETENCIA | DESEMPEÑO PRECISADO | CRITERIOS DE EVALUACIÓN | INSTRUMENTO DE EVALUACIÓN |\n"
                        "| :--- | :--- | :--- | :--- | :--- | :--- | :--- |\n\n"
                        "REGLAS ABSOLUTAS DE TRANSCRIPCIÓN DE CNEB_DATOS.PY:\n"
                        "- En 'ESTÁNDAR DE LA COMPETENCIA': Copia PALABRA POR PALABRA el Estándar Oficial proporcionado arriba. PROHIBIDO refrasear. Únicamente **resalta en negrita** (`**texto**`) la frase que se ejercita.\n"
                        "- En 'DESEMPEÑO PRECISADO': Copia PALABRA POR PALABRA el Desempeño Oficial del CNEB conservando su NUMERACIÓN ORIGINAL (ejemplo: 1.1.-, 1.2.-). **Resalta en negrita** (`**texto en negrita**`) la frase del CNEB tomada y lo agregado al final para precisarlo.\n"
                        "- En 'CRITERIOS DE EVALUACIÓN': Formula OBLIGATORIAMENTE EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN por cada sesión separados con <br>.\n"
                        "- En 'INSTRUMENTO DE EVALUACIÓN': Lista de cotejo / Rúbrica.\n\n"
                        "5. ENFOQUES TRANSVERSALES PRIORIZADOS (Tabla con línea separadora).\n"
                        "6. MATERIALES Y RECURSOS DIDÁCTICOS."
                    )

                    pedido_u = "Crea una unidad para " + str(grado_u) + " con duración de " + str(duracion_u) + ". Contexto del problema: " + str(problema_u)
                    
                    resultado_u = generar_respuesta_ia(client, instrucciones_u, pedido_u)
                    resultado_u = limpiar_texto(resultado_u)
                    
                    if not st.session_state["autenticado"]:
                        st.session_state["unidades_generadas"] += 1
                        if st.session_state["unidades_generadas"] >= MAX_UNIDADES_GRATIS and st.session_state["sesiones_generadas"] >= MAX_SESIONES_GRATIS:
                            st.session_state["dnis_bloqueados_servidor"].add(st.session_state["dni_usuario"])
                        
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
    st.write("Genera el desarrollo de una sesión diaria paso a paso extrayendo datos reales del CNEB.")
    with st.form("form_sesion"):
        col1, col2 = st.columns(2)
        with col1:
            grado_s = st.selectbox("Grado de Primaria:", ["1° Grado", "2° Grado", "3° Grado", "4° Grado", "5° Grado", "6° Grado"], index=2, key="s1")
            ie_s = st.text_input("I.E. N°:", value="22314 Vicenta Aquije de Huamán", key="s_ie")
            docente_s = st.text_input("Docente:", value="Mario Garcia Torres", key="s_doc")
        with col2:
            competencia_s = st.selectbox("Competencia Principal:", ["Se desenvuelve de manera autónoma a través de su motricidad", "Asume una vida saludable", "Interactúa a través de sus habilidades sociomotrices"], key="s2")
            fecha_s = st.text_input("Fecha:", value="29/07/2026", key="s_fec")
            tiempo_s = st.selectbox("Tiempo:", ["45 minutos", "90 minutos", "135 minutos"], index=1, key="s_time")

        detalles_s = st.text_area("Tema de la clase o materiales disponibles:", placeholder="Ej. Coordinación óculo-manual lanzando y recibiendo pelotas de plástico.", key="s3")
        boton_sesion = st.form_submit_button("⚡ Generar Sesión en Word")

    if boton_sesion and detalles_s:
        if not st.session_state["autenticado"] and st.session_state["sesiones_generadas"] >= MAX_SESIONES_GRATIS:
            st.session_state["dnis_bloqueados_servidor"].add(st.session_state["dni_usuario"])
            mostrar_bloqueo_pago("Sesión de Aprendizaje")
        else:
            with st.spinner("Extrayendo matriz oficial de cneb_datos.py y diseñando la sesión..."):
                try:
                    client = genai.Client(api_key=api_key)
                    grado_cneb_s = mapear_grado_cneb(grado_s)
                    ciclo_s = obtener_ciclo_primaria(grado_cneb_s)
                    
                    estandar_base = ""
                    desempenos_base = ""
                    if CNEB_PRIMARIA and competencia_s in CNEB_PRIMARIA:
                        estandar_base = CNEB_PRIMARIA[competencia_s]["estandares"].get(ciclo_s, "")
                        desempenos_lista = CNEB_PRIMARIA[competencia_s]["desempenos"].get(grado_cneb_s, [])
                        desempenos_base = "\n".join(desempenos_lista)

                    instrucciones = (
                        "Actúa como un docente experto de Educación Física de nivel Primaria en Perú, especialista en el CNEB de MINEDU.\n\n"
                        "DATOS OFICIALES DE CNEB_DATOS.PY:\n"
                        "- Grado y Ciclo: " + str(grado_s) + " (" + str(ciclo_s) + ")\n"
                        "- Competencia principal: " + str(competencia_s) + "\n"
                        "- Tema / Propósito motriz: " + str(detalles_s) + "\n"
                        "- ESTÁNDAR CNEB OFICIAL LITERAL: \"" + str(estandar_base) + "\"\n"
                        "- DESEMPEÑOS CNEB OFICIALES DISPONIBLES:\n" + str(desempenos_base) + "\n\n"
                        "ESTRUCTURA OBLIGATORIA A GENERAR EN MARKDOWN:\n\n"
                        "# SESIÓN DE APRENDIZAJE N°.......\n"
                        "**Título:** [Crea un título motivador e innovador sobre " + str(detalles_s) + "]\n\n"
                        "## 1. DATOS INFORMATIVOS\n"
                        "| Campo | Detalle |\n"
                        "| :--- | :--- |\n"
                        "| **I.E N°** | " + str(ie_s) + " |\n"
                        "| **Grado y Sección** | " + str(grado_s) + " |\n"
                        "| **Area** | Educación física |\n"
                        "| **Docente** | " + str(docente_s) + " |\n"
                        "| **Fecha** | " + str(fecha_s) + " |\n"
                        "| **tiempo** | " + str(tiempo_s) + " |\n\n"
                        "## 2. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE\n"
                        "| COMPETENCIA / CAPACIDADES | ESTÁNDAR CNEB | DESEMPEÑOS PRECISADO | CRITERIOS DE EVALUACIÓN | EVIDENCIA Y PRODUCTO | INSTRUMENTO DE EVALUACIÓN |\n"
                        "| :--- | :--- | :--- | :--- | :--- | :--- |\n"
                        "REGLAS DE TRANSCRIPCIÓN:\n"
                        "- Columna 1: Competencia (" + str(competencia_s) + ") y sus capacidades separadas con <br>.\n"
                        "- Columna 2: Transcribe PALABRA POR PALABRA Y COMPLETO el estándar oficial (\"" + str(estandar_base) + "\"). **Resalta en negrita** solo la parte trabajada hoy.\n"
                        "- Columna 3: Selecciona el desempeño de " + str(grado_s) + " de la lista provista arriba. Conserva su numeración exacta (ej. 1.1.-), transcribe palabra por palabra sin modificar, y **resalta en negrita** la frase tomada del CNEB original y lo que le agregas al final para precisarlo al tema (" + str(detalles_s) + ").\n"
                        "- Columna 4: Formula EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN separados con <br>. Implícitos (Acción+Contenido+Condición) SIN escribir las palabras '(Acción)' ni '(Contenido)' por escrito.\n"
                        "- Columna 5: Evidencia de aprendizaje.\n"
                        "- Columna 6: Lista de cotejo.\n\n"
                        "## 3. ENFOQUE TRANSVERSAL\n"
                        "| Enfoque Transversal Priorizado | Valor | Actitud / Comportamiento Observable |\n"
                        "| :--- | :--- | :--- |\n\n"
                        "## 4. PREPARACIÓN DE LA SESIÓN\n"
                        "| ¿Qué necesitamos hacer antes de la sesión? | Recursos o Materiales a utilizar |\n"
                        "| :--- | :--- |\n\n"
                        "## 5. SECUENCIA DIDÁCTICA (MOMENTOS DE LA SESIÓN)\n"
                        "Redacta en PRIMERA PERSONA ('Recibo...', 'Explico...', 'Organizo...') y TIEMPO PRESENTE.\n"
                        "### A) INICIO (20% - 18 min): Motivación, saberes previos, problematización, propósito, acuerdos, Activación Corporal (calentamiento y toma de pulso inicial).\n"
                        "### B) DESARROLLO (60% - 54 min): 3 a 4 actividades en progresión, hidratación, normas de seguridad y retroalimentación.\n"
                        "### C) CIERRE (20% - 18 min): Vuelta a la calma (estiramientos, respiración, pulso final), Hábitos de higiene (aseo y lavado de manos) y metacognición.\n\n"
                        "## 6. ANEXO: INSTRUMENTO DE EVALUACIÓN\n"
                        "Tabla de Lista de Cotejo con los 3 criterios formulados y filas para nombres de estudiantes."
                    )
                    
                    pedido = "Diseña una sesión para " + str(grado_s) + ". Competencia: " + str(competencia_s) + ". Detalles del tema: " + str(detalles_s)
                    
                    resultado_s = generar_respuesta_ia(client, instrucciones, pedido)
                    resultado_s = limpiar_texto(resultado_s)
                    
                    if not st.session_state["autenticado"]:
                        st.session_state["sesiones_generadas"] += 1
                        if st.session_state["unidades_generadas"] >= MAX_UNIDADES_GRATIS and st.session_state["sesiones_generadas"] >= MAX_SESIONES_GRATIS:
                            st.session_state["dnis_bloqueados_servidor"].add(st.session_state["dni_usuario"])
                        
                    st.success("¡Sesión generada con éxito!")
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
                instrucciones_r = (
                    "Actúa como un Evaluador Pedagógico experto en Educación Física para Primaria.\n"
                    "Diseña una rúbrica analítica estructurada con los niveles: En Inicio, En Proceso, Logrado y Logro Destacado para el desempeño solicitado, utilizando exactamente 3 criterios claros y observables alineados al CNEB sin etiquetar explícitamente '(Acción)' ni '(Contenido)'."
                )

                pedido_r = "Crea una rúbrica para " + str(grado_r) + ". Competencia: " + str(competencia_r) + ". Desempeño: " + str(criterio_r)
                
                resultado_r = generar_respuesta_ia(client, instrucciones_r, pedido_r)
                resultado_r = limpiar_texto(resultado_r)
                
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

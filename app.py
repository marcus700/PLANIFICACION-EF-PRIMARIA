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

# Configuración visual de la plataforma
st.set_page_config(page_title="PlanificaEF", page_icon="🏃‍♂️", layout="centered")

# ==============================================================================
# CONFIGURACIÓN DE PAGO Y PINES VÁLIDOS (CAMBIA TU NÚMERO Y PINES AQUÍ)
# ==============================================================================
NUMERO_WHATSAPP = "51937287225"  # 👈 REEMPLAZA CON TU NÚMERO DE WHATSAPP (ej. 51987654321)
NUMERO_YAPE_PLIN = "937 287 225" # 👈 REEMPLAZA CON TU NÚMERO DE YAPE / PLIN

# Lista de PINES mensuales activos que les entregarás a los docentes que te paguen
PINES_ACTIVOS = st.secrets.get("PINES_ACTIVOS", ["EF2026", "PLANIFICA15", "PROFE1", "MAESTRO2026"])

# Estado de autenticación
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

# ==============================================================================
# PANTALLA DE BLOQUEO Y PAGO (PAYWALL)
# ==============================================================================
if not st.session_state["autenticado"]:
    st.title("🔒 PlanificaEF - Acceso Suscriptores")
    st.subheader("Plataforma Inteligente de Planificación para Educación Física Primaria CNEB")
    
    st.warning("⚠️ Esta plataforma es de acceso exclusivo para docentes suscriptores.")
    
    st.markdown("---")
    
    col_pago, col_login = st.columns(2)
    
    with col_pago:
        st.markdown("### 📲 ¿Cómo suscribirte?")
        st.write("💰 **Costo:** **S/ 15.00 soles al mes** (Acceso ilimitado).")
        st.write(f"1. Realiza el Yape o Plin de **S/ 15.00** al número: **{NUMERO_YAPE_PLIN}**")
        st.write("2. Envía la captura del pago por WhatsApp.")
        st.write("3. Te enviaremos tu **PIN de Acceso Mensual** al instante.")
        
        # Enlace directo a WhatsApp
        mensaje_wa = "Hola, deseo suscribirme a PlanificaEF por S/ 15 soles al mes. Adjunto mi voucher de pago."
        link_wa = f"https://wa.me/{NUMERO_WHATSAPP}?text={re.sub(r' ', '%20', mensaje_wa)}"
        
        st.markdown(f'''
            <a href="{link_wa}" target="_blank">
                <button style="background-color:#25D366; color:white; border:none; padding:12px 20px; border-radius:8px; font-weight:bold; cursor:pointer; width:100%;">
                    📲 Solicitar PIN por WhatsApp
                </button>
            </a>
        ''', unsafe_allow_html=True)

    with col_login:
        st.markdown("### 🔑 Ingresar PIN de Acceso")
        st.write("Si ya realizaste tu pago y tienes tu PIN, ingrésalo aquí:")
        pin_usuario = st.text_input("Ingresa tu PIN mensual:", type="password", key="input_pin")
        
        if st.button("🔓 Desbloquear Plataforma", use_container_width=True):
            if pin_usuario in PINES_ACTIVOS:
                st.session_state["autenticado"] = True
                st.success("¡PIN Correcto! Bienvenido a PlanificaEF.")
                st.rerun()
            else:
                st.error("❌ PIN incorrecto o vencido. Si aún no te has suscrito, solicita tu PIN por WhatsApp.")

    st.markdown("---")
    st.info("💡 Con tu suscripción de S/ 15/mes generas Unidades de Aprendizaje, Sesiones diarias y Rúbricas oficiales del CNEB listas para descargar en Word.")
    
    # Detiene la ejecución para no mostrar las pestañas ni gastar API de IA si no ha pagado
    st.stop()

# ==============================================================================
# CONTENIDO DE LA PLATAFORMA (SOLO SE MUESTRA SI YA INGRESÓ SU PIN VÁLIDO)
# ==============================================================================
st.title("🏃‍♂️ PlanificaEF")
st.caption("✅ Sesión Activa de Suscriptor | Para cerrar sesión borra la caché del navegador.")

# Enlace automático a la clave secreta guardada de forma segura
api_key = st.secrets.get("GEMINI_API_KEY", None)

if not api_key:
    st.error("⚠️ No se encontró la GEMINI_API_KEY en los secretos de Streamlit.")

# Funciones auxiliares de Word e IA
def limpiar_texto(texto):
    if not texto:
        return ""
    return texto.replace('||', '|\n|')

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
                            est_u_list.append(f"• COMPETENCIA: {comp_name}\nESTÁNDAR OFICIAL: \"{est_val}\"")
                        if des_list:
                            des_u_list.append(f"• COMPETENCIA: {comp_name}\nDESEMPEÑOS CON NUMERACIÓN OFICIAL CNEB PARA {grado_u}:\n" + "\n".join(des_list))
                    matriz_estandares_u = "\n\n".join(est_u_list)
                    matriz_desempenos_u = "\n\n".join(des_u_list)

                instrucciones_u = f"""Actúa como un Especialista Curricular experto en Educación Física para Primaria bajo el enfoque del CNEB de Perú (MINEDU).
Diseña una UNIDAD DE APRENDIZAJE completa estructurada EN TABLAS MARKDOWN CON LÍNEAS SEPARADORAS (|---|---|).

MATRIZ OFICIAL EXACTA Y PALABRA POR PALABRA EXTRAÍDA DE CNEB_DATOS.PY PARA {grado_u} ({ciclo_u}):

ESTÁNDARES OFICIALES DEL CICLO ({ciclo_u}):
{matriz_estandares_u}

DESEMPEÑOS OFICIALES DEL GRADO ({grado_u}):
{matriz_desempenos_u}

ESTRUCTURA DE LA UNIDAD DE APRENDIZAJE:

1. DATOS INFORMATIVOS:
Genera esta TABLA EXACTA en Markdown con línea separadora:
| Campo | Detalle |
| :--- | :--- |
| **DRE / UGEL** | DRE [.....] / UGEL [.....] |
| **Institución Educativa** | I.E. N° [.....] |
| **Lugar / Localidad** | [.....] |
| **Ciclo** | {ciclo_u} |
| **Grado y Sección** | {grado_u}, Secciones: [.....] |
| **Docente del Área** | [.....] |
| **Director(a)** | [.....] |
| **Duración y Periodo** | {duracion_u} (Del [Día/Mes] al [Día/Mes]) |

2. TÍTULO DE LA UNIDAD DE APRENDIZAJE (Significativo, retador e innovador).

3. SITUACIÓN SIGNIFICATIVA (Contexto del problema, Reto en pregunta y Producto de la unidad).

4. PROPÓSITOS DE APRENDIZAJE Y SECUENCIA DE SESIONES:
Genera esta TABLA de 7 COLUMNAS con línea separadora:
| ACTIVIDAD (SESIÓN) | DESCRIPCIÓN PEDAGÓGICA | COMPETENCIA / CAPACIDADES | ESTÁNDAR DE LA COMPETENCIA | DESEMPEÑO PRECISADO | CRITERIOS DE EVALUACIÓN | INSTRUMENTO DE EVALUACIÓN |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |

REGLAS ABSOLUTAS Y ESTRICTAS DE TRANSCRIPCIÓN DE CNEB_DATOS.PY:
- En 'ESTÁNDAR DE LA COMPETENCIA': Copia PALABRA POR PALABRA el Estándar Oficial proporcionado arriba para la competencia correspondiente. Queda PROHIBIDO refrasear o cambiar palabras. Únicamente inserta **negrita** (`**texto**`) sobre las palabras del estándar original que se ejercitan en esa sesión.
- En 'DESEMPEÑO PRECISADO': Copia PALABRA POR PALABRA el Desempeño Oficial del CNEB de la lista de {grado_u} provista arriba, conservando su NUMERACIÓN ORIGINAL EXACTA (ejemplo: `1.1.-`, `1.2.-`, `2.1.-`, `3.1.-`). PROHIBIDO alterar o resumir las palabras originales del CNEB. Inserta **negrita** (`**texto en negrita**`) ÚNICAMENTE en dos partes: 1) la frase tomada del CNEB original que se ejercita, y 2) lo que le agregas al final para precisarlo con el tema de la sesión.
- En 'CRITERIOS DE EVALUACIÓN': Formula OBLIGATORIAMENTE EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN por cada sesión usando `<br>` para separarlos. La redacción debe contemplar de forma implícita los 3 elementos (Acción + Contenido + Condición) sin etiquetas explícitas.
- En 'INSTRUMENTO DE EVALUACIÓN': Lista de cotejo / Rúbrica.

5. ENFOQUES TRANSVERSALES PRIORIZADOS (Tabla con línea separadora).
6. MATERIALES Y RECURSOS DIDÁCTICOS."""

                pedido_u = f"Crea una unidad para {grado_u} con duración de {duracion_u}. Contexto del problema: {problema_u}"
                
                resultado_u = generar_respuesta_ia(client, instrucciones_u, pedido_u)
                resultado_u = limpiar_texto(resultado_u)
                
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

                instrucciones = f"""Actúa como un docente experto de Educación Física de nivel Primaria en Perú, especialista en el enfoque por competencias del CNEB de MINEDU.

MATRIZ OFICIAL LITERAL EXTRAÍDA DIRECTAMENTE DE CNEB_DATOS.PY:
- Grado y Ciclo: {grado_s} ({ciclo_s})
- Competencia principal: {competencia_s}
- Tema / Propósito motriz: {detalles_s}
- ESTÁNDAR CNEB OFICIAL LITERAL PARA {ciclo_s}: "{estandar_base}"
- DESEMPEÑOS CNEB OFICIALES CON NUMERACIÓN EXACTA PARA {grado_s} Y PARA LA COMPETENCIA "{competencia_s}":
{desempenos_base}

ESTRUCTURA OBLIGATORIA A GENERAR EN TABLAS CON LÍNEAS SEPARADORAS MARKDOWN (|---|---|):

# SESIÓN DE APRENDIZAJE N°.......
**Título:** [Crea un título motivador, lúdico e innovador sobre {detalles_s}]

## 1. DATOS INFORMATIVOS
Genera esta TABLA EXACTA de 2 columnas con línea separadora:
| Campo | Detalle |
| :--- | :--- |
| **I.E N°** | {ie_s} |
| **Grado y Sección** | {grado_s} |
| **Area** | Educación fisica |
| **Docente** | {docente_s} |
| **Fecha** | {fecha_s} |
| **tiempo** | {tiempo_s} |

## 2. PROPÓSITOS Y EVIDENCIAS DE APRENDIZAJE
Genera esta TABLA EXACTA de 6 COLUMNAS en una sola fila por cada competencia con línea separadora:
| COMPETENCIA / CAPACIDADES | ESTÁNDAR CNEB | DESEMPEÑOS PRECISADO | CRITERIOS DE EVALUACIÓN | EVIDENCIA Y PRODUCTO | INSTRUMENTO DE EVALUACIÓN |
| :--- | :--- | :--- | :--- | :--- | :--- |

REGLAS ABSOLUTAS Y ESTRICTAS DE TRANSCRIPCIÓN DE CNEB_DATOS.PY:
- **Columna 1:** Transcribe la competencia ({competencia_s}) y sus capacidades oficiales separadas con `<br>`.
- **Columna 2 (ESTÁNDAR CNEB):** Transcribe PALABRA POR PALABRA Y DE MANERA COMPLETA el estándar oficial proporcionado arriba ("{estandar_base}"). Queda estrictamente PROHIBIDO refrasear, cambiar palabras o usar puntos suspensivos '...'. Únicamente **resalta en negrita (**texto**)** la frase exacta del estándar original que se aplica directamente en la clase de hoy.
- **Columna 3 (DESEMPEÑO PRECISADO):** Selecciona el desempeño oficial del CNEB de la lista de {grado_s} arriba provista. CONSERVA SU NUMERACIÓN OFICIAL EXACTA (ejemplo: `1.1.-`, `1.2.-`, `2.1.-`, `3.1.-`), transcribe PALABRA POR PALABRA el texto original de cneb_datos.py sin modificar ninguna de sus palabras y **resalta en negrita (**texto en negrita**)** únicamente dos partes: 1) la frase o acción tomada del CNEB original que se ejercita hoy, y 2) la adición de la precisión agregada al final para el tema ({detalles_s}).
- **Columna 4 (CRITERIOS DE EVALUACIÓN):** Formula OBLIGATORIAMENTE EXACTAMENTE 3 CRITERIOS DE EVALUACIÓN separados con `<br>` dentro de la celda. Cada criterio debe contener de forma implícita los 3 elementos pedagógicos (Acción + Contenido + Condición) en una oración fluida, PERO QUEDA ESTRICTAMENTE PROHIBIDO ESCRIBIR O ETIQUETAR LAS PALABRAS '(Acción)' O '(Contenido)' EXPLÍCITAMENTE.
- **Columna 5:** Define la Evidencia de aprendizaje (producto o actuación medible).
- **Columna 6:** Lista de cotejo.

## 3. ENFOQUE TRANSVERSAL
Genera esta TABLA con línea separadora:
| Enfoque Transversal Priorizado | Valor | Actitud / Comportamiento Observable |
| :--- | :--- | :--- |

## 4. PREPARACIÓN DE LA SESIÓN
Genera esta TABLA de ESTRICTAMENTE 2 COLUMNAS con línea separadora:
| ¿Qué necesitamos hacer antes de la sesión? | Recursos o Materiales a utilizar |
| :--- | :--- |

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
Diseña una TABLA de **Lista de Cotejo** con línea separadora con los 3 criterios de evaluación planteados al inicio y filas para nombres de estudiantes."""

                pedido = f"Diseña una sesión para {grado_s}. Competencia: {competencia_s}. Detalles del tema: {detalles_s}"
                
                resultado_s = generar_respuesta_ia(client, instrucciones, pedido)
                resultado_s = limpiar_texto(resultado_s)
                
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
                instrucciones_r = f"""Actúa como un Evaluador Pedagógico experto en Educación Física para Primaria.
Diseña una rúbrica analítica estructurada con los niveles: En Inicio, En Proceso, Logrado y Logro Destacado para el desempeño solicitado, utilizando exactamente 3 criterios claros y observables alineados al CNEB sin etiquetar explícitamente '(Acción)' ni '(Contenido)'."""

                pedido_r = f"Crea una rúbrica para {grado_r}. Competencia: {competencia_r}. Desempeño: {criterio_r}"
                
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

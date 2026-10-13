import io
import os
import json
import urllib.request
import urllib.parse
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt

# Configuración de la interfaz web
st.set_page_config(
    page_title="Generador Profesional de Presentaciones",
    page_icon="🎓",
    layout="centered",
)

# --- SISTEMA DE LOGIN SENCILLO ---
CREDENCIALES = {"profesor": "profe2026", "estudiante": "clase2026"}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔐 Acceso a la Plataforma Educativa")
    st.write("Inicia sesión para generar tus presentaciones y recursos didácticos.")

    with st.form("login_form"):
        user = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Ingresar")

        if entrar:
            if user in CREDENCIALES and CREDENCIALES[user] == password:
                st.session_state.autenticado = True
                st.session_state.usuario = user
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
    st.stop()

# --- APLICACIÓN PRINCIPAL ---
st.sidebar.success(f"Conectado como: **{st.session_state.usuario}**")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.autenticado = False
    st.rerun()

st.title("🎓 Generador Universal de Presentaciones Académicas")
st.write(
    "Escribe **cualquier tema, asignatura o concepto** y la app buscará "
    "información real en internet para estructurar una presentación con "
    "datos veraces, imágenes ilustrativas y enlaces educativos."
)

with st.form("form_presentacion"):
    tema = st.text_input(
        "¿Sobre qué tema, asignatura o concepto deseas la presentación?",
        "La Inteligencia Artificial",
    )
    col1, col2 = st.columns(2)
    with col1:
        num_slides = st.slider("Número de diapositivas", min_value=3, max_value=8, value=4)
    with col2:
        incluir_yt = st.checkbox("Incluir enlaces de videos de YouTube", value=True)

    generar_btn = st.form_submit_button("🚀 Generar y Descargar Presentación")


# --- MOTOR DE BÚSQUEDA Y GENERACIÓN DE CONTENIDO REAL ---
def buscar_info_wikipedia(query):
    """Busca información resumida e imagen en Wikipedia de forma gratuita."""
    url = f"https://es.wikipedia.org/api/rest_v1/page/summary/{urllib.parse.quote(query)}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'GeneradorPresentaciones/1.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            return {
                'titulo': data.get('title', query),
                'extracto': data.get('extract', 'Información no disponible en este subtema.'),
                'imagen_url': data.get('thumbnail', {}).get('source')
            }
    except Exception:
        return {'titulo': query, 'extracto': 'No se encontró información específica en fuentes abiertas.', 'imagen_url': None}


def generar_contenido_unico(query, total_slides):
    slides = []
    
    # Subtemas dinámicos para enriquecer la búsqueda y evitar repetición
    subtemas = [
        f"{query} introducción",
        f"{query} historia fundamentos",
        f"{query} aplicaciones practicas",
        f"{query} ventajas beneficios",
        f"{query} retos desafios",
        f"{query} futuro tendencias",
        f"{query} metodologia",
        f"{query} casos de estudio"
    ]
    
    for i in range(total_slides):
        subtema = subtemas[i % len(subtemas)]
        info = buscar_info_wikipedia(subtema)
        
        # Si el extracto es muy corto, buscamos el tema general como respaldo
        if len(info['extracto']) < 150:
            info = buscar_info_wikipedia(query)
            
        # Limpieza y formato del título
        titulo_limpio = info['titulo'].replace(query, "").strip().strip(":")
        if not titulo_limpio or titulo_limpio.lower() == query.lower():
            titulo = f"Aspectos Clave: {query}" if i > 0 else f"Introducción a {query}"
        else:
            titulo = f"{i+1}. {titulo_limpio}"
            
        slides.append({
            'titulo': titulo,
            'texto': info['extracto'],
            'imagen_url': info['imagen_url']
        })

    # Enlaces de YouTube optimizados para contenido educativo
    busqueda_url = urllib.parse.quote(query)
    videos = [
        {
            "titulo": f"Clase magistral y explicación detallada: {query}",
            "url": f"https://www.youtube.com/results?search_query={busqueda_url}+clase+magistral+educativa"
        },
        {
            "titulo": f"Guía práctica, tutoriales y ejemplos visuales: {query}",
            "url": f"https://www.youtube.com/results?search_query={busqueda_url}+explicacion+facil+ejemplos"
        },
    ]
    
    return slides, videos


# --- GENERACIÓN DEL ARCHIVO POWERPOINT ---
if generar_btn:
    if not tema.strip():
        st.warning("Por favor, escribe un tema válido.")
    else:
        barra = st.progress(0)
        estado = st.empty()

        try:
            estado.text(f"🔍 Buscando información real y verificada sobre '{tema}'...")
            barra.progress(30)
            
            contenido_slides, lista_videos = generar_contenido_unico(tema, num_slides)

            estado.text("📊 Compilando archivo PowerPoint con diseño optimizado...")
            barra.progress(60)

            prs = Presentation()

            for slide_info in contenido_slides:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = slide_info['titulo']
                
                # Insertar texto con ajuste de palabra
                tf = slide.placeholders[1].text_frame
                tf.text = slide_info['texto']
                tf.word_wrap = True

                # --- LÓGICA DE IMÁGENES INTELIGENTE ---
                img_url = slide_info['imagen_url']
                if not img_url:
                    # Fallback: Generador de imágenes IA gratuito y sin API Key (Pollinations)
                    tema_img = urllib.parse.quote(tema)
                    img_url = f"https://image.pollinations.ai/prompt/{tema_img}%20educational%20presentation%20professional?width=600&height=400&nologo=true"
                
                try:
                    resp_img = requests.get(img_url, timeout=8)
                    if resp_img.status_code == 200:
                        image_stream = io.BytesIO(resp_img.content)
                        # Posición: lado derecho, equilibrado con el texto
                        left = Inches(5.8)
                        top = Inches(2.2)
                        width = Inches(3.5)
                        slide.shapes.add_picture(image_stream, left, top, width=width)
                except Exception:
                    pass  # Si falla la imagen, la diapositiva se mantiene limpia solo con texto

            # Diapositiva final con enlaces directos a YouTube
            if incluir_yt and lista_videos:
                slide_yt = prs.slides.add_slide(prs.slide_layouts[1])
                slide_yt.shapes.title.text = "🎥 Recursos y Videos en YouTube"
                detalle_yt = "Enlaces directos recomendados para profundizar en la clase:\n\n"
                for v in lista_videos:
                    detalle_yt += f"• {v['titulo']}\n  🔗 {v['url']}\n\n"
                slide_yt.placeholders[1].text = detalle_yt

            barra.progress(100)
            estado.text("¡Completado con éxito!")

            archivo_salida = f"presentacion_{os.getpid()}.pptx"
            prs.save(archivo_salida)

            with open(archivo_salida, "rb") as f:
                st.success("🎉 ¡Tu presentación profesional está lista con **contenido real de internet**!")
                st.download_button(
                    label="📥 Descargar Presentación (.pptx)",
                    data=f,
                    file_name=f"{tema.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            if os.path.exists(archivo_salida):
                os.remove(archivo_salida)

        except Exception as err:
            st.error(f"Ocurrió un error durante la generación: {err}")

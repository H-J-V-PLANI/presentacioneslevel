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
st.set_page_config(page_title="Generador Profesional de Presentaciones", page_icon="🎓", layout="centered")

# --- SISTEMA DE LOGIN SENCILLO ---
CREDENCIALES = {"profesor": "profe2026", "estudiante": "clase2026"}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔐 Acceso a la Plataforma Educativa")
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

st.sidebar.success(f"Conectado como: **{st.session_state.usuario}**")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.autenticado = False
    st.rerun()

st.title("🎓 Generador de Presentaciones con Búsqueda Web Real")
st.write("Escribe un tema y el sistema buscará **artículos reales en internet**, extraerá la información clave y generará tu PowerPoint.")

with st.form("form_presentacion"):
    tema = st.text_input("Tema de la presentación", "Inteligencia Artificial en la educación")
    num_slides = st.slider("Número de diapositivas", min_value=3, max_value=8, value=4)
    generar_btn = st.form_submit_button("🚀 Buscar en Internet y Generar")

# --- MOTOR DE BÚSQUEDA WEB REAL (GRATUITO Y SIN API KEY) ---
def buscar_en_internet(query):
    """Usa el endpoint público de Jina AI para buscar y extraer texto limpio de la web."""
    url = f"https://s.jina.ai/{urllib.parse.quote(query)}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0', 'X-Return-Format': 'text'})
        with urllib.request.urlopen(req, timeout=15) as response:
            texto_completo = response.read().decode('utf-8')
            
            # Limpiar el texto: Jina devuelve "Title: ...\nURL: ...\n\nContenido..."
            lineas = texto_completo.split('\n')
            titulo = "Información encontrada"
            contenido = []
            
            for linea in lineas:
                if linea.startswith("Title:"):
                    titulo = linea.replace("Title:", "").strip()
                elif linea.startswith("URL:") or linea.startswith("Published:") or linea.startswith("Description:"):
                    continue # Saltar metadatos
                elif linea.strip() and not linea.startswith("Image:"):
                    contenido.append(linea.strip())
            
            texto_limpio = " ".join(contenido)[:1200] # Máximo 1200 caracteres para que quepa en la diapositiva
            
            if len(texto_limpio) < 50:
                return None
                
            return {'titulo': titulo, 'extracto': texto_limpio}
    except Exception:
        return None

def buscar_wikipedia_robusto(query):
    """Wikipedia de respaldo: primero busca el título exacto, luego el extracto."""
    try:
        # Paso 1: Buscar el título correcto
        search_url = f"https://es.wikipedia.org/w/api.php?action=query&list=search&srsearch={urllib.parse.quote(query)}&srlimit=1&utf8=&format=json"
        with urllib.request.urlopen(search_url, timeout=5) as response:
            data = json.loads(response.read().decode())
            if not data['query']['search']:
                return None
            page_title = data['query']['search'][0]['title']
            
            # Paso 2: Obtener el extracto del título exacto
            extract_url = f"https://es.wikipedia.org/w/api.php?action=query&prop=extracts|pageimages&exintro=1&explaintext=1&pithumbsize=600&titles={urllib.parse.quote(page_title)}&format=json"
            with urllib.request.urlopen(extract_url, timeout=5) as response2:
                data2 = json.loads(response2.read().decode())
                pages = data2['query']['pages']
                page_id = list(pages.keys())[0]
                if page_id == '-1':
                    return None
                
                return {
                    'titulo': page_title,
                    'extracto': pages[page_id].get('extract', 'Sin resumen disponible.')[:1200],
                    'imagen_url': pages[page_id].get('thumbnail', {}).get('source')
                }
    except Exception:
        return None

def generar_contenido(tema_principal, total_slides):
    slides = []
    # Subtemas inteligentes para variar la búsqueda
    subtemas = [
        tema_principal,
        f"{tema_principal} historia origen",
        f"{tema_principal} aplicaciones practicas",
        f"{tema_principal} ventajas beneficios",
        f"{tema_principal} desafios problemas",
        f"{tema_principal} futuro tendencias"
    ]
    
    for i in range(total_slides):
        subtema = subtemas[i % len(subtemas)]
        
        # Intentar primero búsqueda web real, si falla, usar Wikipedia robusto
        info = buscar_en_internet(subtema)
        if not info:
            info = buscar_wikipedia_robusto(subtema)
            
        if not info:
            info = {'titulo': f"Aspectos de {tema_principal}", 'extracto': f"Información detallada sobre {subtema} en el contexto actual.", 'imagen_url': None}

        slides.append({
            'titulo': f"{i+1}. {info['titulo']}" if i > 0 else f"Introducción: {info['titulo']}",
            'texto': info['extracto'],
            'imagen_url': info.get('imagen_url')
        })
        
    return slides

# --- GENERACIÓN DEL POWERPOINT ---
if generar_btn:
    if not tema.strip():
        st.warning("Por favor, escribe un tema válido.")
    else:
        barra = st.progress(0)
        estado = st.empty()

        try:
            estado.text("🌐 Buscando información real en internet (esto puede tomar unos segundos)...")
            barra.progress(30)
            
            contenido_slides = generar_contenido(tema, num_slides)
            
            estado.text("📊 Diseñando las diapositivas...")
            barra.progress(70)

            prs = Presentation()
            
            for slide_info in contenido_slides:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = slide_info['titulo']
                
                tf = slide.placeholders[1].text_frame
                tf.text = slide_info['texto']
                tf.word_wrap = True

                # Lógica de imagen: Usa la de Wikipedia o genera una con IA gratuita (Pollinations)
                img_url = slide_info.get('imagen_url')
                if not img_url:
                    tema_img = urllib.parse.quote(tema)
                    img_url = f"https://image.pollinations.ai/prompt/{tema_img}%20professional%20educational%20presentation?width=600&height=400&nologo=true"
                
                try:
                    resp_img = requests.get(img_url, timeout=10)
                    if resp_img.status_code == 200:
                        slide.shapes.add_picture(io.BytesIO(resp_img.content), Inches(5.5), Inches(2.0), width=Inches(3.8))
                except Exception:
                    pass # Si la imagen falla, la diapositiva se guarda solo con texto

            barra.progress(100)
            estado.text("¡Completado con éxito!")

            archivo_salida = f"presentacion_{os.getpid()}.pptx"
            prs.save(archivo_salida)

            with open(archivo_salida, "rb") as f:
                st.success("🎉 ¡Presentación generada con información REAL de internet!")
                st.download_button(
                    label="📥 Descargar Presentación (.pptx)",
                    data=f,
                    file_name=f"{tema.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

            if os.path.exists(archivo_salida):
                os.remove(archivo_salida)

        except Exception as err:
            st.error(f"Ocurrió un error inesperado: {err}")

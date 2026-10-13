import io
import os
import json
import urllib.request
import urllib.parse
import requests
import streamlit as st
from bs4 import BeautifulSoup # Necesitarás: pip install beautifulsoup4
from pptx import Presentation
from pptx.util import Inches, Pt

# ... (Mantén tu sistema de login y configuración inicial igual que antes) ...

# --- MOTOR DE BÚSQUEDA MODULAR ---
def obtener_info_duckduckgo(query):
    """Busca resúmenes rápidos sin API Key usando DuckDuckGo."""
    url = f"https://api.duckduckgo.com/?q={urllib.parse.quote(query)}&format=json&no_html=1&skip_disambig=1"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=5) as response:
            data = json.loads(response.read().decode())
            extracto = data.get('AbstractText', '')
            imagen = data.get('Image', '')
            
            if not extracto:
                return {'titulo': query, 'extracto': 'No se encontró resumen rápido.', 'imagen_url': None}
            
            return {
                'titulo': data.get('Heading', query),
                'extracto': extracto,
                'imagen_url': imagen if imagen.startswith('http') else None
            }
    except Exception:
        return {'titulo': query, 'extracto': 'Error al consultar DuckDuckGo.', 'imagen_url': None}


def obtener_info_por_url(url_objetivo):
    """Extrae el texto principal de una página web específica (Scraping)."""
    try:
        req = urllib.request.Request(url_objetivo, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req, timeout=5).read()
        soup = BeautifulSoup(html, 'html.parser')
        
        # Extrae todos los párrafos y los une (puedes ajustar los selectores según la web)
        parrafos = [p.get_text().strip() for p in soup.find_all('p') if len(p.get_text()) > 50]
        texto_completo = "\n\n".join(parrafos[:5]) # Limita a los primeros 5 párrafos para no saturar
        
        return {
            'titulo': soup.title.string if soup.title else "Contenido Web",
            'extracto': texto_completo,
            'imagen_url': None
        }
    except Exception as e:
        return {'titulo': 'Error', 'extracto': f'No se pudo leer la URL: {e}', 'imagen_url': None}


def generar_contenido_unico(query, total_slides, fuente="duckduckgo", url_personalizada=None):
    slides = []
    subtemas = [
        query,
        f"{query} historia",
        f"{query} aplicaciones",
        f"{query} ventajas",
        f"{query} desafios",
        f"{query} futuro"
    ]
    
    for i in range(total_slides):
        subtema = subtemas[i % len(subtemas)]
        
        # --- AQUÍ SE ELIGE LA FUENTE ---
        if fuente == "url" and url_personalizada:
            info = obtener_info_por_url(url_personalizada)
        elif fuente == "duckduckgo":
            info = obtener_info_duckduckgo(subtema)
        else: # Fallback a Wikipedia si todo lo demás falla
            url_wiki = f"https://es.wikipedia.org/api/rest_v1/page/summary/{urllib.parse.quote(subtema)}"
            try:
                req = urllib.request.Request(url_wiki, headers={'User-Agent': 'App/1.0'})
                with urllib.request.urlopen(req, timeout=5) as response:
                    data = json.loads(response.read().decode())
                    info = {
                        'titulo': data.get('title', subtema),
                        'extracto': data.get('extract', 'Sin información.'),
                        'imagen_url': data.get('thumbnail', {}).get('source')
                    }
            except Exception:
                info = {'titulo': subtema, 'extracto': 'Información no disponible.', 'imagen_url': None}

        slides.append({
            'titulo': f"{i+1}. {info['titulo']}" if i > 0 else f"Introducción: {info['titulo']}",
            'texto': info['extracto'],
            'imagen_url': info['imagen_url']
        })

    return slides


# --- INTERFAZ DE STREAMLIT ACTUALIZADA ---
st.title("🎓 Generador Universal de Presentaciones (Multi-Fuente)")

with st.form("form_presentacion"):
    tema = st.text_input("Tema de la presentación", "Energía Renovable")
    
    # ¡NUEVO! Selector de fuente de información
    fuente_seleccionada = st.radio(
        "¿De dónde quieres obtener la información?",
        options=["DuckDuckGo (Rápido y sin claves)", "Wikipedia (Estable y detallado)", "URL Personalizada (Scraping)"],
        index=0
    )
    
    url_personalizada = ""
    if fuente_seleccionada == "URL Personalizada (Scraping)":
        url_personalizada = st.text_input("Pega la URL del artículo o página web:", "https://ejemplo.com/articulo-educativo")

    num_slides = st.slider("Número de diapositivas", min_value=3, max_value=8, value=4)
    generar_btn = st.form_submit_button("🚀 Generar Presentación")

if generar_btn:
    if not tema.strip():
        st.warning("Por favor, escribe un tema válido.")
    else:
        barra = st.progress(0)
        estado = st.empty()

        try:
            estado.text(f"🔍 Buscando en {fuente_seleccionada}...")
            barra.progress(30)
            
            # Mapeo de la selección del usuario a la lógica del código
            if "DuckDuckGo" in fuente_seleccionada:
                fuente_codigo = "duckduckgo"
            elif "URL" in fuente_seleccionada:
                fuente_codigo = "url"
            else:
                fuente_codigo = "wikipedia"
                
            contenido_slides = generar_contenido_unico(tema, num_slides, fuente=fuente_codigo, url_personalizada=url_personalizada)

            estado.text("📊 Creando diapositivas...")
            barra.progress(70)

            prs = Presentation()
            for slide_info in contenido_slides:
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                slide.shapes.title.text = slide_info['titulo']
                tf = slide.placeholders[1].text_frame
                tf.text = slide_info['texto']
                tf.word_wrap = True

                # Lógica de imagen (usa la de la fuente o genera una con IA si está vacía)
                img_url = slide_info['imagen_url']
                if not img_url:
                    img_url = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(tema)}%20educational?width=600&height=400&nologo=true"
                
                try:
                    resp_img = requests.get(img_url, timeout=8)
                    if resp_img.status_code == 200:
                        slide.shapes.add_picture(io.BytesIO(resp_img.content), Inches(5.8), Inches(2.2), width=Inches(3.5))
                except Exception:
                    pass

            barra.progress(100)
            archivo_salida = f"presentacion_{os.getpid()}.pptx"
            prs.save(archivo_salida)

            with open(archivo_salida, "rb") as f:
                st.success("🎉 ¡Presentación generada con éxito!")
                st.download_button(
                    label="📥 Descargar .pptx",
                    data=f,
                    file_name=f"{tema.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            if os.path.exists(archivo_salida):
                os.remove(archivo_salida)

        except Exception as err:
            st.error(f"Error: {err}")

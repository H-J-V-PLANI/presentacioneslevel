import io
import os
import json
import urllib.request
import urllib.parse
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Generador Multi-Fuente", page_icon="🌐", layout="centered")

# --- LOGIN ---
CREDENCIALES = {"profesor": "profe2026", "estudiante": "clase2026"}
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔐 Acceso")
    with st.form("login_form"):
        user = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Ingresar")
        if entrar:
            if CREDENCIALES.get(user) == password:
                st.session_state.autenticado = True
                st.session_state.usuario = user
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
    st.stop()

st.sidebar.success(f"Conectado como: **{st.session_state.usuario}**")
if st.sidebar.button("Cerrar Sesión"):
    st.session_state.autenticado = False
    st.rerun()

st.title("🌐 Generador Multi-Fuente de Presentaciones")
st.write("Busca información **real** combinando Wikipedia, DuckDuckGo, Open Library y REST Countries. ¡Cero APIs de pago!")

with st.form("form"):
    tema = st.text_input("Tema:", "La Segunda Guerra Mundial")
    num_slides = st.slider("Diapositivas", 3, 8, 4)
    generar_btn = st.form_submit_button("🚀 Generar")

# --- FUENTES DE DATOS GRATUITAS ---

def fetch_json(url, timeout=5):
    """Helper universal para cualquier API gratuita."""
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'MultiSourceApp/1.0'})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode())
    except Exception:
        return None

def buscar_wikipedia(query):
    data = fetch_json(f"https://es.wikipedia.org/api/rest_v1/page/summary/{urllib.parse.quote(query)}")
    if data and data.get('extract'):
        return {'fuente': 'Wikipedia', 'texto': data['extract'], 
                'imagen': data.get('thumbnail', {}).get('source')}
    return None

def buscar_duckduckgo(query):
    data = fetch_json(f"https://api.duckduckgo.com/?q={urllib.parse.quote(query)}&format=json&no_html=1&skip_disambig=1")
    if data and data.get('AbstractText'):
        return {'fuente': 'DuckDuckGo', 'texto': data['AbstractText'], 
                'imagen': data.get('Image')}
    return None

def buscar_openlibrary(query):
    data = fetch_json(f"https://openlibrary.org/search.json?q={urllib.parse.quote(query)}&limit=1")
    if data and data.get('docs'):
        doc = data['docs'][0]
        texto = f"📚 {doc.get('title', '')} - {doc.get('author_name', ['Autor desconocido'])[0]}\n\n"
        texto += f"Primera publicación: {doc.get('first_publish_year', 'N/A')}\n"
        if 'subject' in doc:
            texto += f"Temas: {', '.join(doc['subject'][:5])}"
        return {'fuente': 'Open Library', 'texto': texto, 'imagen': None}
    return None

def buscar_pais(query):
    data = fetch_json(f"https://restcountries.com/v3.1/name/{urllib.parse.quote(query)}")
    if data and isinstance(data, list):
        p = data[0]
        nombre = p.get('name', {}).get('common', query)
        texto = f"🏳️ País: {nombre}\n\n"
        texto += f"Capital: {p.get('capital', ['N/A'])[0] if 'capital' in p else 'N/A'}\n"
        texto += f"Población: {p.get('population', 'N/A'):,}\n"
        texto += f"Región: {p.get('region', 'N/A')}\n"
        texto += f"Idiomas: {', '.join(p.get('languages', {}).values())}"
        return {'fuente': 'REST Countries', 'texto': texto, 
                'imagen': p.get('flags', {}).get('png')}
    return None

def buscar_inteligente(query, idx):
    """Estrategia multi-fuente: intenta una, si falla usa otra."""
    fuentes = [buscar_wikipedia, buscar_duckduckgo, buscar_openlibrary]
    # Si el tema parece un país, priorizamos REST Countries
    if idx == 0:
        pais = buscar_pais(query)
        if pais: return pais
    
    # Subtemas dinámicos para diversificar las búsquedas
    subtemas = [query, f"{query} historia", f"{query} cultura", 
                f"{query} actualidad", f"{query} economia"]
    subtema = subtemas[idx % len(subtemas)]
    
    for fuente in fuentes:
        resultado = fuente(subtema)
        if resultado and len(resultado.get('texto', '')) > 100:
            return resultado
    
    # Fallback de emergencia
    return {'fuente': 'Sin datos', 'texto': f'Información no disponible para {subtema}.', 'imagen': None}


if generar_btn:
    barra = st.progress(0)
    estado = st.empty()
    try:
        estado.text(f"🔎 Consultando múltiples fuentes gratuitas sobre '{tema}'...")
        barra.progress(30)
        
        slides = []
        fuentes_usadas = set()
        for i in range(num_slides):
            info = buscar_inteligente(tema, i)
            fuentes_usadas.add(info['fuente'])
            slides.append({
                'titulo': f"{i+1}. {info['fuente']} - {tema}",
                'texto': info['texto'],
                'imagen_url': info['imagen_url'] if 'imagen_url' in info else info.get('imagen')
            })
        
        estado.text(f"✨ Se consultaron: {', '.join(fuentes_usadas)}")
        barra.progress(60)
        
        prs = Presentation()
        for s in slides:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = s['titulo']
            slide.placeholders[1].text_frame.text = s['texto']
            
            if s['imagen_url']:
                try:
                    r = requests.get(s['imagen_url'], timeout=8)
                    if r.status_code == 200:
                        slide.shapes.add_picture(io.BytesIO(r.content), 
                                                 Inches(6), Inches(2), width=Inches(3))
                except: pass
        
        archivo = f"pres_{os.getpid()}.pptx"
        prs.save(archivo)
        barra.progress(100)
        
        with open(archivo, "rb") as f:
            st.success(f"🎉 ¡Listo! Se usaron estas fuentes: **{', '.join(fuentes_usadas)}**")
            st.download_button("📥 Descargar", f, f"{tema.replace(' ', '_')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        
        if os.path.exists(archivo): os.remove(archivo)
    except Exception as e:
        st.error(f"Error: {e}")

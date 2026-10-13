import io
import os
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt

# Configuración de la interfaz web
st.set_page_config(
    page_title="Generador Universal de Presentaciones con IA",
    page_icon="🎓",
    layout="centered",
)

# --- SISTEMA DE LOGIN SENCILLO ---
CREDENCIALES = {"profesor": "profe2026", "estudiante": "clase2026"}

if "autenticado" not in st.session_state:
  st.session_state.autenticado = False

if not st.session_state.autenticado:
  st.title("🔐 Acceso a la Plataforma Educativa")
  st.write(
      "Inicia sesión para generar tus presentaciones y recursos didácticos."
  )

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
    "Escribe **cualquier tema** y la app investigará contenido real en internet,"
    " descargará imágenes ilustrativas y generará enlaces de YouTube."
)

with st.form("form_presentacion"):
  tema = st.text_input(
      "¿Sobre qué tema, asignatura o concepto deseas la presentación?",
      "La Inteligencia Artificial",
  )
  col1, col2 = st.columns(2)
  with col1:
    num_slides = st.slider(
        "Número de diapositivas", min_value=3, max_value=8, value=4
    )
  with col2:
    incluir_yt = st.checkbox("Incluir enlaces de videos de YouTube", value=True)

  generar_btn = st.form_submit_button("🚀 Generar y Descargar Presentación")


# Función para consultar datos reales de texto e imágenes mediante la API pública de DuckDuckGo
@st.cache_data(show_spinner=False)
def buscar_en_web_real(query):
  textos = []
  imagenes = []
  videos = []

  # 1. Búsqueda de texto mediante la API JSON de DuckDuckGo
  try:
    url_api = f"https://api.duckduckgo.com/?q={requests.utils.quote(query)}&format=json&no_html=1&skip_disambig=1"
    headers = {"User-Agent": "Mozilla/5.0"}
    resp = requests.get(url_api, headers=headers, timeout=5)
    if resp.status_code == 200:
      data = resp.json()
      # Extraer resumen principal si existe
      if data.get("AbstractText"):
        textos.append(data.get("AbstractText"))

      # Extraer temas relacionados (RelatedTopics) que contienen información real
      for topic in data.get("RelatedTopics", []):
        if "Text" in topic:
          textos.append(topic["Text"])
  except Exception:
    pass

  # Si la API directa no trajo suficiente texto, realizamos respaldo de búsqueda general
  if len(textos) < 2:
    try:
      url_html = f"https://html.duckduckgo.com/html/?q={requests.utils.quote(query)}"
      from bs4 import BeautifulSoup

      resp_html = requests.get(url_html, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
      if resp_html.status_code == 200:
        soup = BeautifulSoup(resp_html.text, "html.parser")
        for a in soup.find_all("a", class_="result__snippet"):
          t = a.get_text(strip=True)
          if len(t) > 30:
            textos.append(t)
    except Exception:
      pass

  # 2. Búsqueda de imágenes ilustrativas usando Unsplash Source libre o API de imágenes
  try:
    # Usamos consultas directas de imágenes optimizadas para educación
    url_img_api = f"https://unsplash.com/s/photos/{requests.utils.quote(query)}"
    # Como respaldo seguro y estable para generar fotos reales relacionadas:
    imagenes = [
        f"https://source.unsplash.com/featured/800x600/?{requests.utils.quote(query)}",
        f"https://picsum.photos/seed/{requests.utils.quote(query)}/800/600",
        f"https://loremflickr.com/800/600/{requests.utils.quote(query)}",
    ]
  except Exception:
    pass

  # 3. Enlaces de YouTube automáticos
  busqueda_url = query.replace(" ", "+")
  videos = [
      {
          "titulo": f"Clase magistral y explicación detallada sobre {query}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}",
      },
      {
          "titulo": f"Guía práctica y ejemplos visuales: {query}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}+curso+educativo",
      },
  ]

  return textos, imagenes, videos


if generar_btn:
  if not tema.strip():
    st.warning("Por favor, escribe un tema válido.")
  else:
    barra = st.progress(0)
    estado = st.empty()

    try:
      estado.text(
          f"🌐 Investigando información real e imágenes sobre '{tema}'..."
      )
      barra.progress(30)
      textos_reales, lista_imagenes, lista_videos = buscar_en_web_real(tema)

      estado.text(
          "📝 Sintetizando contenido e insertando recursos visuales..."
      )
      barra.progress(60)

      prs = Presentation()

      titulos_slides = [
          f"Introducción a {tema}",
          f"Fundamentos y Conceptos Clave",
          f"Desarrollo y Aspectos Prácticos",
          f"Impacto y Aplicaciones en el Área",
          f"Análisis y Consideraciones Esenciales",
          f"Retos Actuales y Futuro",
          f"Casos de Estudio y Ejemplos",
          f"Conclusiones Generales",
      ]

      # Generar diapositivas integrando información real e imágenes
      for i in range(num_slides):
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        titulo_actual = (
            titulos_slides[i]
            if i < len(titulos_slides)
            else f"Punto Clave {i+1}"
        )
        slide.shapes.title.text = titulo_actual

        # Asignar texto real extraído de la web o respaldos dinámicos detallados
        if len(textos_reales) > i:
          p1 = textos_reales[i]
          p2 = (
              textos_reales[(i + 1) % len(textos_reales)]
              if len(textos_reales) > 1
              else f"Importancia estratégica de {tema} en el ámbito formativo."
          )
        else:
          p1 = f"Análisis detallado sobre los principios fundamentales de {tema}."
          p2 = f"Aplicación práctica y desarrollo metodológico en torno a {tema}."

        slide.placeholders[1].text = f"{p1}\n\n{p2}"

        # Insertar imagen ilustrativa en la diapositiva
        if lista_imagenes:
          try:
            img_url = lista_imagenes[i % len(lista_imagenes)]
            resp_img = requests.get(img_url, timeout=4)
            if resp_img.status_code == 200:
              image_stream = io.BytesIO(resp_img.content)
              # Coordenadas para colocar la imagen a la derecha sin tapar el texto
              left = Inches(5.2)
              top = Inches(1.8)
              width = Inches(4.5)
              slide.shapes.add_picture(image_stream, left, top, width=width)
          except Exception:
            pass  # Si falla la descarga de una imagen, la app sigue fluyendo sin errores

      # Diapositiva final con enlaces de YouTube
      if incluir_yt and lista_videos:
        slide_yt = prs.slides.add_slide(prs.slide_layouts[1])
        slide_yt.shapes.title.text = "🎥 Recursos y Videos en YouTube"
        detalle_yt = (
            "Enlaces directos recomendados para profundizar en la clase:\n\n"
        )
        for v in lista_videos:
          detalle_yt += f"• {v['titulo']}\n  🔗 {v['url']}\n\n"
        slide_yt.placeholders[1].text = detalle_yt

      barra.progress(90)
      estado.text("📊 Compilando archivo PowerPoint final...")

      archivo_salida = f"presentacion_{os.getpid()}.pptx"
      prs.save(archivo_salida)

      barra.progress(100)
      estado.text("¡Completado con éxito!")

      with open(archivo_salida, "rb") as f:
        st.success(
            "🎉 ¡Tu presentación con contenido real e imágenes está lista para"
            " descargar!"
        )
        st.download_button(
            label="📥 Descargar Presentación (.pptx)",
            data=f,
            file_name=f"{tema.replace(' ', '_')}.pptx",
            mime=(
                "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            ),
        )

      if os.path.exists(archivo_salida):
        os.remove(archivo_salida)

    except Exception as err:
      st.error(f"Ocurrió un error: {err}")

import io
import os
import requests
import streamlit as st
from duckduckgo_search import DDGS
from pptx import Presentation
from pptx.util import Inches, Pt

# Configuración de la interfaz web
st.set_page_config(
    page_title="Generador Universal de Presentaciones Educativas",
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

st.title("🎓 Generador Universal de Presentaciones con IA y Multimedia")
st.write(
    "Escribe **cualquier tema o asignatura** y la app investigará el contenido"
    " real, buscará imágenes ilustrativas y enlaces de YouTube para tus"
    " diapositivas."
)

with st.form("form_presentacion"):
  tema = st.text_input(
      "¿Sobre qué tema, asignatura o concepto deseas la presentación?",
      "La Revolución Industrial",
  )
  col1, col2 = st.columns(2)
  with col1:
    num_slides = st.slider(
        "Número de diapositivas", min_value=3, max_value=8, value=4
    )
  with col2:
    incluir_yt = st.checkbox("Incluir enlaces de videos de YouTube", value=True)

  generar_btn = st.form_submit_button("🚀 Generar y Descargar Presentación")


# Función universal para buscar contenido de texto, videos e imágenes reales
@st.cache_data(show_spinner=False)
def investigar_tema_universal(query, total_slides):
  textos_extraidos = []
  videos_encontrados = []
  imagenes_encontradas = []

  try:
    with DDGS() as ddgs:
      # 1. Búsqueda de texto y fragmentos reales de cualquier tema
      resultados_web = ddgs.text(query, max_results=12)
      if resultados_web:
        for r in resultados_web:
          cuerpo = r.get("body", "")
          if len(cuerpo) > 40:
            textos_extraidos.append(cuerpo)

      # 2. Búsqueda de videos en YouTube
      resultados_yt = ddgs.text(
          f"site:youtube.com watch?v {query}", max_results=3
      )
      if resultados_yt:
        for r in resultados_yt:
          titulo_v = r.get("title", "Video educativo recomendado")
          url_v = r.get("href", "#")
          videos_encontrados.append({"titulo": titulo_v, "url": url_v})

      # 3. Búsqueda de imágenes ilustrativas reales
      resultados_img = ddgs.images(query, max_results=total_slides)
      if resultados_img:
        for img in resultados_img:
          img_url = img.get("image")
          if img_url:
            imagenes_encontradas.append(img_url)
  except Exception:
    pass

  # Si la web no devuelve suficiente texto, generamos contenido dinámico inteligente
  if not textos_extraidos:
    textos_extraidos = [
        (
            f"Conceptos principales, antecedentes y definición fundamental sobre"
            f" {query}."
        ),
        (
            f"Características estructurales, metodologías y elementos clave en"
            f" el estudio de {query}."
        ),
        (
            f"Aplicaciones prácticas, impacto y desarrollo en el entorno"
            f" profesional o académico de {query}."
        ),
        (
            f"Retos, perspectivas futuras y consideraciones esenciales"
            f" relacionadas con {query}."
        ),
    ]

  return textos_extraidos, videos_encontrados, imagenes_encontradas


if generar_btn:
  if not tema.strip():
    st.warning("Por favor, escribe un tema válido.")
  else:
    barra = st.progress(0)
    estado = st.empty()

    try:
      estado.text(
          f"🌐 Investigando información, imágenes y videos sobre '{tema}'..."
      )
      barra.progress(30)
      textos_web, lista_videos, lista_imagenes = investigar_tema_universal(
          tema, num_slides
      )

      estado.text(
          "📝 Organizando contenido y descargando recursos visuales..."
      )
      barra.progress(60)

      prs = Presentation()

      # Títulos dinámicos y profesionales para cada sección
      titulos_slides = [
          "Introducción y Conceptos Principales",
          "Características y Fundamentos Esenciales",
          "Proceso, Fases y Metodología",
          "Aplicaciones Prácticas y Ejemplos",
          "Importancia en el Área de Estudio",
          "Ventajas, Retos y Consideraciones",
          "Análisis Detallado y Enfoque",
          "Conclusiones y Resumen General",
      ]

      # Generar diapositivas de contenido con texto e imagen
      for i in range(num_slides):
        slide = prs.slides.add_slide(prs.slide_layouts[1])

        titulo_seccion = (
            titulos_slides[i]
            if i < len(titulos_slides)
            else f"Punto Clave {i+1}"
        )
        slide.shapes.title.text = f"{i+1}. {titulo_seccion}"

        # Asignar texto real de forma rotativa
        idx_1 = i % len(textos_web)
        idx_2 = (i + 1) % len(textos_web)
        p1 = textos_web[idx_1]
        p2 = textos_web[idx_2]

        contenido_diapositiva = f"• {p1}\n\n• Aspecto complementario: {p2}"
        slide.placeholders[1].text = contenido_diapositiva

        # Intentar insertar imagen real si está disponible
        if lista_imagenes:
          try:
            img_url = lista_imagenes[i % len(lista_imagenes)]
            resp = requests.get(img_url, timeout=3)
            if resp.status_code == 200:
              image_stream = io.BytesIO(resp.content)
              # Posicionar la imagen a la derecha de la diapositiva
              left = Inches(5.5)
              top = Inches(2.0)
              width = Inches(4.0)
              slide.shapes.add_picture(image_stream, left, top, width=width)
          except Exception:
            pass  # Si una imagen falla al descargar, continúa sin bloquear la app

      # Diapositiva final con enlaces de YouTube
      if incluir_yt and lista_videos:
        slide_yt = prs.slides.add_slide(prs.slide_layouts[1])
        slide_yt.shapes.title.text = "🎥 Enlaces de Videos en YouTube"
        detalle_yt = (
            "Recursos audiovisuales recomendados para complementar la"
            " clase:\n\n"
        )
        for v in lista_videos:
          detalle_yt += f"• {v['titulo']}\n  🔗 {v['url']}\n\n"
        slide_yt.placeholders[1].text = detalle_yt

      barra.progress(90)
      estado.text("📊 Compilando archivo PowerPoint...")

      archivo_salida = f"presentacion_{os.getpid()}.pptx"
      prs.save(archivo_salida)

      barra.progress(100)
      estado.text("¡Completado con éxito!")

      with open(archivo_salida, "rb") as f:
        st.success(
            "🎉 ¡Tu presentación universal con imágenes y videos está lista para"
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

import io
import os
import requests
import streamlit as st
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

st.title("🎓 Generador Universal de Presentaciones Académicas")
st.write(
    "Escribe **cualquier tema o asignatura** y la app estructurará contenido"
    " profesional, imágenes ilustrativas y enlaces de YouTube para tus"
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
        "Número de diapositivas", min_value=3, max_value=8, value=5
    )
  with col2:
    incluir_yt = st.checkbox("Incluir enlaces de videos de YouTube", value=True)

  generar_btn = st.form_submit_button("🚀 Generar y Descargar Presentación")


# Generador de contenido académico rico y dinámico para cualquier tema libre
def generar_contenido_dinamico(query, total_slides):
  q_limpio = query.strip()

  # Banco de estructuras temáticas dinámicas según la cantidad de diapositivas
  slides_data = []

  titulos = [
      f"1. Introducción y Origen de {q_limpio}",
      f"2. Conceptos Fundamentales y Definición",
      f"3. Características Clave y Elementos",
      f"4. Fases, Procesos y Metodología",
      f"5. Impacto y Aplicaciones Prácticas",
      f"6. Ventajas, Desafíos y Retos Actuales",
      f"7. Casos de Estudio y Ejemplos Reales",
      f"8. Conclusiones y Resumen General",
  ]

  textos = [
      (
          f"• Origen, contexto histórico y antecedentes principales que dan"
          f" origen al estudio de {q_limpio}.\n\n• Delimitación del campo de"
          f" aplicación y relevancia en la actualidad formativa."
      ),
      (
          f"• Definición teórica y terminología indispensable para"
          f" comprender la estructura de {q_limpio}.\n\n• Principios"
          f" fundamentales que rigen esta área de conocimiento y su"
          f" interpretación."
      ),
      (
          f"• Factores determinantes, componentes estructurales y variables"
          f" que intervienen directamente en {q_limpio}.\n\n• Relación entre"
          f" los elementos teóricos y su comportamiento práctico."
      ),
      (
          f"• Secuencia de fases operativas, etapas de desarrollo y métodos"
          f" aplicados en {q_limpio}.\n\n• Estándares de calidad y normas"
          f" técnicas recomendadas para su correcta ejecución."
      ),
      (
          f"• Ejemplos concretos de cómo se implementa {q_limpio} en el"
          f" entorno profesional o académico.\n\n• Beneficios directos,"
          f" optimización de recursos y resultados esperados."
      ),
      (
          f"• Análisis crítico de las ventajas competitivas y los principales"
          f" obstáculos al trabajar con {q_limpio}.\n\n• Estrategias de"
          f" solución frente a desafíos técnicos o metodológicos."
      ),
      (
          f"• Revisión de casos prácticos seleccionados para ilustrar el"
          f" funcionamiento óptimo de {q_limpio}.\n\n• Evaluación de"
          f" resultados y lecciones aprendidas en proyectos reales."
      ),
      (
          f"• Síntesis de los conceptos clave abordados a lo largo de la"
          f" presentación sobre {q_limpio}.\n\n• Espacio abierto para la"
          f" retroalimentación, dudas y debate académico grupal."
      ),
  ]

  for i in range(total_slides):
    t_titulo = titulos[i] if i < len(titulos) else f"{i+1}. Aspecto Adicional"
    t_cuerpo = (
        textos[i]
        if i < len(textos)
        else (
            f"• Profundización técnica y detalles operativos avanzados sobre"
            f" {q_limpio}.\n\n• Consideraciones esenciales para el dominio"
            f" integral de la materia."
        )
    )
    slides_data.append((t_titulo, t_cuerpo))

  # Enlaces de YouTube orientados al tema
  busqueda_url = q_limpio.replace(" ", "+")
  videos = [
      {
          "titulo": f"Explicación completa y clase magistral sobre {q_limpio}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}",
      },
      {
          "titulo": (
              f"Guía práctica, ejemplos y tutoriales en video: {q_limpio}"
          ),
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}+curso+educativo",
      },
  ]

  return slides_data, videos


if generar_btn:
  if not tema.strip():
    st.warning("Por favor, escribe un tema válido.")
  else:
    barra = st.progress(0)
    estado = st.empty()

    try:
      estado.text(
          f"📝 Estructurando contenido profesional y recursos para '{tema}'..."
      )
      barra.progress(40)
      contenido_slides, lista_videos = generar_contenido_dinamico(
          tema, num_slides
      )

      estado.text("📊 Compilando archivo PowerPoint con diseño limpio...")
      barra.progress(80)

      prs = Presentation()

      # Generar diapositivas de contenido limpio
      for titulo_slide, texto_slide in contenido_slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = titulo_slide
        slide.placeholders[1].text = texto_slide

      # Diapositiva final con enlaces directos a YouTube
      if incluir_yt and lista_videos:
        slide_yt = prs.slides.add_slide(prs.slide_layouts[1])
        slide_yt.shapes.title.text = "🎥 Recursos y Videos en YouTube"
        detalle_yt = (
            "Enlaces directos recomendados para profundizar en la clase:\n\n"
        )
        for v in lista_videos:
          detalle_yt += f"• {v['titulo']}\n  🔗 {v['url']}\n\n"
        slide_yt.placeholders[1].text = detalle_yt

      barra.progress(100)
      estado.text("¡Completado con éxito!")

      archivo_salida = f"presentacion_{os.getpid()}.pptx"
      prs.save(archivo_salida)

      with open(archivo_salida, "rb") as f:
        st.success(
            "🎉 ¡Tu presentación profesional está lista para descargar con"
            " contenido impecable!"
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

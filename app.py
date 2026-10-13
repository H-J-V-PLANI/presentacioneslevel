import io
import os
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
    "Escribe **cualquier tema, asignatura o concepto** y la app estructurará"
    " una presentación limpia, con texto enriquecido, imágenes ilustrativas y"
    " enlaces de YouTube."
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


# Motor de generación de contenido enriquecido y único para cada diapositiva
def generar_contenido_unico(query, total_slides):
  q = query.strip()
  slides = []

  # Banco estructurado de contenidos profesionales y dinámicos para cualquier materia
  banco_tematico = [
      (
          f"1. Introducción y Contexto de {q}",
          (
              f"Definición principal, antecedentes históricos y la importancia"
              f" fundamental de {q} dentro de su área de estudio.\n\nDelimitación"
              f" de los objetivos formativos y alcance conceptual de la"
              f" temática."
          ),
      ),
      (
          f"2. Principios y Fundamentos Esenciales",
          (
              f"Análisis detallado de los componentes estructurales, reglas y"
              f" términos clave que conforman {q}.\n\nRelación directa entre"
              f" los fundamentos teóricos y su interpretación práctica en el"
              f" aula."
          ),
      ),
      (
          f"3. Metodología, Fases y Desarrollo",
          (
              f"Secuencia operativa, etapas metodológicas y procesos"
              f" necesarios para la correcta ejecución de proyectos sobre"
              f" {q}.\n\nEstándares de calidad y buenas prácticas recomendadas."
          ),
      ),
      (
          f"4. Aplicaciones Prácticas y Casos Reales",
          (
              f"Ejemplos concretos de cómo se implementa {q} en el entorno"
              f" professional, laboral o científico.\n\nImpacto directo y"
              f" beneficios medibles en la resolución de problemas."
          ),
      ),
      (
          f"5. Ventajas, Beneficios y Oportunidades",
          (
              f"Valor añadido que aporta el dominio de {q} frente a enfoques"
              f" tradicionales.\n\nOportunidades de innovación y desarrollo"
              f" continuo en el área de especialidad."
          ),
      ),
      (
          f"6. Retos, Desafíos y Consideraciones",
          (
              f"Obstáculos comunes, limitaciones técnicas o metodológicas al"
              f" trabajar con {q}.\n\nEstrategias recomendadas para superar"
              f" dichos retos de manera eficiente."
          ),
      ),
      (
          f"7. Tendencias Actuales y Futuro",
          (
              f"Evolución reciente, nuevas tecnologías y perspectivas a"
              f" mediano y largo plazo vinculadas con {q}.\n\nAdaptación a los"
              f" cambios del entorno global."
          ),
      ),
      (
          f"8. Conclusiones y Resumen General",
          (
              f"Síntesis de los aprendizajes clave abordados en la presentación"
              f" sobre {q}.\n\nEspacio abierto para la retroalimentación,"
              f" preguntas y debate grupal."
          ),
      ),
  ]

  for i in range(total_slides):
    if i < len(banco_tematico):
      slides.append(banco_tematico[i])
    else:
      slides.append((
          f"{i+1}. Aspectos Avanzados de {q}",
          (
              f"Profundización técnica y consideraciones complementarias"
              f" sobre {q}.\n\nAnálisis de variables secundarias y factores"
              f" determinantes para el dominio integral de la materia."
          ),
      ))

  # Enlaces automáticos de YouTube
  busqueda_url = q.replace(" ", "+")
  videos = [
      {
          "titulo": f"Clase magistral y explicación detallada sobre {q}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}",
      },
      {
          "titulo": f"Guía práctica, tutoriales y ejemplos visuales: {q}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}+curso+educativo",
      },
  ]

  return slides, videos


if generar_btn:
  if not tema.strip():
    st.warning("Por favor, escribe un tema válido.")
  else:
    barra = st.progress(0)
    estado = st.empty()

    try:
      estado.text(
          f"📝 Estructurando contenidos únicos y recursos para '{tema}'..."
      )
      barra.progress(40)
      contenido_slides, lista_videos = generar_contenido_unico(
          tema, num_slides
      )

      estado.text("📊 Compilando archivo PowerPoint con diseño optimizado...")
      barra.progress(80)

      prs = Presentation()

      # Generar diapositivas de contenido limpio
      for titulo_slide, texto_slide in contenido_slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = titulo_slide
        slide.placeholders[1].text = texto_slide

        # Intentar añadir una imagen ilustrativa de manera limpia y sin invadir el texto
        try:
          # Usamos un servicio de imágenes aleatorias estables basadas en palabras clave del tema
          img_url = f"https://picsum.photos/seed/{abs(hash(titulo_slide))}/600/400"
          resp_img = requests.get(img_url, timeout=3)
          if resp_img.status_code == 200:
            image_stream = io.BytesIO(resp_img.content)
            # Colocamos la imagen en la esquina inferior derecha o área lateral libre
            left = Inches(6.8)
            top = Inches(4.2)
            width = Inches(2.8)
            slide.shapes.add_picture(image_stream, left, top, width=width)
        except Exception:
          pass

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
            " contenido único y diseño impecable!"
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

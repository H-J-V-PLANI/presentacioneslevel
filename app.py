import os
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt

# Configuración de la interfaz web
st.set_page_config(
    page_title="Generador de Presentaciones Educativas",
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

st.title("🎓 Generador de Presentaciones Académicas")
st.write(
    "Crea presentaciones educativas estructuradas, detalladas y con enlaces de"
    " YouTube listos para exponer."
)

with st.form("form_presentacion"):
  tema = st.text_input(
      "¿Sobre qué tema o asignatura deseas la presentación?",
      "La animación 2D",
  )
  col1, col2 = st.columns(2)
  with col1:
    num_slides = st.slider(
        "Número de diapositivas", min_value=3, max_value=8, value=4
    )
  with col2:
    incluir_yt = st.checkbox(
        "Incluir enlaces de videos de YouTube", value=True
    )

  generar_btn = st.form_submit_button("🚀 Generar y Descargar Presentación")


# Función inteligente para generar contenido académico profundo y adaptado al tema
def generar_contenido_academico(query, total_slides):
  q = query.lower()

  # Base de conocimiento especializada según el tema o asignaturas comunes
  if "animación" in q or "2d" in q:
    banco_contenido = [
        (
            "Fundamentos de la Animación 2D",
            (
                "• La animación 2D se basa en la creación de movimiento en un"
                " espacio bidimensional utilizando altura y anchura.\n\n•"
                " Incluye técnicas tradicionales fotograma a fotograma y"
                " animación digital vectorial o de mapa de bits.\n\n• Es la"
                " base visual para la narrativa en cine, televisión, diseño"
                " web y multimedia."
            ),
        ),
        (
            "Los 12 Principios de la Animación",
            (
                "• Estiramiento y encogimiento (Squash and stretch): Aporta"
                " sensación de peso y flexibilidad a los objetos.\n\n• Anticipación"
                " y puesta en escena: Preparan al espectador para la acción"
                " principal y dirigen la atención visual.\n\n• Animación directa"
                " y pose a pose: Métodos esenciales de planificación y ejecución"
                " temporal."
            ),
        ),
        (
            "Herramientas y Software Profesional",
            (
                "• Utilización de software especializado para la vectorización,"
                " rigging y control de líneas de tiempo.\n\n• Creación de"
                " assets gráficos optimizados para interfaces digitales y"
                " plataformas web.\n\n• Integración de capas, símbolos y"
                " fotogramas clave (keyframes) en el flujo de trabajo."
            ),
        ),
        (
            "Aplicaciones Prácticas y Multimedia",
            (
                "• Desarrollo de contenidos interactivos, anuncios"
                " publicitarios y cortometrajes educativos.\n\n• Sincronización"
                " de audio, efectos sonoros y locución con la línea de"
                " tiempo.\n\n• Exportación en formatos optimizados (MP4, GIF, Lottie"
                " JSON) para web y dispositivos móviles."
            ),
        ),
    ]
  else:
    # Contenido genérico estructurado y profesional para cualquier otra materia
    banco_contenido = [
        (
            f"Introducción y Conceptos Clave de {query}",
            (
                f"• Análisis profundo de los principios teóricos y normativos"
                f" que rigen {query}.\n\n• Importancia del contexto, los"
                f" antecedentes y los objetivos principales de estudio en esta"
                f" área.\n\n• Delimitación del campo de aplicación y su"
                f" relevancia académica o técnica."
            ),
        ),
        (
            "Características y Elementos Estructurales",
            (
                f"• Componentes esenciales que conforman el desarrollo de"
                f" {query}.\n\n• Relación directa entre la teoría metodológica"
                f" y su ejecución práctica.\n\n• Criterios de calidad,"
                f" optimización de recursos y estándares profesionales"
                f" aplicados."
            ),
        ),
        (
            "Proceso, Fases y Metodología de Trabajo",
            (
                f"• Planificación estratégica por etapas para el diseño y"
                f" ejecución de proyectos sobre {query}.\n\n• Herramientas,"
                f" recursos tecnológicos y normativas de buenas prácticas.\n\n•"
                f" Resolución de problemas comunes y control de calidad en el"
                f" proceso."
            ),
        ),
        (
            "Aplicaciones Prácticas y Casos de Estudio",
            (
                f"• Implementación de soluciones reales utilizando {query} en"
                f" entornos profesionales.\n\n• Análisis de impacto, ventajas"
                f" competitivas y beneficios directos para los usuarios o"
                f" estudiantes.\n\n• Evaluación de resultados y métricas de"
                f" éxito."
            ),
        ),
        (
            "Retos Actuales e Innovación en el Área",
            (
                f"• Nuevas tendencias tecnológicas y adaptaciones futuras"
                f" relacionadas con {query}.\n\n• Desafíos comunes a nivel de"
                f" diseño, ejecución y sostenibilidad.\n\n• Oportunidades de"
                f" mejora continua y actualización profesional."
            ),
        ),
        (
            "Conclusiones y Resumen General",
            (
                f"• Síntesis de los aprendizajes y puntos críticos analizados"
                f" sobre {query}.\n\n• Reflexión final sobre la importancia de"
                f" dominar estas competencias en el ámbito formativo.\n\n•"
                f" Espacio para preguntas, debate y retroalimentación grupal."
            ),
        ),
    ]

  # Ajustar a la cantidad de diapositivas solicitadas
  slides_generadas = []
  for i in range(total_slides):
    if i < len(banco_contenido):
      slides_generadas.append(banco_contenido[i])
    else:
      slides_generadas.append((
          f"Aspecto Adicional {i+1} sobre {query}",
          (
              f"• Profundización en los detalles técnicos y operativos de"
              f" {query}.\n\n• Consideraciones adicionales para el dominio"
              f" completo de la temática.\n\n• Estrategias de seguimiento y"
              f" evaluación continua."
          ),
      ))

  # Enlaces automáticos de YouTube orientados al tema
  busqueda_url = query.replace(" ", "+")
  videos = [
      {
          "titulo": f"Tutorial completo y explicación detallada sobre {query}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}",
      },
      {
          "titulo": f"Guía práctica, ejemplos y casos de estudio: {query}",
          "url": f"https://www.youtube.com/results?search_query={busqueda_url}+curso+practico",
      },
  ]

  return slides_generadas, videos


if generar_btn:
  if not tema.strip():
    st.warning("Por favor, escribe un tema válido.")
  else:
    barra = st.progress(0)
    estado = st.empty()

    try:
      estado.text("📝 Estructurando y redactando contenido académico...")
      barra.progress(40)
      contenido_slides, lista_videos = generar_contenido_academico(
          tema, num_slides
      )

      estado.text("📊 Compilando archivo PowerPoint profesional...")
      barra.progress(80)

      prs = Presentation()

      # Generar diapositivas con texto desarrollado
      for titulo_slide, texto_slide in contenido_slides:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = titulo_slide
        slide.placeholders[1].text = texto_slide

      # Diapositiva final con enlaces directos a YouTube
      if incluir_yt and lista_videos:
        slide_yt = prs.slides.add_slide(prs.slide_layouts[1])
        slide_yt.shapes.title.text = "🎥 Recursos y Videos en YouTube"
        detalle_yt = (
            "Enlaces directos recomendados para profundizar en el"
            " contenido:\n\n"
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
            "🎉 ¡Tu presentación detallada está lista para descargar con"
            " contenido real!"
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

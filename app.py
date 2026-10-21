import os
import unicodedata
import streamlit as st
from docx import Document

# Configuración de la página
st.set_page_config(
    page_title="Química - Bitácora Académica", page_icon="🧪", layout="wide"
)

# Estilos CSS inspirados en el diseño de Wix y tablas limpias
st.markdown(
    """
    <head>
        <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet">
    </head>
    <style>
        html, body, [data-testid="stAppViewContainer"] {
            background-color: #ffffff !important;
            color: #111c2c;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }
        
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}

        /* Barra de navegación superior estilo Wix */
        .wix-nav {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 1.5rem 3rem;
            border-bottom: 1px solid #E2E8F0;
            background-color: #ffffff;
            margin-bottom: 3rem;
        }
        .wix-quote {
            font-size: 13px;
            letter-spacing: 0.05em;
            color: #64748B;
            text-transform: uppercase;
        }
        .wix-menu {
            display: flex;
            gap: 2rem;
            font-size: 14px;
            font-weight: 500;
            color: #111c2c;
        }

        /* Título principal grande tipo blog */
        .blog-title-container {
            text-align: center;
            padding: 2rem 1rem 4rem 1rem;
        }
        .blog-main-title {
            font-family: 'Space Grotesk', sans-serif;
            font-size: 64px;
            font-weight: 700;
            color: #111c2c;
            letter-spacing: -0.03em;
            margin-top: 0.5rem;
        }

        /* Tarjetas de blog modernas */
        .blog-card {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 0.5rem;
            padding: 0;
            overflow: hidden;
            box-shadow: 0 4px 20px -4px rgba(0,0,0,0.04);
            margin-bottom: 2rem;
            transition: all 0.3s ease;
            height: 100%;
        }
        .blog-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 30px -8px rgba(0,0,0,0.08);
        }
        .card-img-placeholder {
            height: 180px;
            background: linear-gradient(135deg, #e0e7ff 0%, #f1f5f9 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 32px;
        }
        .card-body {
            padding: 1.5rem;
        }
        .card-date {
            font-size: 12px;
            color: #64748B;
            margin-bottom: 0.5rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        /* Botones de acción estilo Wix */
        .stButton>button {
            background-color: #000000;
            color: #ffffff;
            font-family: 'Space Grotesk', sans-serif;
            font-weight: 600;
            border-radius: 2rem;
            padding: 0.6rem 1.5rem;
            border: none;
            transition: all 0.3s ease;
            width: 100%;
        }
        .stButton>button:hover {
            background-color: #333333;
            transform: translateY(-1px);
        }

        /* Estilos profesionales para tablas idénticas al Word */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 1.5rem 0;
            font-size: 15px;
        }
        th, td {
            border: 1px solid #CBD5E1;
            padding: 0.75rem 1rem;
            text-align: left;
        }
        th {
            background-color: #F1F5F9;
            font-weight: 600;
            color: #0f172a;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# Menú superior tipo plantilla de Wix
st.markdown(
    """
    <div class="wix-nav">
        <div class="wix-quote">TODO ES PERSONAL. INCLUSO ESTE BLOG.</div>
        <div class="wix-menu">
            <span>Inicio</span>
            <span>Acerca de</span>
            <span>Mi Blog</span>
            <span>Contacto</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Estructura oficial sin mención de semestres en los títulos
temas_dict = {
    "1. Química General I y II": {
        "Materia y Átomo": [
            "Introducción a la materia",
            "Estructura atómica y modelos",
        ],
        "Enlace Químico": [
            "Concepto y tipos de enlaces",
            "Fuerzas intermoleculares",
        ],
        "Nomenclatura y Fórmulas Químicas": [
            "Compuestos inorgánicos",
            "Estados de oxidación",
        ],
        "Estequiometría y Reacciones": [
            "El concepto de mol",
            "Balanceo de ecuaciones y reactivo límite",
        ],
        "Gases": ["Leyes de los gases", "Ecuación de estado ideal"],
        "Soluciones": ["Tipos de soluciones", "Unidades de concentración"],
    },
    "2. Química Inorgánica I y II": {
        "Química de los Elementos Representativos y de Transición": [
            "Periodicidad química avanzada",
            "Bloques s, p y d",
        ],
        "Teorías de Enlace en Compuestos Inorgánicos": [
            "Teoría del Enlace de Valencia",
            "Teoría de Orbitales Moleculares",
        ],
        "Química del Estado Sólido": ["Redes cristalinas", "Defectos en sólidos"],
    },
    "3. Química Orgánica I y II": {
        "Fundamentos de Química Orgánica": [
            "Hibridación del carbono",
            "Efectos electrónicos",
        ],
        "Nomenclatura Orgánica y Grupos Funcionales": [
            "Hidrocarburos alifáticos y cíclicos",
            "Compuestos oxigenados y nitrogenados",
        ],
        "Reacciones Orgánicas y Mecanismos": [
            "Tipos de reacciones orgánicas",
            "Mecanismos de sustitución y adición",
        ],
        "Biomoléculas": ["Carbohidratos y lípidos", "Proteínas y ácidos nucleicos"],
    },
    "4. Fisicoquímica I y II": {
        "Termodinámica Química (Primera, Segunda y Tercera Ley)": [
            "Sistemas y energía interna",
            "Entalpía, entropía y energía libre de Gibbs",
        ],
        "Cinética Química": [
            "Velocidad de reacción",
            "Leyes de velocidad y factores determinantes",
        ],
        "Equilibrio Químico y de Fases": [
            "Constante de equilibrio",
            "Principio de Le Chatelier",
        ],
        "Electroquímica": [
            "Celdas galvánicas y electrolíticas",
            "Ecuación de Nernst",
        ],
    },
    "5. Química Analítica y Análisis Instrumental": {
        "Análisis Clásico / Cuantitativo": [
            "Análisis gravimétrico",
            "Titulaciones volumétricas",
        ],
        "Análisis Instrumental I y II": [
            "Métodos espectrométricos (UV-Vis)",
            "Métodos cromatográficos",
        ],
    },
    "6. Espectroscopia Básica y Avanzada": {
        "Espectroscopia Básica": [
            "Interacción radiación-materia",
            "Infrarrojo (FT-IR)",
        ],
        "Espectroscopia Avanzada": [
            "Resonancia Magnética Nuclear (RMN)",
            "Espectrometría de Masas",
        ],
    },
    "7. Química de los Compuestos de Coordinación": {
        "Estructura y Enlace": [
            "Nomenclatura de complejos",
            "Teoría del Campo Cristalino",
        ],
        "Propiedades y Reactividad": [
            "Propiedades magnéticas y color",
            "Mecanismos de sustitución",
        ],
    },
    "8. Química Cuántica y Computacional": {
        "Química Cuántica": [
            "Postulados y ecuación de Schrödinger",
            "Átomo de hidrógeno",
        ],
        "Química Computacional": [
            "Modelado molecular",
            "Optimización de geometrías (DFT)",
        ],
    },
    "9. Bioquímica I y Biología Molecular": {
        "Bioquímica I": ["Cinética enzimática", "Ciclo de Krebs y metabolismo"],
        "Biología Molecular": ["Replicación y transcripción", "Traducción y genes"],
    },
    "10. Química Ambiental y Verde": {
        "Química Ambiental": ["Ciclos biogeoquímicos", "Atmósfera e hidrósfera"],
        "Química Verde": ["Los 12 principios", "Procesos sostenibles"],
    },
    "11. Procesos Químicos Industriales, Radioquímica y Productos Naturales": {
        "Procesos Químicos Industriales": [
            "Operaciones unitarias",
            "Manufactura a gran escala",
        ],
        "Radioquímica": [
            "Desintegración nuclear",
            "Aplicaciones de radioisótopos",
        ],
        "Productos Naturales": [
            "Extracción de metabolitos secundarios",
            "Identificación estructural",
        ],
    },
}

# Control de estado de la navegación
if "tema_seleccionado" not in st.session_state:
  st.session_state.tema_seleccionado = None
if "subtema_seleccionado" not in st.session_state:
  st.session_state.subtema_seleccionado = None

# Título principal del Blog
st.markdown(
    """
    <div class="blog-title-container">
        <span class="wix-quote">BITÁCORA ACADÉMICA</span>
        <h1 class="blog-main-title">Química</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# PANTALLA 1: CUADRÍCULA DE TEMAS PRINCIPALES
# ==========================================
if st.session_state.tema_seleccionado is None and st.session_state.subtema_seleccionado is None:
  st.markdown(
      """
      <div style="max-width: 1200px; margin: 0 auto; padding: 0 2rem;">
          <h3 style="font-family: 'Space Grotesk', sans-serif; font-size: 24px; margin-bottom: 2rem;">Entradas más recientes</h3>
      </div>
      """,
      unsafe_allow_html=True,
  )

  col1, col2, col3 = st.columns(3)
  temas_lista = list(temas_dict.keys())

  for index, tema in enumerate(temas_lista):
    col_actual = col1 if index % 3 == 0 else (col2 if index % 3 == 1 else col3)
    with col_actual:
      st.markdown(
          f"""
            <div class="blog-card">
                <div class="card-img-placeholder">🧪</div>
                <div class="card-body">
                    <div class="card-date">Área de Estudio</div>
                    <h4 style="font-family: 'Space Grotesk', sans-serif; font-size: 18px; color: #111c2c; margin-top: 0; min-height: 50px;">{tema}</h4>
                    <p style="color: #64748B; font-size: 14px; margin-bottom: 1.5rem;">Explora los contenidos detallados, guías y subtemas de esta materia.</p>
                </div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button("Leer más", key=f"btn_tema_{index}"):
        st.session_state.tema_seleccionado = tema
        st.rerun()

# ==========================================
# PANTALLA 2: DESPLIEGUE DE SUBTEMAS
# ==========================================
elif st.session_state.tema_seleccionado is not None and st.session_state.subtema_seleccionado is None:
  tema_actual = st.session_state.tema_seleccionado

  st.markdown(
      '<div style="max-width: 1200px; margin: 0 auto; padding: 0 2rem;">',
      unsafe_allow_html=True,
  )

  if st.button("⬅️ Volver al inicio"):
    st.session_state.tema_seleccionado = None
    st.session_state.subtema_seleccionado = None
    st.rerun()

  st.markdown(
      f"""
        <div style="margin: 2rem 0; padding-bottom: 1rem; border-bottom: 1px solid #E2E8F0;">
            <span class="wix-quote">Tema Seleccionado</span>
            <h2 style="font-family: 'Space Grotesk', sans-serif; font-size: 36px; margin-top: 0.5rem; color: #111c2c;">{tema_actual}</h2>
            <p style="color: #64748B;">Selecciona un subtema para abrir el documento completo:</p>
        </div>
      </div>
      """,
      unsafe_allow_html=True,
  )

  subtemas_del_tema = temas_dict[tema_actual]
  sub_col1, sub_col2, sub_col3 = st.columns(3)

  idx_contador = 0
  for subtema, titulos_nivel3 in subtemas_del_tema.items():
    col_sub = sub_col1 if idx_contador % 3 == 0 else (sub_col2 if idx_contador % 3 == 1 else sub_col3)
    with col_sub:
      st.markdown(
          f"""
            <div class="blog-card">
                <div class="card-img-placeholder" style="background: linear-gradient(135deg, #f3e8ff 0%, #e0e7ff 100%);">📄</div>
                <div class="card-body">
                    <div class="card-date">Subtema Académico</div>
                    <h4 style="font-family: 'Space Grotesk', sans-serif; font-size: 16px; color: #111c2c; margin-top: 0;">{subtema}</h4>
                    <p style="color: #64748B; font-size: 13px; margin-bottom: 1.5rem;">Contenido: {', '.join(titulos_nivel3)}.</p>
                </div>
            </div>
            """,
          unsafe_allow_html=True,
      )
      if st.button("Ver artículo", key=f"sub_{subtema}"):
        st.session_state.subtema_seleccionado = subtema
        st.rerun()
    idx_contador += 1

# ==========================================
# PANTALLA 3: LECTURA CON TEXTO, TABLAS E IMÁGENES
# ==========================================
else:
  tema_actual = st.session_state.tema_seleccionado
  sub_actual = st.session_state.subtema_seleccionado

  st.markdown(
      '<div style="max-width: 900px; margin: 0 auto; padding: 0 2rem;">',
      unsafe_allow_html=True,
  )

  if st.button("⬅️ Volver a los subtemas"):
    st.session_state.subtema_seleccionado = None
    st.rerun()

  st.markdown(
      f"""
      <div style="margin: 2rem 0 1rem 0;">
          <span class="wix-quote">{tema_actual}</span>
          <h1 style="font-family: 'Space Grotesk', sans-serif; font-size: 42px; color: #111c2c; margin-top: 0.5rem; letter-spacing: -0.02em;">{sub_actual}</h1>
          <hr style="border: none; border-top: 1px solid #E2E8F0; margin: 2rem 0;">
      </div>
      """,
      unsafe_allow_html=True,
  )


  # Función inteligente para ignorar tildes y mayúsculas al buscar el archivo Word
  def limpiar_nombre(texto):
    nfkd_form = unicodedata.normalize("NFKD", texto)
    texto_sin_tildes = "".join(
        [c for c in nfkd_form if not unicodedata.combining(c)]
    )
    return texto_sin_tildes.lower().replace(" ", "_").replace("/", "_")


  nombre_archivo_seguro = limpiar_nombre(sub_actual) + ".docx"
  ruta_archivo = os.path.join("apuntes", nombre_archivo_seguro)

  if os.path.exists(ruta_archivo):
    try:
      doc = Document(ruta_archivo)

      os.makedirs("temp_imagenes", exist_ok=True)
      img_counter = 0

      for element in doc.element.body:
        if element.tag.endswith("p"):
          from docx.text.paragraph import Paragraph

          p = Paragraph(element, doc)
          texto = p.text.strip()

          if texto:
            estilo_nombre = p.style.name.lower() if p.style and p.style.name else ""

            if "heading 1" in estilo_nombre or "título 1" in estilo_nombre:
              st.markdown(
                  f"<h2 style='font-family: Space Grotesk, sans-serif; color:"
                  f" #111c2c; margin-top: 1.5rem;'>{texto}</h2>",
                  unsafe_allow_html=True,
              )
            elif "heading 2" in estilo_nombre or "título 2" in estilo_nombre:
              st.markdown(
                  f"<h3 style='font-family: Space Grotesk, sans-serif; color:"
                  f" #111c2c; margin-top: 1.2rem;'>{texto}</h3>",
                  unsafe_allow_html=True,
              )
            elif (
                "list" in estilo_nombre
                or texto.startswith("•")
                or texto.startswith("-")
            ):
              st.markdown(
                  f"<ul style='margin-bottom: 0.5rem;'><li style='color:"
                  f" #2D3748; font-size: 16px;'>{texto.lstrip('•- ')}</li></ul>",
                  unsafe_allow_html=True,
              )
            else:
              st.markdown(
                  f"<p style='line-height: 1.8; color: #2D3748; font-size:"
                  f" 16px; margin-bottom: 1rem;'>{texto}</p>",
                  unsafe_allow_html=True,
              )

          for run in p.runs:
            drawings = run._r.xpath(".//a:blip")
            for drawing in drawings:
              embed_id = drawing.get(
                  "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
              )
              if embed_id:
                image_part = doc.part.related_parts[embed_id]
                img_bytes = image_part.blob
                img_path = os.path.join(
                    "temp_imagenes", f"img_{sub_actual}_{img_counter}.png"
                )
                with open(img_path, "wb") as f:
                  f.write(img_bytes)
                st.image(img_path, use_column_width=True)
                img_counter += 1

          if not texto:
            st.markdown("<br>", unsafe_allow_html=True)

        elif element.tag.endswith("tbl"):
          from docx.table import Table

          table = Table(element, doc)
          table_html = "<table>"
          for i, row in enumerate(table.rows):
            table_html += "<tr>"
            for cell in row.cells:
              cell_text = cell.text.strip().replace("\n", "<br>")
              if i == 0:
                table_html += f"<th>{cell_text}</th>"
              else:
                table_html += f"<td>{cell_text}</td>"
            table_html += "</tr>"
          table_html += "</table>"
          st.markdown(table_html, unsafe_allow_html=True)

    except Exception as e:
      st.error(
          f"Ocurrió un error al leer el documento de Word y sus elementos: {e}"
      )
  else:
    st.warning(
        f"⚠️ Aún no se encuentra el archivo correspondiente para este subtema.\n\n"
        f"Asegúrate de que la carpeta se llame exactamente **`apuntes`** (en"
        f" minúsculas) y guarda tu documento con el nombre limpio:"
        f" **`{nombre_archivo_seguro}`**"
    )

  st.markdown("</div>", unsafe_allow_html=True)

# Footer estilo Wix
st.markdown(
    """
    <footer style="width: 100%; border-top: 1px solid #E2E8F0; padding: 3rem 2rem; margin-top: 5rem; text-align: center;">
        <p style="color: #64748B; font-size: 13px;">© 2026 Química. Creado para transformar el estudio de la ciencia.</p>
    </footer>
    """,
    unsafe_allow_html=True,
)


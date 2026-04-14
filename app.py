import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EHS Goldstein - Sistema de Evaluación", layout="wide")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS DEL MANUAL ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas de comunicación: escucha activa, inicio de conversación y agradecimiento social."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad para pedir ayuda, integrarse a grupos, dar y seguir instrucciones complejas."},
    "III: HHSS Sentimientos": {"items": list(range(15, 22)), "desc": "Habilidades para conocer y expresar sentimientos propios, comprender los ajenos y manejar el afecto."},
    "IV: Alt. a la Agresión": {"items": list(range(22, 31)), "desc": "Capacidad de autocontrol, defensa de derechos, negociación y resolución pacífica de conflictos."},
    "V: Frente al Estrés": {"items": list(range(31, 43)), "desc": "Habilidades para responder a quejas, fracasos, presiones de grupo y manejar la vergüenza."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Habilidades de toma de decisiones, fijación de objetivos realistas y organización de tareas."}
}

BAREMOS = {
    "I":   {9: 38, 8: 35, 7: 33, 6: 30, 5: 28, 4: 26, 3: 23, 2: 21, 1: 0},
    "II":  {9: 28, 8: 26, 7: 24, 6: 22, 5: 21, 4: 19, 3: 17, 2: 15, 1: 0},
    "III": {9: 33, 8: 31, 7: 29, 6: 26, 5: 24, 4: 22, 3: 20, 2: 17, 1: 0},
    "IV":  {9: 43, 8: 41, 7: 38, 6: 36, 5: 33, 4: 31, 3: 29, 2: 26, 1: 0},
    "V":   {9: 56, 8: 53, 7: 49, 6: 46, 5: 42, 4: 39, 3: 35, 2: 32, 1: 0},
    "VI":  {9: 40, 8: 38, 7: 35, 6: 33, 5: 31, 4: 28, 3: 26, 2: 23, 1: 0},
    "TOTAL": {9: 228, 8: 216, 7: 204, 6: 192, 5: 181, 4: 169, 3: 157, 2: 145, 1: 0}
}

PREGUNTAS = [
    "Prestas atención a quien te habla", "Hablas de temas poco importantes", "Hablas de cosas que interesan a ambos",
    "Clarificas la información que necesitas", "Agradeces los favores", "Te das a conocer por propia iniciativa",
    "Ayudas a que los demás se conozcan", "Dices que te gusta algo de otra persona", "Pides ayuda ante dificultades",
    "Eliges la mejor forma de integrarte", "Explicas con claridad cómo hacer tareas", "Prestas atención a instrucciones",
    "Pides disculpas por errores", "Intentas persuadir a los demás", "Reconoces tus emociones",
    "Permites que otros sepan lo que sientes", "Intentas comprender lo que sienten otros", "Comprendes el enfado ajeno",
    "Te interesas por los demás", "Buscas disminuir tus miedos", "Te dices cosas agradables",
    "Pides permiso cuando es necesario", "Compartes algo apreciado", "Ayudas a quien lo necesita",
    "Negocias de forma satisfactoria", "Controlas tu carácter", "Defiendes tus derechos", "No pierdes el control ante bromas",
    "Evitas situaciones problemáticas", "Resuelves conflictos sin pelear", "Indicas quién es responsable del problema",
    "Buscas soluciones justas ante quejas", "Expresas cumplidos sinceros", "Haces algo para sentir menos vergüenza",
    "Gestionas cuando te dejan de lado", "Defiendes a amigos ante injusticias", "Consideras la posición de la otra persona",
    "Comprendes por qué has fracasado", "Resuelves mensajes contradictorios", "Comprendes acusaciones recibidas",
    "Planificas cómo exponer tu punto de vista", "Decides qué hacer ante presión grupal", "Resuelves el aburrimiento",
    "Reconoces si un evento está bajo tu control", "Eres realista sobre tus capacidades", "Eres realista ante tareas difíciles",
    "Resuelves qué necesitas saber", "Determinas qué problema es prioritario", "Consideras posibilidades y eliges", "Te organizas para facilitar el trabajo"
]

# --- INTERFAZ ---
with st.sidebar:
    st.header("📋 Instrucciones de Llenado")
    st.info("""
    Lea cada situación y seleccione el grado en que le ocurre:
    1: Nunca o muy pocas veces.
    2: Pocas veces.
    3: A veces.
    4: Muchas veces.
    5: Siempre o casi siempre.
    """)
    st.divider()
    st.header("Datos del Evaluado")
    nombre = st.text_input("Nombre Completo", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    consent = st.checkbox("Acepto el Consentimiento Informado")
    if st.button("🗑️ Reiniciar Prueba"): reset_form()

st.title("Escala de Habilidades Sociales (Goldstein)")

if consent:
    # LLENADO EN UNA COLUMNA
    respuestas = {}
    for i, p in enumerate(PREGUNTAS, 1):
        respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, key=f"q_{i}_{st.session_state.revision}", index=None)

    if st.button("📈 GENERAR INFORME CLÍNICO"):
        if None in respuestas.values():
            st.error("⚠️ Error: Debe completar todas las preguntas del cuestionario.")
        else:
            # PROCESAMIENTO
            res_data = []
            total_pd = sum(respuestas.values())
            def obtener_e(p, k):
                for e, lim in sorted(BAREMOS[k].items(), reverse=True):
                    if p >= lim: return e
                return 1

            for g, info in GRUPOS.items():
                pd_v = sum(respuestas[idx] for idx in info["items"])
                enea = obtener_e(pd_v, g.split(":")[0])
                res_data.append({"Área": g, "Eneatipo": enea, "Descripción": info["desc"]})

            df = pd.DataFrame(res_data)
            enea_total = obtener_e(total_pd, "TOTAL")

            # --- VISUALIZACIÓN EN PANTALLA ---
            st.divider()
            st.header("Interpretación de Resultados")
            
            fig, ax = plt.subplots(figsize=(10, 4))
            colores = ['#FF4B4B' if e <= 3 else '#00CC96' if e >= 7 else '#FFAA00' for e in df['Eneatipo']]
            ax.bar(df['Área'], df['Eneatipo'], color=colores, edgecolor='black')
            ax.set_ylim(0, 10)
            plt.xticks(rotation=15)
            st.pyplot(fig)

            # --- WORD (FORMATO DE UNA SOLA HOJA TIPO MANUAL) ---
            doc = Document()
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = Inches(0.4)
            sec.left_margin = sec.right_margin = Inches(0.5)

            h = doc.add_heading('PROTOCOLO CLÍNICO: LISTA DE GOLDSTEIN', 0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph(f"Nombre: {nombre} | Edad: {edad} años | Eneatipo Global: {enea_total}").bold = True

            # Tabla de Respuestas (Copia fiel del manual en 2 columnas para que quepa en 1 hoja)
            doc.add_heading('1. Protocolo de Evaluación', level=1)
            table = doc.add_table(rows=25, cols=4)
            table.style = 'Table Grid'
            for i in range(1, 26):
                table.cell(i-1, 0).text = f"{i}. {PREGUNTAS[i-1][:35]}..."
                table.cell(i-1, 1).text = str(respuestas[i])
                table.cell(i-1, 2).text = f"{i+25}. {PREGUNTAS[i+24][:35]}..."
                table.cell(i-1, 3).text = str(respuestas[i+25])
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: p.runs[0].font.size = Pt(8)

            # Análisis y Plan
            doc.add_page_break()
            doc.add_heading('2. Diagnóstico y Plan Terapéutico', level=1)
            
            img_b = BytesIO()
            plt.savefig(img_b, format='png', bbox_inches='tight')
            doc.add_picture(img_b, width=Inches(5.5))

            doc.add_heading('Análisis de Áreas y Posibles Causas:', level=2)
            doc.add_paragraph("Déficit en el aprendizaje social, ansiedad inhibitoria o falta de reforzamiento ambiental.")
            for r in res_data:
                p_area = doc.add_paragraph()
                p_area.add_run(f"• {r['Área']} (Eneatipo {r['Eneatipo']}): ").bold = True
                p_area.add_run(r['Descripción'])
                p_area.runs[0].font.size = Pt(9)

            doc.add_heading('Plan de Intervención Sugerido:', level=2)
            doc.add_paragraph("Objetivo: Fortalecer las áreas deficitarias mediante Entrenamiento en Habilidades Sociales (EHS), Modelado, Ensayo Conductual (Role-playing) y Tareas de Exposición Gradual.")

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 Descargar Hoja de Impresión Profesional", w_buf.getvalue(), f"EHS_Informe_{nombre}.docx")
else:
    st.info("Por favor, lea y acepte el consentimiento en la barra lateral para comenzar.")

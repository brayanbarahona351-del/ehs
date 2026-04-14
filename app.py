import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN CLÍNICA ---
st.set_page_config(page_title="Sistema de Evaluación EHS - Goldstein", layout="centered")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS Y BAREMOS ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas de comunicación: escucha activa, inicio de conversación y agradecimiento."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad para pedir ayuda, integrarse, dar y seguir instrucciones."},
    "III: HHSS de Sentimientos": {"items": list(range(15, 22)), "desc": "Expresión y comprensión de emociones, manejo del afecto y empatía."},
    "IV: Alt. a la Agresión": {"items": list(range(22, 31)), "desc": "Autocontrol, defensa de derechos, negociación y resolución de conflictos."},
    "V: Frente al Estrés": {"items": list(range(31, 43)), "desc": "Resiliencia ante el fracaso, la vergüenza, quejas y presión grupal."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Toma de decisiones, fijación de objetivos y organización de tareas."}
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
    "Prestas atención a la persona que te está hablando...", "Hablas con los demás de temas poco importantes...",
    "Hablas con otras personas sobre cosas que interesan a ambos", "Clarificas la información que necesitas...",
    "Permites que los demás sepan que les agradeces los favores", "Te das a conocer a los demás por propia iniciativa",
    "Ayudas a que los demás se conozcan entre sí", "Dices que te gusta algún aspecto de la otra persona...",
    "Pides que te ayuden cuando tienes alguna dificultad", "Eliges la mejor forma para integrarte en un grupo...",
    "Explicas con claridad a los demás cómo hacer una tarea...", "Prestas atención a las instrucciones...",
    "Pides disculpas a los demás por haber hecho algo mal", "Intentas persuadir a los demás...",
    "Intentas reconocer las emociones que experimentas", "Permites que los demás conozcan lo que sientes",
    "Intentas comprender lo que sienten los demás", "Intentas comprender el enfado de la otra persona",
    "Permites que los demás sepan que te interesas por ellos", "Piensas porqué estás asustado y buscas disminuirlo",
    "Te dices cosas agradables cuando mereces recompensa", "Reconoces cuando es necesario pedir permiso...",
    "Te ofreces para compartir algo apreciado por otros", "Ayudas a quien lo necesita",
    "Estableces un sistema de negociación satisfactorio", "Controlas tu carácter para no perder el control",
    "Defiendes tus derechos dando a conocer tu postura", "Te las arreglas sin perder el control ante bromas",
    "Te mantienes al margen de situaciones problemáticas", "Encuentras formas de resolver conflictos sin pelear",
    "Dices a los demás quién es responsable de un problema", "Intentas llegar a una solución justa ante quejas",
    "Expresas un sincero cumplido por cómo han jugado", "Haces algo para sentir menos vergüenza",
    "Eres consciente si te dejan de lado y buscas sentirte mejor", "Manifiestas si han tratado injustamente a un amigo",
    "Consideras la posición de la otra persona antes de decidir", "Comprendes por qué has fracasado y cómo mejorar",
    "Resuelves la confusión ante mensajes contradictorios", "Comprendes una acusación y cómo relacionarte",
    "Planificas la mejor forma para exponer tu punto de vista", "Decides qué hacer ante la presión grupal",
    "Resuelves el aburrimiento con nuevas actividades", "Reconoces si un evento está bajo tu control",
    "Tomas decisiones realistas sobre tus capacidades", "Eres realista sobre cómo desenvolverte en una tarea",
    "Resuelves qué necesitas saber y cómo conseguir info", "Determinas qué problema es el más importante",
    "Consideras posibilidades y eliges la mejor", "Te organizas y preparas para facilitar tu trabajo"
]

# --- INTERFAZ ---
st.title("Escala de Competencia Social")

with st.sidebar:
    st.header("Identificación")
    nombre = st.text_input("Nombre del Evaluado", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    st.divider()
    if st.button("🗑️ Nueva Evaluación"): reset_form()

# CONSENTIMIENTO E INSTRUCCIONES FIJAS
with st.container(border=True):
    st.subheader("Instrucciones y Consentimiento")
    consent = st.checkbox("Acepto participar voluntariamente y confirmo que he leído las instrucciones.")
    st.markdown("""
    **Guía de Llenado:** Seleccione la opción que mejor describa su comportamiento habitual. 
    * **1:** Nunca / Muy pocas veces.
    * **5:** Siempre / Casi siempre.
    """)

if consent:
    # FORMULARIO EN UNA COLUMNA
    respuestas = {}
    for i, p in enumerate(PREGUNTAS, 1):
        respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, index=None, key=f"q_{i}_{st.session_state.revision}")

    if st.button("📈 GENERAR DIAGNÓSTICO E INFORME"):
        if None in respuestas.values():
            st.error("⚠️ Por favor, responda todos los ítems antes de continuar.")
        else:
            # PROCESAMIENTO
            res_data = []
            total_pd = sum(respuestas.values())
            def get_e(pd, k):
                for e, lim in sorted(BAREMOS[k].items(), reverse=True):
                    if pd >= lim: return e
                return 1

            for g, info in GRUPOS.items():
                pd_g = sum(respuestas[idx] for idx in info["items"])
                enea = get_e(pd_g, g.split(":")[0])
                res_data.append({"Área": g, "PD": pd_g, "Eneatipo": enea, "Descripción": info["desc"]})

            enea_total = get_e(total_pd, "TOTAL")
            
            # --- RESULTADOS CLÍNICOS ---
            st.header("Resultados y Plan Terapéutico")
            df = pd.DataFrame(res_data)
            
            # Gráfico de Perfil
            fig, ax = plt.subplots(figsize=(10, 4))
            colors = ['#ff4b4b' if e <= 3 else '#00cc96' if e >= 7 else '#ffa500' for e in df['Eneatipo']]
            ax.bar(df['Área'], df['Eneatipo'], color=colors, edgecolor='black')
            ax.set_ylim(0, 10)
            plt.xticks(rotation=15)
            st.pyplot(fig)

            # --- ANÁLISIS EXTENSO ---
            st.subheader("Interpretación por Áreas")
            for r in res_data:
                label = "Fortaleza" if r['Eneatipo'] >= 7 else "Deficiencia" if r['Eneatipo'] <= 3 else "Promedio"
                st.write(f"**{r['Área']} ({label}):** {r['Descripción']}")

            # --- PLAN TERAPÉUTICO ---
            st.subheader("🎯 Plan Terapéutico Sugerido")
            deficitarias = [r['Área'] for r in res_data if r['Eneatipo'] <= 3]
            if deficitarias:
                st.write(f"Se requiere intervención prioritaria en: {', '.join(deficitarias)}.")
                st.info("""
                **Objetivos Técnicos:**
                1. Reestructuración cognitiva sobre creencias de incompetencia social.
                2. Entrenamiento en habilidades específicas mediante ensayo conductual y role-playing.
                3. Tareas de exposición gradual en entornos naturales.
                4. Reforzamiento positivo de aproximaciones sucesivas a la conducta meta.
                """)
            else:
                st.success("Competencia social global adecuada. Se sugiere mantenimiento.")

            # --- GENERACIÓN DE WORD (COPIA FIEL) ---
            doc = Document()
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = Inches(0.4)
            sec.left_margin = sec.right_margin = Inches(0.5)

            doc.add_heading('INFORME TÉCNICO: ESCALA DE GOLDSTEIN', 0)
            doc.add_paragraph(f"Evaluado: {nombre} | Edad: {edad} años | Eneatipo Global: {enea_total}").bold = True

            # TABLA QUE IMITA EL FORMULARIO
            doc.add_heading('I. Protocolo de Evaluación (Copia de Llenado)', level=1)
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Ítem / Situación Evaluación'
            hdr_cells[1].text = 'Respuesta (1-5)'

            for i, p in enumerate(PREGUNTAS, 1):
                row_cells = table.add_row().cells
                row_cells[0].text = f"{i}. {p}"
                row_cells[1].text = str(respuestas[i])
                row_cells[1].paragraphs[0].alignment = 1

            for row in table.rows:
                for cell in row.cells:
                    for p_graph in cell.paragraphs: p_graph.runs[0].font.size = Pt(8.5)

            # GRÁFICO E INTERPRETACIÓN
            doc.add_page_break()
            doc.add_heading('II. Perfil Psicológico y Análisis', level=1)
            img_b = BytesIO()
            plt.savefig(img_b, format='png', bbox_inches='tight')
            doc.add_picture(img_b, width=Inches(5.8))

            doc.add_heading('III. Análisis de Causas y Plan Terapéutico', level=1)
            doc.add_paragraph("Posibles Causas: Déficits en la adquisición (falta de modelos), inhibición por ansiedad social o falta de reforzamiento ambiental.")
            
            doc.add_heading('Plan de Intervención:', level=2)
            doc.add_paragraph(f"Prioridad: {', '.join(deficitarias) if deficitarias else 'Mantenimiento'}")
            doc.add_paragraph("Técnicas: Entrenamiento en Habilidades Sociales (EHS), Modelado, Ensayo Conductual y Retroalimentación.")

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 DESCARGAR INFORME E IMPRESIÓN COMPLETA", w_buf.getvalue(), f"Informe_EHS_{nombre}.docx")

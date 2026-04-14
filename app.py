import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE ESTILO ---
st.set_page_config(page_title="EHS Goldstein - Reporte Clínico", layout="wide")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas: escuchar, iniciar conversaciones y dar las gracias."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad para pedir ayuda, integrarse y seguir instrucciones."},
    "III: Sentimientos": {"items": list(range(15, 22)), "desc": "Expresión de afecto, autorrecompensa y empatía emocional."},
    "IV: Alt. Agresión": {"items": list(range(22, 31)), "desc": "Autocontrol, defensa de derechos y negociación de conflictos."},
    "V: Estrés": {"items": list(range(31, 43)), "desc": "Resiliencia ante quejas, fracasos, vergüenza y presión grupal."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Fijar objetivos, toma de decisiones y organización de tareas."}
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

# --- APP ---
st.title("Sistema de Diagnóstico de Habilidades Sociales")

with st.sidebar:
    st.header("Identificación")
    nombre = st.text_input("Nombre", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    if st.button("🗑️ Nueva Evaluación"): reset_form()

consent = st.checkbox("Acepto el consentimiento informado y la veracidad de mis respuestas")

if consent:
    st.write("---")
    st.subheader("Cuestionario Autoinformado")
    respuestas = {}
    cols = st.columns(2)
    for i, p in enumerate(PREGUNTAS, 1):
        with cols[0 if i <= 25 else 1]:
            respuestas[i] = st.radio(f"{i}. {p}", [1,2,3,4,5], horizontal=True, key=f"q_{i}_{st.session_state.revision}")

    if st.button("📈 Generar Reporte Final"):
        # Cálculos y Diagnóstico
        res_data = []
        for g, info in GRUPOS.items():
            pd = sum(respuestas[idx] for idx in info["items"])
            enea = 1
            for e, lim in sorted(BAREMOS[g.split(":")[0]].items(), reverse=True):
                if pd >= lim: 
                    enea = e
                    break
            res_data.append({"Área": g, "Eneatipo": enea, "Descripción": info["desc"]})

        df = pd.DataFrame(res_data)
        enea_total = 1
        total_puntos = sum(respuestas.values())
        for e, lim in sorted(BAREMOS["TOTAL"].items(), reverse=True):
            if total_puntos >= lim:
                enea_total = e
                break

        # --- GENERACIÓN WORD ---
        doc = Document()
        sec = doc.sections[0]
        sec.top_margin = sec.bottom_margin = Inches(0.5)
        sec.left_margin = sec.right_margin = Inches(0.6)

        # Encabezado
        h = doc.add_heading('PROTOCOLO Y RESULTADOS: ESCALA DE GOLDSTEIN', 0)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        info_p = doc.add_paragraph()
        run = info_p.add_run(f"Evaluado: {nombre}  |  Edad: {edad} años  |  Eneatipo Global: {enea_total}")
        run.bold = True
        run.font.size = Pt(11)

        # --- PARTE 1: PROTOCOLO (SEPARADO) ---
        doc.add_heading('I. PROTOCOLO DE LLENADO (Autoinforme)', level=1)
        doc.add_paragraph("Registro exacto de la selección del evaluado:").runs[0].font.size = Pt(9)
        
        tab_prot = doc.add_table(rows=25, cols=4)
        tab_prot.style = 'Table Grid'
        for i in range(1, 26):
            tab_prot.cell(i-1, 0).text = f"{i}. {PREGUNTAS[i-1][:32]}..."
            tab_prot.cell(i-1, 1).text = f"[{respuestas[i]}]"
            tab_prot.cell(i-1, 2).text = f"{i+25}. {PREGUNTAS[i+24][:32]}..."
            tab_prot.cell(i-1, 3).text = f"[{respuestas[i+25]}]"
        
        for row in tab_prot.rows:
            for cell in row.cells:
                for p in cell.paragraphs: p.runs[0].font.size = Pt(8)

        # --- PARTE 2: RESULTADOS (SEPARADO) ---
        doc.add_page_break()
        doc.add_heading('II. RESULTADOS Y DIAGNÓSTICO CLÍNICO', level=1)
        
        # Gráfico
        fig, ax = plt.subplots(figsize=(10, 4))
        colors = ['#d9534f' if e <= 3 else '#5cb85c' if e >= 7 else '#f0ad4e' for e in df['Eneatipo']]
        ax.bar(df['Area' if 'Area' in df else 'Área'], df['Eneatipo'], color=colors, edgecolor='black')
        ax.set_ylim(0, 10)
        plt.xticks(rotation=15)
        
        buf = BytesIO()
        plt.savefig(buf, format='png')
        doc.add_picture(buf, width=Inches(5.5))

        doc.add_heading('III. ANÁLISIS POR ÁREAS Y RECOMENDACIONES', level=1)
        for index, row in df.iterrows():
            ap = doc.add_paragraph()
            ap.add_run(f"• {row['Área']} (Eneatipo {row['Eneatipo']}): ").bold = True
            ap.add_run(row['Descripción'])
            if row['Eneatipo'] <= 3:
                ap.add_run(" [ALERTA: Se requiere intervención]").font.color.rgb = RGBColor(200, 0, 0)
        
        doc.add_heading('IV. POSIBLES CAUSAS Y CONSEJOS', level=1)
        doc.add_paragraph("Causas: Déficits de aprendizaje social, ansiedad social inhibitoria o falta de reforzamiento positivo ambiental.")
        doc.add_paragraph("Consejos: Practicar técnicas de ensayo conductual, relajación y exposición gradual a situaciones sociales.")

        w_buf = BytesIO()
        doc.save(w_buf)
        st.download_button("📥 Descargar Informe Profesional", w_buf.getvalue(), f"Reporte_{nombre}.docx")

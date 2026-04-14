import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EHS Goldstein - Informe Profesional", layout="wide")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas de comunicación inicial."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad de integración y seguimiento de instrucciones."},
    "III: Sentimientos": {"items": list(range(15, 22)), "desc": "Expresión emocional y empatía."},
    "IV: Alt. Agresión": {"items": list(range(22, 31)), "desc": "Autocontrol y resolución de conflictos."},
    "V: Estrés": {"items": list(range(31, 43)), "desc": "Resiliencia ante la presión y el fracaso."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Toma de decisiones y organización."}
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
st.title("📋 Escala de Habilidades Sociales de Goldstein")

with st.sidebar:
    st.header("Identificación")
    nombre = st.text_input("Nombre", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    if st.button("🗑️ Reiniciar Todo"): reset_form()

st.subheader("1. Consentimiento Informado")
consent = st.checkbox("Acepto que mis datos sean procesados para esta evaluación psicológica.")

if consent:
    st.info("**Instrucciones:** Seleccione su respuesta de 1 (Nunca) a 5 (Siempre). Sea honesto/a.")
    
    respuestas = {}
    c1, c2 = st.columns(2)
    for i, p in enumerate(PREGUNTAS, 1):
        with (c1 if i <= 25 else c2):
            respuestas[i] = st.radio(f"{i}. {p}", [1,2,3,4,5], horizontal=True, key=f"q_{i}_{st.session_state.revision}", index=None)

    if st.button("📈 GENERAR REPORTE E IMPRESIÓN"):
        if None in respuestas.values():
            st.error("⚠️ Error: Faltan preguntas por responder.")
        else:
            # PROCESAMIENTO
            res_data = []
            total_puntos = sum(respuestas.values())
            for g, info in GRUPOS.items():
                pd = sum(respuestas[idx] for idx in info["items"])
                k = g.split(":")[0]
                enea = 1
                for e, lim in sorted(BAREMOS[k].items(), reverse=True):
                    if pd >= lim: 
                        enea = e
                        break
                res_data.append({"Área": g, "Eneatipo": enea, "Descripción": info["desc"]})

            df = pd.DataFrame(res_data)
            
            # Gráfico
            fig, ax = plt.subplots(figsize=(10, 4))
            ax.bar(df['Área'], df['Eneatipo'], color='skyblue', edgecolor='black')
            ax.set_ylim(0, 10)
            plt.xticks(rotation=15)
            st.pyplot(fig)

            # --- WORD ---
            doc = Document()
            # Márgenes estrechos para que quepa en una hoja
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = Inches(0.4)
            sec.left_margin = sec.right_margin = Inches(0.5)

            doc.add_heading('PROTOCOLO OFICIAL - ESCALA DE GOLDSTEIN', 0)
            doc.add_paragraph(f"Evaluado: {nombre} | Edad: {edad} años").bold = True

            # Tabla de Protocolo (2 columnas)
            doc.add_heading('Registro de Respuestas (Protocolo)', level=1)
            table = doc.add_table(rows=25, cols=4)
            table.style = 'Table Grid'
            for i in range(1, 26):
                table.cell(i-1, 0).text = f"{i}. {PREGUNTAS[i-1][:35]}..."
                table.cell(i-1, 1).text = str(respuestas[i])
                table.cell(i-1, 2).text = f"{i+25}. {PREGUNTAS[i+24][:35]}..."
                table.cell(i-1, 3).text = str(respuestas[i+25])
            
            # Ajustar fuente de la tabla
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: p.runs[0].font.size = Pt(8)

            # Resultados y Explicación
            doc.add_heading('Interpretación por Áreas', level=1)
            img_b = BytesIO()
            plt.savefig(img_b, format='png')
            doc.add_picture(img_b, width=Inches(5))

            for r in res_data:
                area_p = doc.add_paragraph()
                run = area_p.add_run(f"• {r['Área']} (Enea: {r['Eneatipo']}): ")
                run.bold = True
                area_p.add_run(r['Descripción'])
                area_p.paragraph_format.space_after = Pt(0)
                area_p.runs[0].font.size = Pt(9)

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 Descargar Formulario de Impresión", w_buf.getvalue(), f"EHS_{nombre}.docx")
else:
    st.warning("Debe aceptar el consentimiento para realizar la prueba.")

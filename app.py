import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Evaluación Clínica EHS", layout="centered")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS DEL MANUAL ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas de comunicación: escuchar, iniciar y mantener conversaciones."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad para pedir ayuda, dar instrucciones y seguir reglas sociales."},
    "III: HHSS Sentimientos": {"items": list(range(15, 22)), "desc": "Habilidades para conocer y expresar sentimientos y comprender los ajenos."},
    "IV: Alt. Agresión": {"items": list(range(22, 31)), "desc": "Habilidades de autocontrol, defensa de derechos y negociación de conflictos."},
    "V: Frente al Estrés": {"items": list(range(31, 43)), "desc": "Capacidad para responder a quejas, presiones de grupo y manejar el fracaso."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Habilidades de toma de decisiones, fijar objetivos y organización de tareas."}
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
st.title("📋 Escala de Competencia Social")

with st.sidebar:
    st.header("Identificación")
    nombre = st.text_input("Nombre", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    if st.button("🗑️ Reiniciar Evaluación"): reset_form()

st.subheader("1. Consentimiento Informado")
consent = st.checkbox("Acepto participar voluntariamente y la confidencialidad de los resultados.")

if consent:
    st.subheader("2. Instrucciones")
    st.warning("Marque de 1 (Nunca) a 5 (Siempre) según su conducta habitual. No deje preguntas vacías.")
    
    # LLENADO EN UNA SOLA COLUMNA
    respuestas = {}
    for i, p in enumerate(PREGUNTAS, 1):
        respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, key=f"q_{i}_{st.session_state.revision}", index=None)

    if st.button("📈 GENERAR INFORME PROFESIONAL"):
        if None in respuestas.values():
            st.error("⚠️ Falta responder algunas preguntas.")
        else:
            # Procesamiento de datos
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
            
            # Gráfico interactivo
            df = pd.DataFrame(res_data)
            fig, ax = plt.subplots(figsize=(10, 4))
            colors = ['red' if e <= 3 else 'green' if e >= 7 else 'orange' for e in df['Eneatipo']]
            ax.bar(df['Área'], df['Eneatipo'], color=colors, edgecolor='black')
            ax.set_ylim(0, 10)
            plt.xticks(rotation=15)
            st.pyplot(fig)

            # --- GENERACIÓN DE WORD (FORMATO FORMULARIO ÚNICO) ---
            doc = Document()
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = Inches(0.4)
            sec.left_margin = sec.right_margin = Inches(0.5)

            # Cabecera
            h = doc.add_heading('PROTOCOLO Y DIAGNÓSTICO: LISTA DE GOLDSTEIN', 1)
            h.alignment = 1
            info_p = doc.add_paragraph()
            info_p.add_run(f"Evaluado: {nombre} | Edad: {edad} | Eneatipo Global: {enea_total}").bold = True
            
            # Protocolo Compacto (2 columnas)
            doc.add_heading('Registro de Protocolo (Autoinforme)', level=2)
            table = doc.add_table(rows=25, cols=4)
            table.style = 'Table Grid'
            for i in range(1, 26):
                table.cell(i-1, 0).text = f"{i}. {PREGUNTAS[i-1][:35]}..."
                table.cell(i-1, 1).text = f"[{respuestas[i]}]"
                table.cell(i-1, 2).text = f"{i+25}. {PREGUNTAS[i+24][:35]}..."
                table.cell(i-1, 3).text = f"[{respuestas[i+25]}]"
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: p.runs[0].font.size = Pt(8)

            # Perfil Gráfico
            doc.add_heading('Perfil de Competencia Social', level=2)
            img_b = BytesIO()
            plt.savefig(img_b, format='png', bbox_inches='tight')
            doc.add_picture(img_b, width=Inches(5.5))

            # Explicación de Resultados
            doc.add_heading('Interpretación Técnica por Áreas', level=2)
            for r in res_data:
                ap = doc.add_paragraph()
                ap.add_run(f"{r['Área']} (Eneatipo {r['Eneatipo']}): ").bold = True
                ap.add_run(r['Descripción'])
                ap.paragraph_format.space_after = Pt(0)
                ap.runs[0].font.size = Pt(8.5)

            # Causas y Recomendaciones
            doc.add_heading('Posibles Causas y Recomendaciones', level=2)
            doc.add_paragraph("Causas: Ansiedad inhibitoria, falta de modelos o reforzamiento. Recomendación: Entrenamiento en habilidades específicas.", style='Normal').runs[0].font.size = Pt(9)

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 Descargar Hoja de Impresión", w_buf.getvalue(), f"EHS_Final_{nombre}.docx")
else:
    st.info("Debe aceptar el consentimiento para iniciar.")

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EHS Goldstein Pro - Gráficos", layout="wide")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS ---
GRUPOS = {
    "I: Primeras HHSS": list(range(1, 9)),
    "II: HHSS Avanzadas": list(range(9, 15)),
    "III: Sentimientos": list(range(15, 22)),
    "IV: Alt. Agresión": list(range(22, 31)),
    "V: Estrés": list(range(31, 43)),
    "VI: Planificación": list(range(43, 51))
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
    "Planificas la mejor forma para exponer tu punto de vista", "Decides lo que quieres hacer cuando los demás quieren otra cosa",
    "Resuelves el aburrimiento con actividades nuevas", "Reconoces si un evento está bajo tu control",
    "Tomas decisiones realistas sobre tus capacidades", "Eres realista sobre cómo desenvolverte en una tarea",
    "Resuelves qué necesitas saber y cómo conseguir info", "Determinas qué problema es el más importante",
    "Consideras posibilidades y eliges la mejor", "Te organizas y preparas para facilitar tu trabajo"
]

# --- INTERFAZ ---
st.title("📊 Evaluación de Competencia Social")

with st.sidebar:
    nombre = st.text_input("Nombre Evaluado", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    if st.button("🗑️ Nuevo Llenado"): reset_form()

st.info("Seleccione su respuesta para cada ítem (1: Nunca - 5: Siempre)")

respuestas = {}
for i, p in enumerate(PREGUNTAS, 1):
    respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, index=None, key=f"q_{i}_{st.session_state.revision}")

if st.button("📈 Generar Informe con Gráficos"):
    if None in respuestas.values():
        st.error("Por favor, complete todas las preguntas.")
    else:
        # CÁLCULOS
        res_data = []
        for g, items in GRUPOS.items():
            pd = sum(respuestas[idx] for idx in items)
            key = g.split(":")[0]
            # Función local para eneatipo
            eneatipo = 1
            for e, lim in sorted(BAREMOS[key].items(), reverse=True):
                if pd >= lim: 
                    eneatipo = e
                    break
            res_data.append({"Área": g, "Eneatipo": eneatipo})

        df = pd.DataFrame(res_data)

        # --- GRÁFICO ---
        fig, ax = plt.subplots(figsize=(10, 5))
        colors = ['red' if e <= 3 else 'green' if e >= 7 else 'orange' for e in df['Eneatipo']]
        ax.bar(df['Área'], df['Eneatipo'], color=colors)
        ax.set_ylim(0, 9)
        ax.set_ylabel('Eneatipo')
        ax.set_title(f'Perfil de Habilidades Sociales: {nombre}')
        plt.xticks(rotation=45, ha='right')
        
        # Mostrar en Streamlit
        st.pyplot(fig)
        st.table(df)

        # --- GUARDAR GRÁFICO PARA WORD ---
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', bbox_inches='tight')
        img_buffer.seek(0)

        # --- WORD ---
        doc = Document()
        doc.add_heading('INFORME PSICOLÓGICO DE COMPETENCIA SOCIAL', 0)
        doc.add_paragraph(f"Nombre: {nombre}\nEdad: {edad}")
        
        doc.add_heading('Gráfico de Perfil Psicológico', level=1)
        doc.add_picture(img_buffer, width=Inches(6))
        
        doc.add_heading('Interpretación y Causas', level=1)
        doc.add_paragraph("Los niveles bajos (rojo) sugieren áreas de intervención prioritaria debido a posibles factores de ansiedad social o falta de modelos de aprendizaje.")

        doc.add_heading('Protocolo Detallado', level=1)
        for i, p in enumerate(PREGUNTAS, 1):
            doc.add_paragraph(f"{i}. {p} -> Puntaje: {respuestas[i]}")

        word_buffer = BytesIO()
        doc.save(word_buffer)
        st.download_button("📥 Descargar Informe con Gráficos", word_buffer.getvalue(), f"Informe_{nombre}.docx")

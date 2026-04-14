import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import time

# --- CONFIGURACIÓN TÉCNICA ---
st.set_page_config(page_title="EHS Goldstein Pro", layout="wide")

# Inicializar estado de la aplicación para el botón de borrar
if 'form_id' not in st.session_state:
    st.session_state.form_id = 0

def reiniciar_aplicacion():
    st.session_state.form_id += 1
    st.rerun()

# --- DATOS DEL MANUAL ---
GRUPOS = {
    "I: Primeras Habilidades": list(range(1, 9)),
    "II: Habilidades Avanzadas": list(range(9, 15)),
    "III: Habilidades de Sentimientos": list(range(15, 22)),
    "IV: Alternativas a la Agresión": list(range(22, 31)),
    "V: Frente al Estrés": list(range(31, 43)),
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
    "Planificas la mejor forma para exponer tu punto de vista", "Decides qué hacer frente a la presión grupal",
    "Resuelves el aburrimiento con actividades nuevas", "Reconoces si un evento está bajo tu control",
    "Tomas decisiones realistas sobre lo que eres capaz de realizar", "Eres realista sobre cómo desenvolverte en una tarea",
    "Resuelves qué necesitas saber y cómo conseguir info", "Determinas qué problema es el más importante",
    "Consideras posibilidades y eliges la mejor", "Te organizas y preparas para facilitar tu trabajo"
]

# --- INTERFAZ ---
st.title("🚀 Escala de Habilidades Sociales - Goldstein")

with st.sidebar:
    st.header("Opciones")
    if st.button("🗑️ Borrar y Nuevo Llenado", on_click=reiniciar_aplicacion):
        st.success("Formulario reiniciado.")
    st.divider()
    nombre = st.text_input("Nombre del Evaluado", key=f"name_{st.session_state.form_id}")
    edad = st.number_input("Edad", 12, 90, 20, key=f"age_{st.session_state.form_id}")

st.info("Responda de 1 a 5 (1: Nunca, 5: Siempre)")

# Formulario de preguntas
respuestas = {}
for i, p in enumerate(PREGUNTAS, 1):
    respuestas[i] = st.select_slider(f"{i}. {p}", options=[1, 2, 3, 4, 5], key=f"q_{i}_{st.session_state.form_id}")

if st.button("📊 Generar Informe y Diagnóstico"):
    total_pd = sum(respuestas.values())
    res_finales = []
    
    def get_enea(score, key):
        for e, limit in sorted(BAREMOS[key].items(), reverse=True):
            if score >= limit: return e
        return 1

    for g, items in GRUPOS.items():
        pd_v = sum(respuestas[idx] for idx in items)
        enea = get_enea(pd_v, g.split(":")[0])
        res_finales.append({"Área": g, "PD": pd_v, "Eneatipo": enea})

    enea_total = get_enea(total_pd, "TOTAL")
    
    # --- ANIMACIONES SEGÚN RESULTADOS ---
    if enea_total >= 7:
        st.balloons()
        st.success("¡Excelente! El evaluado posee competencias sociales destacadas.")
    elif enea_total <= 3:
        st.snow()
        st.warning("Atención: Se identifican áreas con déficits significativos que requieren intervención.")
    else:
        st.info("Resultado dentro del promedio normal.")

    st.table(pd.DataFrame(res_finales))
    st.metric("Eneatipo Total", enea_total)

    # --- GENERAR WORD ---
    doc = Document()
    doc.add_heading('INFORME CLÍNICO DE HABILIDADES SOCIALES', 0)
    doc.add_paragraph(f"Nombre: {nombre}\nEdad: {edad}")
    
    doc.add_heading('Protocolo de Respuestas', level=1)
    for i, p in enumerate(PREGUNTAS, 1):
        doc.add_paragraph(f"{i}. {p} -> Respuesta: {respuestas[i]}")

    buffer = BytesIO()
    doc.save(buffer)
    st.download_button("📥 Descargar Informe (Word)", buffer.getvalue(), f"Informe_{nombre}.docx")

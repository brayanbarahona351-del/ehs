import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="EHS Goldstein - Evaluación Profesional", layout="wide")

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
    "Resuelves la sensación de aburrimiento iniciando una nueva actividad", "Reconoces si un evento está bajo tu control",
    "Tomas decisiones realistas sobre tus capacidades", "Eres realista sobre cómo desenvolverte en una tarea",
    "Resuelves qué necesitas saber y cómo conseguir info", "Determinas qué problema es el más importante",
    "Consideras posibilidades y eliges la mejor", "Te organizas y preparas para facilitar tu trabajo"
]

# --- INTERFAZ ---
st.title("⚖️ Sistema de Evaluación: Habilidades Sociales")

with st.sidebar:
    st.header("Control de Sesión")
    if st.button("🗑️ Reiniciar Prueba"): reset_form()
    nombre = st.text_input("Nombre del Evaluado", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")

# 1. CONSENTIMIENTO INFORMADO
st.subheader("1. Consentimiento Informado")
with st.expander("Leer Consentimiento Legal"):
    st.write("""
    Por la presente, acepto participar de manera voluntaria en esta evaluación psicológica. 
    Entiendo que los resultados son orientativos y serán tratados con confidencialidad. 
    Autorizo el procesamiento de mis respuestas para la generación del informe técnico.
    """)
consentimiento = st.checkbox("Acepto los términos y el consentimiento informado", key=f"c_{st.session_state.revision}")

if consentimiento:
    st.subheader("2. Instrucciones de Llenado")
    st.warning("""
    **Instrucciones:** Lea cada situación y seleccione el grado en que le ocurre habitualmente:
    1: Nunca | 2: Pocas veces | 3: A veces | 4: Muchas veces | 5: Siempre.
    No deje ninguna pregunta sin responder.
    """)

    # 3. CUESTIONARIO
    respuestas = {}
    for i, p in enumerate(PREGUNTAS, 1):
        respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, index=None, key=f"q_{i}_{st.session_state.revision}")

    if st.button("📈 GENERAR INFORME CLÍNICO"):
        if None in respuestas.values():
            st.error("⚠️ Debe completar todas las preguntas del cuestionario.")
        else:
            # PROCESAMIENTO
            res_data = []
            total_pd = sum(respuestas.values())
            
            def obtener_e(p, k):
                for e, lim in sorted(BAREMOS[k].items(), reverse=True):
                    if p >= lim: return e
                return 1

            for g, items in GRUPOS.items():
                pd_v = sum(respuestas[idx] for idx in items)
                enea = obtener_e(pd_v, g.split(":")[0])
                res_data.append({"Área": g, "Eneatipo": enea, "PD": pd_v})

            df = pd.DataFrame(res_data)
            enea_total = obtener_e(total_pd, "TOTAL")

            if enea_total >= 7: st.balloons()
            elif enea_total <= 3: st.snow()

            # GRÁFICO
            fig, ax = plt.subplots(figsize=(10, 5))
            colores = ['#FF4B4B' if e <= 3 else '#00CC96' if e >= 7 else '#FFAA00' for e in df['Eneatipo']]
            bars = ax.bar(df['Área'], df['Eneatipo'], color=colores)
            ax.set_ylim(0, 10)
            ax.set_title(f"Perfil de Competencia Social - {nombre}")
            plt.xticks(rotation=45, ha='right')
            st.pyplot(fig)

            # WORD
            doc = Document()
            doc.add_heading('INFORME PSICOLÓGICO: LISTA DE GOLDSTEIN', 0)
            
            # Consentimiento en el Word
            doc.add_heading('Consentimiento e Instrucciones', level=1)
            doc.add_paragraph("El evaluado ha aceptado el consentimiento informado de manera digital.")
            doc.add_paragraph("Instrucciones aplicadas: Escala Likert de 1 a 5 (Nunca a Siempre).")

            # Datos
            doc.add_heading('Datos Generales', level=1)
            doc.add_paragraph(f"Paciente: {nombre}\nEdad: {edad}\nEneatipo Total: {enea_total}")

            # Imagen
            img_buf = BytesIO()
            plt.savefig(img_buf, format='png', bbox_inches='tight')
            img_buf.seek(0)
            doc.add_picture(img_buf, width=Inches(5.5))

            # Causas y Consejos
            doc.add_heading('Análisis y Recomendaciones', level=1)
            doc.add_paragraph("Posibles Causas: Déficit de aprendizaje social, ansiedad inhibitoria o falta de reforzamiento ambiental.")
            doc.add_paragraph("Consejos: Iniciar entrenamiento en asertividad y técnicas de relajación social.")

            # Protocolo
            doc.add_heading('Protocolo de Respuestas', level=1)
            for i, p in enumerate(PREGUNTAS, 1):
                doc.add_paragraph(f"{i}. {p} -> {respuestas[i]}", style='List Bullet')

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 Descargar Informe Completo", w_buf.getvalue(), f"Informe_{nombre}.docx")
else:
    st.info("Debe aceptar el consentimiento informado para visualizar y realizar la prueba.")

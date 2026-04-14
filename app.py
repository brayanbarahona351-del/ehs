import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="EHS Goldstein - Sistema Clínico Pro", layout="centered")

if 'revision' not in st.session_state:
    st.session_state.revision = 0

def reset_form():
    st.session_state.revision += 1
    st.rerun()

# --- DATOS TÉCNICOS ---
GRUPOS = {
    "I: Primeras HHSS": {"items": list(range(1, 9)), "desc": "Habilidades básicas de comunicación: escuchar, iniciar y mantener conversaciones, así como dar las gracias."},
    "II: HHSS Avanzadas": {"items": list(range(9, 15)), "desc": "Capacidad para pedir ayuda, integrarse a grupos, dar y seguir instrucciones complejas."},
    "III: HHSS de Sentimientos": {"items": list(range(15, 22)), "desc": "Habilidades para conocer y expresar sentimientos propios, comprender los ajenos y manejar el afecto."},
    "IV: Alt. a la Agresión": {"items": list(range(22, 31)), "desc": "Capacidad de autocontrol, defensa de derechos, negociación y resolución pacífica de conflictos."},
    "V: Frente al Estrés": {"items": list(range(31, 43)), "desc": "Capacidad para responder a quejas, fracasos, presiones de grupo y manejar situaciones de vergüenza."},
    "VI: Planificación": {"items": list(range(43, 51)), "desc": "Toma de decisiones, fijar objetivos de forma realista y organización eficaz del trabajo."}
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
st.title("🚀 Sistema de Diagnóstico y Plan Terapéutico EHS")

with st.sidebar:
    st.header("Identificación")
    nombre = st.text_input("Nombre Completo", key=f"n_{st.session_state.revision}")
    edad = st.number_input("Edad", 12, 99, 20, key=f"e_{st.session_state.revision}")
    if st.button("🗑️ Reiniciar Prueba"): reset_form()

# INSTRUCCIONES SIEMPRE VISIBLES
st.subheader("1. Consentimiento e Instrucciones")
consent = st.checkbox("Acepto el consentimiento informado.")

st.warning("""
**INSTRUCCIONES DE LLENADO:**
Lea cada pregunta y seleccione la frecuencia según su comportamiento habitual:
- **1:** Nunca o muy pocas veces.
- **2:** Pocas veces.
- **3:** A veces.
- **4:** Muchas veces.
- **5:** Siempre o casi siempre.
""")

if consent:
    # CUESTIONARIO EN UNA COLUMNA
    respuestas = {}
    for i, p in enumerate(PREGUNTAS, 1):
        respuestas[i] = st.radio(f"**{i}. {p}**", [1,2,3,4,5], horizontal=True, key=f"q_{i}_{st.session_state.revision}", index=None)

    if st.button("📈 GENERAR DIAGNÓSTICO Y PLAN TERAPÉUTICO"):
        if None in respuestas.values():
            st.error("⚠️ Error: Debe completar todas las preguntas.")
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
            
            # Gráfico
            df = pd.DataFrame(res_data)
            fig, ax = plt.subplots(figsize=(10, 4))
            colors = ['#ff4b4b' if e <= 3 else '#00cc96' if e >= 7 else '#ffa500' for e in df['Eneatipo']]
            ax.bar(df['Área'], df['Eneatipo'], color=colors, edgecolor='black')
            ax.set_ylim(0, 10)
            plt.xticks(rotation=15)
            st.pyplot(fig)

            # PLAN TERAPÉUTICO AUTOMÁTICO
            st.header("📋 Plan Terapéutico Sugerido")
            deficit_areas = [r['Área'] for r in res_data if r['Eneatipo'] <= 3]
            
            if not deficit_areas:
                st.success("No se observan déficits significativos. Se recomienda mantenimiento de habilidades actuales.")
            else:
                st.write("Basado en el perfil, se recomienda intervención en:")
                for area in deficit_areas:
                    st.write(f"- **{area}**")
                
                st.info("""
                **Estrategias Generales:**
                1. **Modelado:** Observar a personas competentes en estas áreas.
                2. **Ensayo Conductual:** Practicar las conductas en sesión (Role-playing).
                3. **Retroalimentación:** Revisar y corregir el desempeño.
                4. **Tareas en casa:** Aplicar lo aprendido en entornos reales controlados.
                """)

            # --- GENERACIÓN DE WORD ---
            doc = Document()
            sec = doc.sections[0]
            sec.top_margin = sec.bottom_margin = Inches(0.5)
            
            doc.add_heading('INFORME PSICOLÓGICO Y PLAN TERAPÉUTICO', 0)
            doc.add_paragraph(f"Evaluado: {nombre} | Edad: {edad} años\nEneatipo Global: {enea_total}").bold = True

            # Protocolo
            doc.add_heading('1. Protocolo de Respuestas', level=1)
            for i, p in enumerate(PREGUNTAS, 1):
                p_doc = doc.add_paragraph(f"{i}. {p}: [{respuestas[i]}]")
                p_doc.runs[0].font.size = Pt(9)

            # Gráfico
            img_b = BytesIO()
            plt.savefig(img_b, format='png', bbox_inches='tight')
            doc.add_heading('2. Perfil de Competencia Social', level=1)
            doc.add_picture(img_b, width=Inches(5.5))

            # Análisis Extenso
            doc.add_heading('3. Análisis por Áreas', level=1)
            for r in res_data:
                area_p = doc.add_paragraph()
                area_p.add_run(f"{r['Área']} (Eneatipo {r['Eneatipo']}): ").bold = True
                area_p.add_run(r['Descripción'])

            # Plan Terapéutico
            doc.add_heading('4. Plan Terapéutico Estructurado', level=1)
            doc.add_paragraph(f"Objetivo: Fortalecer las áreas de {', '.join(deficit_areas)}.")
            doc.add_paragraph("Fases de Intervención:\n1. Reestructuración cognitiva sobre creencias sociales.\n2. Entrenamiento en habilidades específicas mediante ensayo.\n3. Generalización a entornos naturales.")

            w_buf = BytesIO()
            doc.save(w_buf)
            st.download_button("📥 DESCARGAR INFORME Y PLAN COMPLETO", w_buf.getvalue(), f"Informe_EHS_{nombre}.docx")

import streamlit as st

# Set page configuration as the first Streamlit command
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Import remaining dependencies
import pandas as pd
import base64
import io
import datetime
import logging
import uuid

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Custom CSS
st.markdown("""
    <style>
        .header { font-size: 2.5em; font-weight: bold; color: #2c3e50; text-align: center; margin-bottom: 20px; }
        .subheader { font-size: 1.8em; font-weight: bold; color: #34495e; margin-top: 20px; }
        .success { color: #27ae60; font-weight: bold; margin-top: 10px; }
        .stButton>button { width: 100%; border-radius: 5px; padding: 10px; }
        .stRadio>div { flex-direction: row; gap: 10px; }
        .dashboard-box { background-color: #f8f9fa; padding: 15px; border-radius: 10px; margin-bottom: 20px; }
        .error { color: #e74c3c; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# Cache questions data
@st.cache_data
def load_questions():
    return {
        "Human-Centered Employee Experience": {
            "English": [
                "Do employees feel their work contributes meaningfully to organizational success and personal growth?",
                "Are leaders evaluated on both business results and creating positive work environments?",
                "Are process improvements balanced with employee well-being?",
                "Do employees at all levels have input in decisions affecting their work?",
                "Is psychological safety actively cultivated?"
            ],
            "Español": [
                "¿Sienten los empleados que su trabajo contribuye al éxito organizacional y su crecimiento personal?",
                "¿Se evalúa a los líderes por resultados comerciales y crear entornos positivos?",
                "¿Se equilibran las mejoras de procesos con el bienestar del empleado?",
                "¿Tienen los empleados de todos niveles participación en decisiones que afectan su trabajo?",
                "¿Se cultiva activamente la seguridad psicológica?"
            ]
        },
        "Ethical Process Improvement": {
            "English": [
                "Does the organization consider human impacts when eliminating waste?",
                "Are process changes evaluated for negative consequences before implementation?",
                "Do continuous improvement initiatives include ethical considerations?",
                "Is retraining provided for employees affected by automation?",
                "Are metrics balanced between efficiency and human impact?"
            ],
            "Español": [
                "¿Considera la organización impactos humanos al eliminar desperdicios?",
                "¿Se evalúan los cambios en procesos por consecuencias negativas antes de implementar?",
                "¿Incluyen las iniciativas de mejora continua consideraciones éticas?",
                "¿Se proporciona recapacitación para empleados afectados por automatización?",
                "¿Se equilibran las métricas entre eficiencia e impacto humano?"
            ]
        },
        "Value Creation for All Stakeholders": {
            "English": [
                "Is success measured by value created for all stakeholders?",
                "Are customer needs balanced with employee capabilities in process design?",
                "Are long-term human and social impacts considered when cutting costs?",
                "Do process improvements benefit both employees and customers?",
                "Is community/societal impact considered in operational decisions?"
            ],
            "Español": [
                "¿Se mide el éxito por el valor creado para todos los interesados?",
                "¿Se equilibran las necesidades del cliente con las capacidades de empleados en el diseño de procesos?",
                "¿Se consideran impactos humanos y sociales a largo plazo al reducir costos?",
                "¿Benefician las mejoras de procesos a empleados y clientes?",
                "¿Se considera el impacto comunitario/social en decisiones operativas?"
            ]
        }
    }

# Initialize session state
def init_session_state():
    if 'initialized' not in st.session_state:
        questions = load_questions()
        st.session_state.update({
            'responses': {cat: [0]*len(questions[cat]["English"]) for cat in questions},
            'current_category': 0,
            'completed_categories': set(),
            'lang': 'English',
            'audit_timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'initialized': True
        })

init_session_state()

# Sidebar: Language selection
st.session_state.lang = st.sidebar.selectbox(
    "Language / Idioma",
    ["English", "Español"],
    key="lang_select"
)

# Sidebar: Reset session
if st.sidebar.button("Reset Session" if st.session_state.lang == "English" else "Restablecer Sesión", key="reset_session"):
    st.session_state.clear()
    init_session_state()
    st.rerun()

# Main title
st.markdown(
    f'<div class="header">{"Ethical Lean Audit" if st.session_state.lang == "English" else "Auditoría Lean Ética"}</div>',
    unsafe_allow_html=True
)

# Dashboard
questions = load_questions()
categories = list(questions.keys())
completed_count = len(st.session_state.completed_categories)
total_categories = len(categories)
st.markdown('<div class="dashboard-box">', unsafe_allow_html=True)
st.subheader("Progress Overview" if st.session_state.lang == "English" else "Resumen de Progreso")
st.write(
    f"{'Completed:' if st.session_state.lang == 'English' else 'Completado:'} {completed_count}/{total_categories} "
    f"({(completed_count/total_categories)*100:.1f}%)"
)
if st.session_state.completed_categories:
    st.write(
        f"{'Completed Categories:' if st.session_state.lang == 'English' else 'Categorías Completadas:'} "
        f"{', '.join(st.session_state.completed_categories)}"
    )
st.markdown('</div>', unsafe_allow_html=True)

# Progress bar
progress = (st.session_state.current_category / max(1, len(categories))) * 100
st.progress(min(int(progress), 100))

# Category navigation
st.sidebar.subheader("Progress" if st.session_state.lang == "English" else "Progreso")
category_index = st.sidebar.slider(
    "Select Category / Seleccionar Categoría",
    0, len(categories)-1,
    st.session_state.current_category,
    key="category_slider"
)
st.session_state.current_category = category_index
category = categories[category_index]

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Collect responses
st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
current_scores = st.session_state.responses.get(category, [0] * len(questions[category][st.session_state.lang]))
for idx, q in enumerate(questions[category][st.session_state.lang]):
    score = st.radio(
        q,
        options=list(range(1, 6)),
        format_func=lambda x: f"{x} - {labels[st.session_state.lang][x-1]}",
        key=f"{category}_{idx}",
        horizontal=True,
        index=None if current_scores[idx] == 0 else current_scores[idx]-1
    )
    st.session_state.responses[category][idx] = score if score is not None else 0

# Calculate and display category score
scores = st.session_state.responses.get(category, [])
score_sum = sum(scores)
max_score = len(questions[category][st.session_state.lang]) * 5
score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
st.write(
    f"{'Current Category Score:' if st.session_state.lang == 'English' else 'Puntuación Actual:'} "
    f"{score_sum}/{max_score} ({score_percent:.1f}%)"
)

# Mark category as completed
if all(1 <= score <= 5 for score in scores):
    st.session_state.completed_categories.add(category)
    st.markdown(
        f'<div class="success">{"Category completed!" if st.session_state.lang == "English" else "¡Categoría completada!"}</div>',
        unsafe_allow_html=True
    )

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button(
        "Previous Category" if st.session_state.lang == "English" else "Categoría Anterior",
        disabled=category_index == 0,
        key="prev_button"
    ):
        st.session_state.current_category = max(0, category_index - 1)
        st.rerun()
with col2:
    if st.button(
        "Next Category" if st.session_state.lang == "English" else "Siguiente Categoría",
        disabled=category_index == len(categories)-1,
        key="next_button"
    ):
        st.session_state.current_category = min(len(categories)-1, category_index + 1)
        st.rerun()

# Generate report
if st.button("Generate Report" if st.session_state.lang == "English" else "Generar Informe", key="generate_report"):
    incomplete_categories = [
        cat for cat in categories
        if not all(1 <= score <= 5 for score in st.session_state.responses.get(cat, []))
    ]
    if incomplete_categories:
        st.error(
            f"{'Please complete all questions for:' if st.session_state.lang == 'English' else 'Por favor complete todas las preguntas para:'} "
            f"{', '.join(incomplete_categories)}"
        )
    else:
        # Prepare results data
        results = []
        detailed_results = []
        for cat in categories:
            scores = st.session_state.responses.get(cat, [])
            if not scores:
                continue
            total = sum(scores)
            percent = (total / (len(scores)*5)) * 100 if scores else 0
            results.append({
                "Category": cat,
                "Score": total,
                "Percent": percent
            })
            for idx, (score, question) in enumerate(zip(scores, questions[cat][st.session_state.lang])):
                detailed_results.append({
                    "Category": cat,
                    "Question": question,
                    "Score": score,
                    "Rating": labels[st.session_state.lang][score-1]
                })
        
        df = pd.DataFrame(results)
        df_detailed = pd.DataFrame(detailed_results)
        
        # Display results
        st.subheader("Audit Results" if st.session_state.lang == "English" else "Resultados de la Auditoría")
        st.dataframe(df.style.format({"Score": "{:.0f}", "Percent": "{:.1f}%"}))
        st.subheader("Detailed Results" if st.session_state.lang == "English" else "Resultados Detallados")
        st.dataframe(df_detailed)
        
        # Export to CSV
        csv_buffer = io.StringIO()
        csv_buffer.write(f"Ethical Lean Audit Report\n")
        csv_buffer.write(f"Date: {st.session_state.audit_timestamp}\n\n")
        csv_buffer.write("Summary\n")
        df.to_csv(csv_buffer, index=False)
        csv_buffer.write("\nDetailed Results\n")
        df_detailed.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        b64_csv = base64.b64encode(csv_data.encode()).decode()
        csv_href = (
            f'<a href="data:text/csv;base64,{b64_csv}" '
            f'download="ethical_lean_audit_results.csv">'
            f'{"Download CSV Report" if st.session_state.lang == "English" else "Descargar Informe CSV"}</a>'
        )
        st.markdown(csv_href, unsafe_allow_html=True)

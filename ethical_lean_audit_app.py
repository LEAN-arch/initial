import streamlit as st

# Set page configuration as the first Streamlit command
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Import minimal dependencies
import pandas as pd
import base64
import io
import datetime
import sys

# Check if script is run correctly
if not sys.argv[0].endswith('streamlit-script-runner.py'):
    st.error("Please run this script using 'streamlit run ethical_lean_audit_minimal.py'.")
    sys.exit(1)

# Custom CSS
st.markdown("""
    <style>
        .header { font-size: 2em; font-weight: bold; color: #2c3e50; text-align: center; margin-bottom: 20px; }
        .subheader { font-size: 1.5em; font-weight: bold; color: #34495e; margin-top: 20px; }
        .success { color: #27ae60; font-weight: bold; margin-top: 10px; }
        .error { color: #e74c3c; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# Cache questions data
@st.cache_data
def load_questions():
    return {
        "Employee Experience": {
            "English": [
                "Does work contribute to organizational success and personal growth?",
                "Are leaders evaluated on creating positive work environments?"
            ],
            "Español": [
                "¿Contribuye el trabajo al éxito organizacional y crecimiento personal?",
                "¿Se evalúa a los líderes por crear entornos positivos?"
            ]
        },
        "Process Improvement": {
            "English": [
                "Are human impacts considered when eliminating waste?",
                "Are process changes evaluated for negative consequences?"
            ],
            "Español": [
                "¿Se consideran impactos humanos al eliminar desperdicios?",
                "¿Se evalúan los cambios por consecuencias negativas?"
            ]
        }
    }

# Initialize session state
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

# Sidebar: Language selection
st.session_state.lang = st.sidebar.selectbox("Language / Idioma", ["English", "Español"])

# Sidebar: Reset session
if st.sidebar.button("Reset" if st.session_state.lang == "English" else "Restablecer"):
    st.session_state.clear()
    st.rerun()

# Main title
st.markdown(f'<div class="header">{"Ethical Lean Audit" if st.session_state.lang == "English" else "Auditoría Lean Ética"}</div>', unsafe_allow_html=True)

# Category navigation
questions = load_questions()
categories = list(questions.keys())
category_index = st.sidebar.slider(
    "Category / Categoría",
    0, len(categories)-1,
    st.session_state.current_category
)
st.session_state.current_category = category_index
category = categories[category_index]

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely", "Partially", "Mostly", "Fully"],
    "Español": ["No practicado", "Raramente", "Parcialmente", "Mayormente", "Totalmente"]
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

# Display category score
scores = st.session_state.responses.get(category, [])
score_sum = sum(scores)
max_score = len(questions[category][st.session_state.lang]) * 5
score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
st.write(f"{'Score:' if st.session_state.lang == 'English' else 'Puntuación:'} {score_sum}/{max_score} ({score_percent:.1f}%)")

# Mark category as completed
if all(1 <= score <= 5 for score in scores):
    st.session_state.completed_categories.add(category)
    st.markdown(f'<div class="success">{"Category completed!" if st.session_state.lang == "English" else "¡Categoría completada!"}</div>', unsafe_allow_html=True)

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Previous" if st.session_state.lang == "English" else "Anterior", disabled=category_index == 0):
        st.session_state.current_category = max(0, category_index - 1)
        st.rerun()
with col2:
    if st.button("Next" if st.session_state.lang == "English" else "Siguiente", disabled=category_index == len(categories)-1):
        st.session_state.current_category = min(len(categories)-1, category_index + 1)
        st.rerun()

# Generate report
if st.button("Generate Report" if st.session_state.lang == "English" else "Generar Informe"):
    incomplete_categories = [cat for cat in categories if not all(1 <= score <= 5 for score in st.session_state.responses.get(cat, []))]
    if incomplete_categories:
        st.markdown(f'<div class="error">{"Complete all questions for: " if st.session_state.lang == "English" else "Complete todas las preguntas para: "}{", ".join(incomplete_categories)}</div>', unsafe_allow_html=True)
    else:
        results = []
        for cat in categories:
            scores = st.session_state.responses.get(cat, [])
            total = sum(scores)
            percent = (total / (len(scores)*5)) * 100 if scores else 0
            results.append({"Category": cat, "Score": total, "Percent": percent})
        
        df = pd.DataFrame(results)
        csv_buffer = io.StringIO()
        csv_buffer.write(f"Ethical Lean Audit Report\nDate: {st.session_state.audit_timestamp}\n\n")
        df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        b64_csv = base64.b64encode(csv_data.encode()).decode()
        csv_href = f'<a href="data:text/csv;base64,{b64_csv}" download="audit_results.csv">{"Download CSV" if st.session_state.lang == "English" else "Descargar CSV"}</a>'
        st.markdown(csv_href, unsafe_allow_html=True)

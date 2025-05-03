import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import base64
import io
import numpy as np
import seaborn as sns

# Custom CSS for engagement and visual appeal
st.markdown("""
    <style>
        /* Your existing CSS styles */
    </style>
""", unsafe_allow_html=True)

# Set page configuration
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Theme selection
theme = st.sidebar.selectbox("Theme / Tema", ["Light", "Dark"], help="Choose your preferred theme / Elige tu tema preferido")
if theme == "Dark":
    st.markdown('<style>body, .main { background-color: #1e1e1e; color: #e0e0e0; } .main { background-color: #2e2e2e; }</style>', unsafe_allow_html=True)

# Bilingual support
LANG = st.sidebar.selectbox("Language / Idioma", ["English", "Español"], help="Select your preferred language / Selecciona tu idioma preferido", key="lang_select")

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Enhanced Audit categories and questions
questions = {
    "Human-Centered Employee Experience": {
        "English": [
            "Do employees feel their work contributes meaningfully to both organizational success and personal growth?",
            "Are leaders evaluated on both business results AND their ability to create positive work environments?",
            "When process improvements are made, is employee well-being given equal weight to efficiency gains?",
            "Do employees at all levels have genuine input in decisions that affect their daily work?",
            "Is psychological safety actively cultivated?"
        ],
        "Español": [
            "¿Sienten los empleados que su trabajo contribuye significativamente al éxito organizacional y su crecimiento personal?",
            "¿Se evalúa a los líderes tanto por resultados comerciales COMO por crear entornos de trabajo positivos?",
            "En mejoras de procesos, ¿se considera igual el bienestar del empleado que las ganancias de eficiencia?",
            "¿Tienen empleados de todos niveles participación genuina en decisiones que afectan su trabajo diario?",
            "¿Se cultiva activamente la seguridad psicológica?"
        ]
    },
    "Ethical Process Improvement": {
        "English": [
            "When eliminating waste, does the organization consider human impacts alongside productivity?",
            "Are process changes evaluated for potential negative consequences before implementation?",
            "Do continuous improvement initiatives explicitly include ethical considerations?",
            "When automating, is equal attention given to retraining affected employees?",
            "Are metrics balanced between efficiency and human impact?"
        ],
        "Español": [
            "Al eliminar desperdicios, ¿se consideran impactos humanos junto con productividad?",
            "¿Se evalúan cambios en procesos por posibles consecuencias negativas antes de implementar?",
            "¿Las iniciativas de mejora continua incluyen explícitamente consideraciones éticas?",
            "Al automatizar, ¿se presta igual atención a recapacitar empleados afectados?",
            "¿Las métricas equilibran eficiencia e impacto humano?"
        ]
    },
    "Value Creation for All Stakeholders": {
        "English": [
            "Does your organization measure success by value created for ALL stakeholders?",
            "Are customer needs balanced with employee capabilities when designing processes?",
            "When cutting costs, does leadership consider long-term human and social impacts?",
            "Do process improvements create mutual benefit for employees and customers?",
            "Is community/societal impact considered in operational decisions?"
        ],
        "Español": [
            "¿Su organización mide el éxito por valor creado para TODOS los interesados?",
            "¿Necesidades del cliente se equilibran con capacidades de empleados al diseñar procesos?",
            "Al reducir costos, ¿se consideran impactos humanos y sociales a largo plazo?",
            "¿Las mejoras de procesos crean beneficio mutuo para empleados y clientes?",
            "¿Se considera impacto comunitario/social en decisiones operativas?"
        ]
    }
}

# Initialize session state
if 'responses' not in st.session_state:
    st.session_state.responses = {cat: [0]*len(questions[cat][LANG]) for cat in questions}
if 'current_category' not in st.session_state:
    st.session_state.current_category = 0
if 'completed_categories' not in st.session_state:
    st.session_state.completed_categories = set()

# Reset session state button
if st.sidebar.button("Reset Session" if LANG == "English" else "Restablecer Sesión", key="reset_session"):
    st.session_state.clear()
    st.rerun()

# Main title
st.markdown('<div class="header">Ethical Lean Audit</div>' if LANG == "English" else '<div class="header">Auditoría Lean Ética</div>', unsafe_allow_html=True)

# Progress tracking
categories = list(questions.keys())
progress = (st.session_state.current_category / len(categories)) * 100
st.progress(int(progress))

# Category navigation
st.sidebar.subheader("Progress" if LANG == "English" else "Progreso")
category_index = st.sidebar.slider(
    "Select Category / Seleccionar Categoría",
    0, len(categories)-1, st.session_state.current_category,
    key="category_slider"
)
st.session_state.current_category = category_index
category = categories[category_index]

# Collect responses
st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
current_scores = st.session_state.responses[category]

for idx, q in enumerate(questions[category][LANG]):
    score = st.radio(
        q,
        options=list(range(1,6)),
        format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
        key=f"{category}_{idx}",
        horizontal=True
    )
    st.session_state.responses[category][idx] = int(score)  # Ensure numeric

# Calculate and display category score
score_sum = sum(st.session_state.responses[category])
max_score = len(questions[category][LANG]) * 5
score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
st.write(f"{'Current Category Score:' if LANG == 'English' else 'Puntuación Actual:'} {score_sum}/{max_score} ({score_percent:.1f}%)")

# Mark category as completed
if all(1 <= score <= 5 for score in st.session_state.responses[category]):
    st.session_state.completed_categories.add(category)
    st.markdown(f'<div class="success">{"Category completed!" if LANG == "English" else "¡Categoría completada!"}</div>', unsafe_allow_html=True)

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Previous Category" if LANG == "English" else "Categoría Anterior", disabled=category_index == 0):
        st.session_state.current_category -= 1
        st.rerun()
with col2:
    if st.button("Next Category" if LANG == "English" else "Siguiente Categoría", disabled=category_index == len(categories)-1):
        st.session_state.current_category += 1
        st.rerun()

# Generate report
if st.button("Generate Report" if LANG == "English" else "Generar Informe", key="generate_report"):
    # Validate responses
    incomplete_categories = [cat for cat in categories if not all(1 <= score <= 5 for score in st.session_state.responses[cat])]
    if incomplete_categories:
        st.error(f"{'Please complete all questions for:' if LANG == 'English' else 'Por favor complete todas las preguntas para:'} {', '.join(incomplete_categories)}")
        st.stop()

    # Prepare results data
    results = []
    for cat in categories:
        scores = st.session_state.responses[cat]
        total = sum(scores)
        percent = (total / (len(scores)*5)) * 100
        results.append({
            "Category": cat,
            "Score": total,
            "Percent": percent
        })
    
    df = pd.DataFrame(results)
    
    # Display results
    st.dataframe(df.style.format({"Score": "{:.0f}", "Percent": "{:.1f}%"}))
    
    # Visualization
    st.subheader("Audit Results Visualization" if LANG == "English" else "Visualización de Resultados")
    fig, ax = plt.subplots(figsize=(10,6))
    sns.barplot(data=df, x="Percent", y="Category", palette="coolwarm")
    plt.title("Audit Results" if LANG == "English" else "Resultados de la Auditoría")
    plt.xlabel("Percentage" if LANG == "English" else "Porcentaje")
    plt.ylabel("Category" if LANG == "English" else "Categoría")
    st.pyplot(fig)

    # PDF Report
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, "Ethical Lean Audit Report" if LANG == "English" else "Informe de Auditoría Lean Ética", 0, 1, "C")
        pdf.ln(10)
        
        pdf.set_font("Arial", "", 12)
        for _, row in df.iterrows():
            pdf.cell(0, 10, f"{row['Category']}: {row['Percent']:.1f}%", 0, 1)
        
        pdf_output = io.BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)
        b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
        href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_lean_audit.pdf">{"Download PDF Report" if LANG == "English" else "Descargar Informe PDF"}</a>'
        st.markdown(href, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error generating PDF: {str(e)}")

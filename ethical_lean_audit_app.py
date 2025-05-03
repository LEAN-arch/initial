import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf2 import FPDF
import base64
import io
import numpy as np
import seaborn as sns
import datetime

# Custom CSS for enhanced visual appeal
st.markdown("""
    <style>
        .header { font-size: 2.5em; font-weight: bold; color: #2c3e50; text-align: center; margin-bottom: 20px; }
        .subheader { font-size: 1.8em; font-weight: bold; color: #34495e; margin-top: 20px; }
        .success { color: #27ae60; font-weight: bold; margin-top: 10px; }
        .stButton>button { width: 100%; border-radius: 5px; padding: 10px; }
        .stRadio>div { flex-direction: row; gap: 10px; }
        .dashboard-box { background-color: #f8f9fa; padding: 15px; border-radius: 10px; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# Set page configuration
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Initialize session state with proper defaults
def init_session_state():
    defaults = {
        'responses': {cat: [0]*len(questions[cat]["English"]) for cat in questions},
        'current_category': 0,
        'completed_categories': set(),
        'lang': 'English',
        'audit_timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

init_session_state()

# Theme selection
theme = st.sidebar.selectbox(
    "Theme / Tema",
    ["Light", "Dark"],
    help="Choose your preferred theme / Elige tu tema preferido",
    key="theme_select"
)
if theme == "Dark":
    st.markdown(
        '<style>body, .main { background-color: #1e1e1e; color: #e0e0e0; } .main { background-color: #2e2e2e; }</style>',
        unsafe_allow_html=True
    )

# Bilingual support
LANG = st.sidebar.selectbox(
    "Language / Idioma",
    ["English", "Español"],
    help="Select your preferred language / Selecciona tu idioma preferido",
    key="lang_select"
)
st.session_state.lang = LANG

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Audit categories and questions
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

# Reset session state button
if st.sidebar.button("Reset Session" if LANG == "English" else "Restablecer Sesión", key="reset_session"):
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    init_session_state()
    st.rerun()

# Main title
st.markdown(
    '<div class="header">Ethical Lean Audit</div>' if LANG == "English" else
    '<div class="header">Auditoría Lean Ética</div>',
    unsafe_allow_html=True
)

# Dashboard for completed categories
st.markdown('<div class="dashboard-box">', unsafe_allow_html=True)
st.subheader("Progress Overview" if LANG == "English" else "Resumen de Progreso")
completed_count = len(st.session_state.completed_categories)
total_categories = len(questions)
st.write(
    f"{'Completed:' if LANG == 'English' else 'Completado:'} {completed_count}/{total_categories} "
    f"({(completed_count/total_categories)*100:.1f}%)"
)
if st.session_state.completed_categories:
    st.write(
        f"{'Completed Categories:' if LANG == 'English' else 'Categorías Completadas:'} "
        f"{', '.join(st.session_state.completed_categories)}"
    )
st.markdown('</div>', unsafe_allow_html=True)

# Progress tracking
categories = list(questions.keys())
progress = min((st.session_state.current_category / len(categories)) * 100, 100.0)
st.progress(int(progress))

# Category navigation
st.sidebar.subheader("Progress" if LANG == "English" else "Progreso")
category_index = st.sidebar.slider(
    "Select Category / Seleccionar Categoría",
    0, len(categories)-1,
    st.session_state.current_category,
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
        options=list(range(1, 6)),
        format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
        key=f"{category}_{idx}",
        horizontal=True,
        index=None if current_scores[idx] == 0 else current_scores[idx]-1
    )
    st.session_state.responses[category][idx] = int(score or 0)

# Calculate and display category score
score_sum = sum(st.session_state.responses[category])
max_score = len(questions[category][LANG]) * 5
score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
st.write(
    f"{'Current Category Score:' if LANG == 'English' else 'Puntuación Actual:'} "
    f"{score_sum}/{max_score} ({score_percent:.1f}%)"
)

# Mark category as completed
if all(1 <= score <= 5 for score in st.session_state.responses[category]):
    st.session_state.completed_categories.add(category)
    st.markdown(
        f'<div class="success">{"Category completed!" if LANG == "English" else "¡Categoría completada!"}</div>',
        unsafe_allow_html=True
    )

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button(
        "Previous Category" if LANG == "English" else "Categoría Anterior",
        disabled=category_index == 0,
        key="prev_button"
    ):
        st.session_state.current_category -= 1
        st.rerun()
with col2:
    if st.button(
        "Next Category" if LANG == "English" else "Siguiente Categoría",
        disabled=category_index == len(categories)-1,
        key="next_button"
    ):
        st.session_state.current_category += 1
        st.rerun()

# Generate report
if st.button("Generate Report" if LANG == "English" else "Generar Informe", key="generate_report"):
    # Validate responses
    incomplete_categories = [
        cat for cat in categories
        if not all(1 <= score <= 5 for score in st.session_state.responses[cat])
    ]
    if incomplete_categories:
        st.error(
            f"{'Please complete all questions for:' if LANG == 'English' else 'Por favor complete todas las preguntas para:'} "
            f"{', '.join(incomplete_categories)}"
        )
    else:
        # Prepare results data
        results = []
        detailed_results = []
        for cat in categories:
            scores = st.session_state.responses[cat]
            total = sum(scores)
            percent = (total / (len(scores)*5)) * 100
            results.append({
                "Category": cat,
                "Score": total,
                "Percent": percent
            })
            for idx, (score, question) in enumerate(zip(scores, questions[cat][LANG])):
                detailed_results.append({
                    "Category": cat,
                    "Question": question,
                    "Score": score,
                    "Rating": labels[LANG][score-1]
                })
        
        df = pd.DataFrame(results)
        df_detailed = pd.DataFrame(detailed_results)
        
        # Display results
        st.subheader("Audit Results" if LANG == "English" else "Resultados de la Auditoría")
        st.dataframe(df.style.format({"Score": "{:.0f}", "Percent": "{:.1f}%"}))
        
        # Interactive Visualization
        try:
            fig = px.bar(
                df,
                x="Percent",
                y="Category",
                color="Percent",
                color_continuous_scale="RdYlGn",
                title="Audit Results" if LANG == "English" else "Resultados de la Auditoría",
                labels={"Percent": "Percentage" if LANG == "English" else "Porcentaje", "Category": "Category" if LANG == "English" else "Categoría"}
            )
            fig.update_layout(
                showlegend=False,
                height=400,
                margin=dict(l=20, r=20, t=50, b=20)
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Detailed radar chart
            radar_data = []
            for cat in categories:
                scores = st.session_state.responses[cat]
                radar_data.append(
                    go.Scatterpolar(
                        r=scores,
                        theta=[q[:30] + "..." for q in questions[cat][LANG]],
                        fill='toself',
                        name=cat
                    )
                )
            radar_fig = go.Figure(
                data=radar_data,
                layout=go.Layout(
                    polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
                    showlegend=True,
                    title="Detailed Category Scores" if LANG == "English" else "Puntuaciones Detalladas por Categoría"
                )
            )
            st.plotly_chart(radar_fig, use_container_width=True)
        except Exception as e:
            st.error(f"Error generating visualizations: {str(e)}")
        
        # Export to CSV
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        b64_csv = base64.b64encode(csv_data.encode()).decode()
        csv_href = (
            f'<a href="data:text/csv;base64,{b64_csv}" '
            f'download="ethical_lean_audit_results.csv">'
            f'{"Download CSV Report" if LANG == "English" else "Descargar Informe CSV"}</a>'
        )
        st.markdown(csv_href, unsafe_allow_html=True)
        
        # Enhanced PDF Report
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(
                0, 10,
                "Ethical Lean Audit Report" if LANG == "English" else "Informe de Auditoría Lean Ética",
                0, 1, "C"
            )
            pdf.set_font("Arial", "", 12)
            pdf.cell(0, 10, f"Date: {st.session_state.audit_timestamp}", 0, 1)
            pdf.ln(10)
            
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Summary" if LANG == "English" else "Resumen", 0, 1)
            pdf.set_font("Arial", "", 12)
            for _, row in df.iterrows():
                pdf.cell(0, 8, f"{row['Category']}: {row['Score']}/{len(questions[row['Category']][LANG])*5} ({row['Percent']:.1f}%)", 0, 1)
            
            pdf.ln(10)
            pdf.set_font("Arial", "B", 14)
            pdf.cell(0, 10, "Detailed Results" if LANG == "English" else "Resultados Detallados", 0, 1)
            pdf.set_font("Arial", "", 12)
            for _, row in df_detailed.iterrows():
                pdf.multi_cell(0, 8, f"[{row['Category']}] {row['Question']}: {row['Rating']} (Score: {row['Score']})")
                pdf.ln(2)
            
            pdf_output = io.BytesIO()
            pdf.output(pdf_output)
            pdf_output.seek(0)
            b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
            href = (
                f'<a href="data:application/pdf;base64,{b64_pdf}" '
                f'download="ethical_lean_audit.pdf">'
                f'{"Download PDF Report" if LANG == "English" else "Descargar Informe PDF"}</a>'
            )
            st.markdown(href, unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error generating PDF: {str(e)}")

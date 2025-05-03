import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
import base64
import io
import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
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

# Set page configuration
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Questions data (static, no caching needed)
QUESTIONS = {
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

# Likert scale labels
LABELS = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Centralized error handler
def handle_error(message, error):
    logger.error(f"{message}: {str(error)}")
    st.error(f"{message}. Please try again.")

# Initialize session state
def init_session_state():
    defaults = {
        'responses': {cat: [0] * len(QUESTIONS[cat]["English"]) for cat in QUESTIONS},
        'current_category': 0,
        'completed_categories': set(),
        'lang': 'English',
        'audit_timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# Render sidebar
def render_sidebar(categories):
    st.sidebar.subheader("Progress" if st.session_state.lang == "English" else "Progreso")
    st.session_state.lang = st.sidebar.selectbox(
        "Language / Idioma",
        ["English", "Español"],
        key="lang_select"
    )
    st.session_state.current_category = st.sidebar.slider(
        "Select Category / Seleccionar Categoría",
        0, len(categories) - 1,
        st.session_state.current_category,
        key="category_slider"
    )
    if st.sidebar.button("Reset Session" if st.session_state.lang == "English" else "Restablecer Sesión", key="reset_session"):
        st.session_state.clear()
        init_session_state()
        st.rerun()

# Render dashboard
def render_dashboard(categories):
    st.markdown('<div class="dashboard-box">', unsafe_allow_html=True)
    st.subheader("Progress Overview" if st.session_state.lang == "English" else "Resumen de Progreso")
    completed_count = len(st.session_state.completed_categories)
    total_categories = len(categories)
    st.write(f"{'Completed:' if st.session_state.lang == 'English' else 'Completado:'} {completed_count}/{total_categories} "
             f"({(completed_count / total_categories) * 100:.1f}%)")
    if st.session_state.completed_categories:
        st.write(f"{'Completed Categories:' if st.session_state.lang == 'English' else 'Categorías Completadas:'} "
                 f"{', '.join(st.session_state.completed_categories)}")
    st.markdown('</div>', unsafe_allow_html=True)
    st.progress(int(min((st.session_state.current_category / len(categories)) * 100, 100.0)))

# Render survey
def render_survey(category, questions):
    st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
    current_scores = st.session_state.responses.get(category, [0] * len(questions[category][st.session_state.lang]))
    
    for idx, q in enumerate(questions[category][st.session_state.lang]):
        score = st.radio(
            q,
            options=list(range(1, 6)),
            format_func=lambda x: f"{x} - {LABELS[st.session_state.lang][x - 1]}",
            key=f"{category}_{idx}",
            horizontal=True,
            index=None if current_scores[idx] == 0 else current_scores[idx] - 1
        )
        st.session_state.responses[category][idx] = int(score or 0)
    
    score_sum = sum(st.session_state.responses[category])
    max_score = len(questions[category][st.session_state.lang]) * 5
    score_percent = (score_sum / max_score * 100) if max_score > 0 else 0  # Corrected line
    st.write(f"{'Current Category Score:' if st.session_state.lang == 'English' else 'Puntuación Actual:'} "
             f"{score_sum}/{max_score} ({score_percent:.1f}%)")
    
    if all(1 <= score <= 5 for score in st.session_state.responses[category]):
        st.session_state.completed_categories.add(category)
        st.markdown(f'<div class="success">{"Category completed!" if st.session_state.lang == "English" else "¡Categoría completada!"}</div>',
                    unsafe_allow_html=True)

# Render navigation
def render_navigation(categories):
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Previous Category" if st.session_state.lang == "English" else "Categoría Anterior",
                     disabled=st.session_state.current_category == 0, key="prev_button"):
            st.session_state.current_category -= 1
            st.rerun()
    with col2:
        if st.button("Next Category" if st.session_state.lang == "English" else "Siguiente Categoría",
                     disabled=st.session_state.current_category == len(categories) - 1, key="next_button"):
            st.session_state.current_category += 1
            st.rerun()

# Generate report
def generate_report(categories, questions):
    if st.button("Generate Report" if st.session_state.lang == "English" else "Generar Informe", key="generate_report"):
        incomplete_categories = [
            cat for cat in categories
            if not all(1 <= score <= 5 for score in st.session_state.responses.get(cat, []))
        ]
        if incomplete_categories:
            st.error(f"{'Please complete all questions for:' if st.session_state.lang == 'English' else 'Por favor complete todas las preguntas para:'} "
                     f"{', '.join(incomplete_categories)}")
            return
        
        results = []
        detailed_results = []
        for cat in categories:
            scores = st.session_state.responses.get(cat, [])
            if not scores:
                continue
            total = sum(scores)
            percent = (total / (len(scores) * 5)) * 100
            results.append({"Category": cat, "Score": total, "Percent": percent})
            for idx, (score, question) in enumerate(zip(scores, questions[cat][st.session_state.lang])):
                detailed_results.append({
                    "Category": cat,
                    "Question": question,
                    "Score": score,
                    "Rating": LABELS[st.session_state.lang][score - 1]
                })
        
        df = pd.DataFrame(results)
        df_detailed = pd.DataFrame(detailed_results)
        
        st.subheader("Audit Results" if st.session_state.lang == "English" else "Resultados de la Auditoría")
        st.dataframe(df.style.format({"Score": "{:.0f}", "Percent": "{:.1f}%"}))
        
        # Bar chart
        fig = px.bar(
            df,
            x="Percent",
            y="Category",
            color="Percent",
            color_continuous_scale="RdYlGn",
            title="Audit Results" if st.session_state.lang == "English" else "Resultados de la Auditoría",
            labels={"Percent": "Percentage" if st.session_state.lang == "English" else "Porcentaje",
                    "Category": "Category" if st.session_state.lang == "English" else "Categoría"}
        )
        fig.update_layout(showlegend=False, height=400, margin=dict(l=20, r=20, t=50, b=20))
        st.plotly_chart(fig, use_container_width=True)
        
        # Radar chart
        radar_data = []
        for cat in categories:
            scores = st.session_state.responses.get(cat, [])
            if not scores:
                continue
            radar_data.append(
                go.Scatterpolar(
                    r=scores,
                    theta=[q[:30] + "..." for q in questions[cat][st.session_state.lang]],
                    fill='toself',
                    name=cat
                )
            )
        radar_fig = go.Figure(
            data=radar_data,
            layout=go.Layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
                showlegend=True,
                title="Detailed Category Scores" if st.session_state.lang == "English" else "Puntuaciones Detalladas por Categoría"
            )
        )
        st.plotly_chart(radar_fig, use_container_width=True)
        
        # CSV export
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False)
        st.download_button(
            label="Download CSV Report" if st.session_state.lang == "English" else "Descargar Informe CSV",
            data=csv_buffer.getvalue(),
            file_name="ethical_lean_audit_results.csv",
            mime="text/csv"
        )
        
        # PDF export
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        
        elements.append(Paragraph(
            "Ethical Lean Audit Report" if st.session_state.lang == "English" else "Informe de Auditoría Lean Ética",
            styles['Title']
        ))
        elements.append(Paragraph(f"Date: {st.session_state.audit_timestamp}", styles['Normal']))
        
        elements.append(Paragraph("Summary" if st.session_state.lang == "English" else "Resumen", styles['Heading2']))
        summary_data = [["Category", "Score", "Percent"]] + [
            [row['Category'], f"{row['Score']}/{len(questions[row['Category']][st.session_state.lang]) * 5}", f"{row['Percent']:.1f}%"]
            for _, row in df.iterrows()
        ]
        summary_table = Table(summary_data)
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(summary_table)
        
        elements.append(Paragraph("Detailed Results" if st.session_state.lang == "English" else "Resultados Detallados", styles['Heading2']))
        for _, row in df_detailed.iterrows():
            elements.append(Paragraph(f"[{row['Category']}] {row['Question']}: {row['Rating']} (Score: {row['Score']})", styles['Normal']))
        
        doc.build(elements)
        st.download_button(
            label="Download PDF Report" if st.session_state.lang == "English" else "Descargar Informe PDF",
            data=pdf_buffer.getvalue(),
            file_name="ethical_lean_audit.pdf",
            mime="application/pdf"
        )

# Main function
def main():
    try:
        init_session_state()
        categories = list(QUESTIONS.keys())
        
        st.markdown('<div class="header">Ethical Lean Audit</div>' if st.session_state.lang == "English" else
                    '<div class="header">Auditoría Lean Ética</div>', unsafe_allow_html=True)
        
        render_sidebar(categories)
        render_dashboard(categories)
        render_survey(categories[st.session_state.current_category], QUESTIONS)
        render_navigation(categories)
        generate_report(categories, QUESTIONS)
        
    except Exception as e:
        handle_error("Error running the application", e)

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
import io
import numpy as np
import seaborn as sns
import datetime
import logging
import uuid

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
try:
    st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")
except Exception as e:
    logger.error(f"Error setting page config: {str(e)}")
    st.error("Failed to initialize the application. Please refresh the page.")
    st.stop()

# Cache questions data
@st.cache_data
def load_questions():
    return {
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
def init_session_state():
    try:
        questions = load_questions()
        if 'initialized' not in st.session_state:
            st.session_state.update({
                'responses': {cat: [0]*len(questions[cat]["English"]) for cat in questions},
                'current_category': 0,
                'completed_categories': set(),
                'lang': 'English',
                'audit_timestamp': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'initialized': True
            })
    except Exception as e:
        logger.error(f"Error initializing session state: {str(e)}")
        st.error("Failed to initialize session state. Please refresh the page.")
        st.stop()

init_session_state()

# Theme selection
try:
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
except Exception as e:
    logger.error(f"Error in theme selection: {str(e)}")
    st.error("Error loading theme selector.")

# Language selection
try:
    st.session_state.lang = st.sidebar.selectbox(
        "Language / Idioma",
        ["English", "Español"],
        help="Select your preferred language / Selecciona tu idioma preferido",
        key="lang_select"
    )
except Exception as e:
    logger.error(f"Error in language selection: {str(e)}")
    st.error("Error loading language selector.")

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Reset session state
try:
    if st.sidebar.button("Reset Session" if st.session_state.lang == "English" else "Restablecer Sesión", key="reset_session"):
        st.session_state.clear()
        init_session_state()
        st.rerun()
except Exception as e:
    logger.error(f"Error in reset session: {str(e)}")
    st.error("Error resetting session.")

# Main title
try:
    st.markdown(
        f'<div class="header">{"Ethical Lean Audit" if st.session_state.lang == "English" else "Auditoría Lean Ética"}</div>',
        unsafe_allow_html=True
    )
except Exception as e:
    logger.error(f"Error rendering main title: {str(e)}")
    st.error("Error rendering title.")

# Dashboard
try:
    st.markdown('<div class="dashboard-box">', unsafe_allow_html=True)
    st.subheader("Progress Overview" if st.session_state.lang == "English" else "Resumen de Progreso")
    questions = load_questions()
    completed_count = len(st.session_state.completed_categories)
    total_categories = len(questions)
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
except Exception as e:
    logger.error(f"Error rendering dashboard: {str(e)}")
    st.error("Error rendering progress dashboard.")

# Progress tracking
try:
    questions = load_questions()
    categories = list(questions.keys())
    progress = (st.session_state.current_category / max(1, len(categories))) * 100 if categories else 0
    st.progress(min(int(progress), 100))
except Exception as e:
    logger.error(f"Error rendering progress bar: {str(e)}")
    st.error("Error rendering progress bar.")

# Category navigation
try:
    st.sidebar.subheader("Progress" if st.session_state.lang == "English" else "Progreso")
    category_index = st.sidebar.slider(
        "Select Category / Seleccionar Categoría",
        0, len(categories)-1,
        st.session_state.current_category,
        key="category_slider"
    )
    st.session_state.current_category = category_index
    category = categories[category_index]
except Exception as e:
    logger.error(f"Error in category navigation: {str(e)}")
    st.error("Error loading category navigation.")

# Collect responses
try:
    st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
    questions = load_questions()
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
except Exception as e:
    logger.error(f"Error collecting responses: {str(e)}")
    st.error("Error loading survey questions.")

# Calculate and display category score
try:
    scores = st.session_state.responses.get(category, [])
    score_sum = sum(scores)
    max_score = len(questions[category][st.session_state.lang]) * 5
    score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
    st.write(
        f"{'Current Category Score:' if st.session_state.lang == 'English' else 'Puntuación Actual:'} "
        f"{score_sum}/{max_score} ({score_percent:.1f}%)"
    )
except Exception as e:
    logger.error(f"Error calculating category score: {str(e)}")
    st.error("Error calculating category score.")

# Mark category as completed
try:
    if all(1 <= score <= 5 for score in st.session_state.responses.get(category, [])):
        st.session_state.completed_categories.add(category)
        st.markdown(
            f'<div class="success">{"Category completed!" if st.session_state.lang == "English" else "¡Categoría completada!"}</div>',
            unsafe_allow_html=True
        )
except Exception as e:
    logger.error(f"Error marking category as completed: {str(e)}")
    st.error("Error updating category status.")

# Navigation buttons
try:
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
except Exception as e:
    logger.error(f"Error in navigation buttons: {str(e)}")
    st.error("Error loading navigation buttons.")

# Generate report
try:
    if st.button("Generate Report" if st.session_state.lang == "English" else "Generar Informe", key="generate_report"):
        questions = load_questions()
        categories = list(questions.keys())
        # Validate responses
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
            
            # Interactive Visualization
            try:
                fig = px.bar(
                    df,
                    x="Percent",
                    y="Category",
                    color="Percent",
                    color_continuous_scale="RdYlGn",
                    title="Audit Results" if st.session_state.lang == "English" else "Resultados de la Auditoría",
                    labels={"Percent": "Percentage" if st.session_state.lang == "English" else "Porcentaje", "Category": "Category" if st.session_state.lang == "English" else "Categoría"}
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
            except Exception as e:
                logger.error(f"Error generating visualizations: {str(e)}")
                st.error(f"Error generating visualizations: {str(e)}")
            
            # Export to CSV
            try:
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
                st.info("PDF export is not available in this environment. Please use the CSV report for detailed results.")
            except Exception as e:
                logger.error(f"Error generating CSV: {str(e)}")
                st.error(f"Error generating CSV: {str(e)}")
except Exception as e:
    logger.error(f"Error generating report: {str(e)}")
    st.error(f"Error generating report: {str(e)}")

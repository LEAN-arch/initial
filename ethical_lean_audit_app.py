import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import base64
import io
import numpy as np
import os

# Set page configuration as the first Streamlit command
st.set_page_config(page_title="Auditoría Ética de Lugar de Trabajo Lean", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for modern, accessible, and responsive UI
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #f5f7fa;
            color: #333;
        }
        .main-container {
            background-color: #ffffff;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
            max-width: 1000px;
            margin: 20px auto;
        }
        .stButton>button {
            background-color: #007bff;
            color: white;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 700;
            transition: background-color 0.3s ease, transform 0.2s ease;
        }
        .stButton>button:hover {
            background-color: #0056b3;
            transform: scale(1.05);
        }
        .stButton>button:focus {
            outline: 2px solid #007bff;
            outline-offset: 2px;
        }
        .stRadio>label {
            background-color: #e9f7ff;
            padding: 12px;
            border-radius: 8px;
            margin: 10px 0;
            font-size: 1.1em;
            transition: background-color 0.3s ease;
        }
        .stRadio>label:hover {
            background-color: #d0e9ff;
        }
        .header {
            color: #007bff;
            font-size: 2.5em;
            text-align: center;
            margin-bottom: 20px;
            font-weight: 700;
        }
        .subheader {
            color: #333;
            font-size: 1.6em;
            margin: 20px 0;
            text-align: center;
            font-weight: 400;
        }
        .sidebar .sidebar-content {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        .download-link {
            color: #007bff;
            font-weight: 700;
            text-decoration: none;
            font-size: 1.1em;
            display: inline-flex;
            align-items: center;
        }
        .download-link:hover {
            color: #0056b3;
            text-decoration: underline;
        }
        .download-link::before {
            content: '📥';
            margin-right: 8px;
        }
        .motivation {
            color: #28a745;
            font-size: 1.2em;
            text-align: center;
            margin: 20px 0;
            font-style: italic;
            background-color: #e6f4ea;
            padding: 10px;
            border-radius: 8px;
        }
        .badge {
            background-color: #28a745;
            color: white;
            padding: 12px 24px;
            border-radius: 20px;
            font-size: 1.2em;
            text-align: center;
            margin: 20px auto;
            display: block;
            width: fit-content;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
        }
        .insights {
            background-color: #e9f7ff;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            border-left: 4px solid #007bff;
        }
        .grade {
            font-size: 1.6em;
            font-weight: 700;
            text-align: center;
            margin: 20px 0;
            padding: 12px;
            border-radius: 8px;
        }
        .grade-excellent { background-color: #28a745; color: white; }
        .grade-good { background-color: #ffd700; color: black; }
        .grade-needs-improvement { background-color: #ff4d4d; color: white; }
        .grade-critical { background-color: #b30000; color: white; }
        .stepper {
            display: flex;
            justify-content: center;
            margin: 20px 0;
        }
        .step {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            background-color: #ccc;
            margin: 0 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 700;
            transition: background-color 0.3s ease;
        }
        .step.active {
            background-color: #007bff;
        }
        .step.completed {
            background-color: #28a745;
        }
        .card {
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
        }
        /* Responsive Design */
        @media (max-width: 768px) {
            .main-container {
                padding: 15px;
                margin: 10px;
            }
            .header {
                font-size: 2em;
            }
            .subheader {
                font-size: 1.4em;
            }
            .stButton>button {
                padding: 10px 20px;
                font-size: 0.9em;
            }
            .stRadio>label {
                font-size: 1em;
                padding: 10px;
            }
            .stepper {
                flex-wrap: wrap;
            }
            .step {
                width: 25px;
                height: 25px;
                margin: 5px;
                font-size: 0.9em;
            }
        }
        /* Accessibility */
        [role="radiogroup"] {
            margin: 10px 0;
        }
        [role="radio"] {
            cursor: pointer;
        }
        [role="radio"]:focus {
            outline: 2px solid #007bff;
            outline-offset: 2px;
        }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'language' not in st.session_state:
    st.session_state.language = "Español"
if 'responses' not in st.session_state:
    st.session_state.responses = {}
if 'current_category' not in st.session_state:
    st.session_state.current_category = 0
if 'prev_language' not in st.session_state:
    st.session_state.prev_language = st.session_state.language
if 'show_intro' not in st.session_state:
    st.session_state.show_intro = True

# Bilingual support
LANG = st.sidebar.selectbox(
    "Idioma / Language", 
    ["Español", "English"], 
    help="Selecciona tu idioma preferido / Select your preferred language",
    key="language_select"
)

# Reset session state if language changes
if LANG != st.session_state.prev_language:
    st.session_state.current_category = 0
    st.session_state.responses = {}
    st.session_state.prev_language = LANG
    st.session_state.show_intro = True

st.session_state.language = LANG

# Introductory modal (simulated with expander)
if st.session_state.show_intro:
    with st.expander("¡Bienvenido a la Auditoría! / Welcome to the Audit!", expanded=True):
        st.markdown(
            """
            Esta auditoría ayuda a crear un lugar de trabajo ético, lean y centrado en las personas. Responde preguntas en 5 categorías (5–10 minutos). Tus aportes generarán un informe detallado con recomendaciones accionables.
            
            **Pasos**:
            1. Responde las preguntas de cada categoría.
            2. Revisa tus respuestas.
            3. Genera y descarga tu informe.
            
            ¡Empecemos!
            """
            if LANG == "Español" else
            """
            This audit helps create an ethical, lean, and human-centered workplace. Answer questions across 5 categories (5–10 minutes). Your insights will generate a detailed report with actionable recommendations.
            
            **Steps**:
            1. Answer questions for each category.
            2. Review your responses.
            3. Generate and download your report.
            
            Let’s get started!
            """
        )
        if st.button("Iniciar Auditoría / Start Audit"):
            st.session_state.show_intro = False

# Main content (only shown after closing intro)
if not st.session_state.show_intro:
    # Main container
    with st.container():
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        
        # Header
        st.markdown(
            '<div class="header">¡Construye un Lugar de Trabajo Ético y Centrado en las Personas!</div>' if LANG == "Español" else 
            '<div class="header">Shape an Ethical & Human-Centered Workplace!</div>', 
            unsafe_allow_html=True
        )
        
        # Likert scale labels
        labels = {
            "Español": ["Nunca", "Raramente", "A veces", "A menudo", "Siempre"],
            "English": ["Not at All", "Rarely", "Sometimes", "Often", "Always"]
        }

        # Audit categories and questions
        questions = {
            "Empoderamiento de Empleados": {
                "Español": [
                    "¿Con qué frecuencia se implementan las sugerencias de los empleados para mejorar los procesos o la cultura laboral?",
                    "¿Proporciona el lugar de trabajo talleres o capacitaciones regulares para desarrollar las habilidades y la confianza de los empleados?",
                    "¿Con qué frecuencia tienen los empleados oportunidades de liderar proyectos o iniciativas que impacten a su equipo?",
                    "¿Qué tan efectivamente fomenta el lugar de trabajo el diálogo abierto entre empleados y la gerencia?"
                ],
                "English": [
                    "How often are employee suggestions implemented to improve workplace processes or culture?",
                    "Does the workplace provide regular workshops or training to develop employee skills and confidence?",
                    "How frequently do employees have opportunities to lead projects or initiatives that impact their team?",
                    "How effectively does the workplace encourage open dialogue between employees and management?"
                ]
            },
            "Liderazgo Ético": {
                "Español": [
                    "¿Con qué consistencia comparten los líderes actualizaciones claras sobre decisiones que afectan a los empleados?",
                    "¿Involucra activamente el liderazgo a los empleados en la formación de políticas laborales o estándares éticos?",
                    "¿Con qué frecuencia reconocen y recompensan los líderes el comportamiento ético o las contribuciones al bienestar del equipo?"
                ],
                "English": [
                    "How consistently do leaders share clear updates on decisions affecting employees?",
                    "Does leadership actively involve employees in shaping workplace policies or ethical standards?",
                    "How often do leaders recognize and reward ethical behavior or contributions to team well-being?"
                ]
            },
            "Operaciones Centradas en las Personas": {
                "Español": [
                    "¿Qué tan efectivamente incorporan los procesos lean la retroalimentación de los empleados para reducir la carga de trabajo innecesaria?",
                    "¿Revisa regularmente el lugar de trabajo las prácticas operativas para asegurar que apoyen el bienestar de los empleados?",
                    "¿Con qué frecuencia se capacita a los empleados para usar herramientas lean de manera que fomenten la colaboración y el respeto?"
                ],
                "English": [
                    "How effectively do lean processes incorporate employee feedback to reduce unnecessary workload?",
                    "Does the workplace regularly review operational practices to ensure they support employee well-being?",
                    "How often are employees trained to use lean tools in ways that enhance collaboration and respect?"
                ]
            },
            "Prácticas Sostenibles y Éticas": {
                "Español": [
                    "¿Reduce activamente el lugar de trabajo el desperdicio (por ejemplo, energía, materiales) a través de iniciativas lean?",
                    "¿Con qué consistencia se asocia el lugar de trabajo con proveedores que priorizan estándares laborales justos y ambientales?",
                    "¿Con qué frecuencia participan los empleados en proyectos de sostenibilidad que benefician al lugar de trabajo o la comunidad?"
                ],
                "English": [
                    "Does the workplace actively reduce waste (e.g., energy, materials) through lean initiatives?",
                    "How consistently does the workplace partner with suppliers who prioritize fair labor and environmental standards?",
                    "How often are employees involved in sustainability projects that benefit the workplace or community?"
                ]
            },
            "Bienestar y Equilibrio": {
                "Español": [
                    "¿Con qué consistencia ofrece el lugar de trabajo recursos (por ejemplo, asesoramiento, horarios flexibles) para gestionar el estrés y la carga de trabajo?",
                    "¿Realiza el lugar de trabajo revisiones regulares para evaluar y abordar el agotamiento o la fatiga de los empleados?",
                    "¿Qué tan efectivamente promueve el lugar de trabajo una cultura donde los empleados se sientan seguros para expresar desafíos personales o profesionales?"
                ],
                "English": [
                    "How consistently does the workplace offer resources (e.g., counseling, flexible schedules) to manage stress and workload?",
                    "Does the workplace conduct regular check-ins to assess and address employee burnout or fatigue?",
                    "How effectively does the workplace promote a culture where employees feel safe to express personal or professional challenges?"
                ]
            }
        }

        # Validate question counts
        for cat in questions:
            if len(questions[cat]["Español"]) != len(questions[cat]["English"]):
                st.error(f"Discrepancia en el número de preguntas en la categoría {cat} entre Español e Inglés.")
                st.stop()

        # Initialize responses
        if not st.session_state.responses or len(st.session_state.responses) != len(questions):
            st.session_state.responses = {cat: [0] * len(questions[cat][LANG]) for cat in questions}

        # Progress stepper
        categories = list(questions.keys())
        stepper_html = '<div class="stepper">'
        for i, cat in enumerate(categories):
            status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
            stepper_html += f'<div class="step {status}" title="{cat}">{i+1}</div>'
        stepper_html += '</div>'
        st.markdown(stepper_html, unsafe_allow_html=True)

        # Progress feedback
        st.markdown(
            f'<div class="motivation">{st.session_state.current_category + 1}/{len(categories)} categorías completadas</div>'
            if LANG == "Español" else
            f'<div class="motivation">{st.session_state.current_category + 1}/{len(categories)} categories completed</div>',
            unsafe_allow_html=True
        )

        # Category questions
        category_index = st.session_state.current_category
        category_index = min(category_index, len(categories) - 1)
        category = categories[category_index]
        
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
            for idx, q in enumerate(questions[category][LANG]):
                with st.container():
                    st.markdown(f"**{q}**")
                    score = st.radio(
                        "",
                        list(range(1, 6)),
                        format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
                        key=f"{category}_{idx}",
                        horizontal=True,
                        help="Selecciona una respuesta para compartir tus ideas." if LANG == "Español" else
                             "Select a response to share your insights."
                    )
                    st.session_state.responses[category][idx] = score
            st.markdown('</div>', unsafe_allow_html=True)

        # Navigation buttons
        col1, col2 = st.columns([1, 1], gap="small")
        with col1:
            if st.button(
                "⬅ Anterior" if LANG == "Español" else "⬅ Previous",
                disabled=category_index == 0,
                use_container_width=True
            ):
                st.session_state.current_category = max(category_index - 1, 0)
                st.rerun()
        with col2:
            if st.button(
                "Siguiente ➡" if LANG == "Español" else "Next ➡",
                disabled=category_index == len(categories) - 1,
                use_container_width=True
            ):
                if all(score != 0 for score in st.session_state.responses[category]):
                    st.session_state.current_category = min(category_index + 1, len(categories) - 1)
                    st.rerun()
                else:
                    st.error(
                        "Por favor, responde todas las preguntas antes de continuar." if LANG == "Español" else
                        "Please answer all questions before proceeding."
                    )

        # Review and Submit
        if category_index == len(categories) - 1:
            with st.expander("Revisar Respuestas / Review Responses", expanded=False):
                for cat in categories:
                    st.markdown(f"**{cat}**")
                    for idx, q in enumerate(questions[cat][LANG]):
                        score = st.session_state.responses[cat][idx]
                        st.markdown(f"- {q}: {score} - {labels[LANG][score-1] if score > 0 else 'No respondida' if LANG == 'Español' else 'Not answered'}")
                if st.button("Enviar Auditoría / Submit Audit"):
                    if all(all(score != 0 for score in scores) for scores in st.session_state.responses.values()):
                        st.session_state.current_category = len(categories)  # Move to report
                        st.rerun()
                    else:
                        st.error(
                            "Por favor, responde todas las preguntas en todas las categorías." if LANG == "Español" else
                            "Please answer all questions in all categories."
                        )

        # Grading matrix function
        def get_grade(score):
            if score >= 85:
                return (
                    "Excelente" if LANG == "Español" else "Excellent",
                    "Tu lugar de trabajo destaca en prácticas éticas, lean y centradas en las personas. ¡Continúa manteniendo estas fortalezas!" if LANG == "Español" else
                    "Your workplace excels in ethical, lean, and human-centered practices. Continue maintaining these strengths!",
                    "grade-excellent"
                )
            elif score >= 70:
                return (
                    "Bueno" if LANG == "Español" else "Good",
                    "Tu lugar de trabajo es sólido pero tiene espacio para mejorar áreas específicas para un rendimiento óptimo. Considera mejoras dirigidas." if LANG == "Español" else
                    "Your workplace is strong but has room to refine specific areas for optimal performance. Consider targeted improvements.",
                    "grade-good"
                )
            elif score >= 50:
                return (
                    "Necesita Mejora" if LANG == "Español" else "Needs Improvement",
                    "Tu lugar de trabajo requiere intervenciones específicas para abordar debilidades moderadas. Prioriza acciones en áreas de baja puntuación." if LANG == "Español" else
                    "Your workplace requires targeted interventions to address moderate weaknesses. Prioritize action in low-scoring areas.",
                    "grade-needs-improvement"
                )
            else:
                return (
                    "Crítico" if LANG == "Español" else "Critical",
                    "Tu lugar de trabajo tiene problemas significativos que requieren acción urgente y comprehensiva. Involucra a expertos para un plan de transformación." if LANG == "Español" else
                    "Your workplace has significant issues requiring urgent, comprehensive action. Engage experts for a transformation plan.",
                    "grade-critical"
                )

        # Generate report
        if st.session_state.current_category >= len(categories):
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(
                f'<div class="subheader">{"Tu Informe de Impacto en el Lugar de Trabajo" if LANG == "Español" else "Your Workplace Impact Report"}</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por construir un lugar de trabajo ético!</div>' if LANG == "Español" else
                '<div class="badge">🏆 Audit Completed! Thank you for shaping an ethical workplace!</div>', 
                unsafe_allow_html=True
            )
            
            # Calculate scores
            results = {cat: sum(scores) for cat, scores in st.session_state.responses.items()}
            df = pd.DataFrame.from_dict(results, orient="index", columns=["Puntuación" if LANG == "Español" else "Score"])
            df["Porcentaje" if LANG == "Español" else "Percent"] = [((score / (len(questions[cat][LANG]) * 5)) * 100) for cat, score in results.items()]
            df["Prioridad" if LANG == "Español" else "Priority"] = df["Porcentaje" if LANG == "Español" else "Percent"].apply(lambda x: "Alta" if x < 50 else "Media" if x < 70 else "Baja" if LANG == "Español" else "High" if x < 50 else "Medium" if x < 70 else "Low")
            
            # Overall score and grade
            overall_score = df["Porcentaje" if LANG == "Español" else "Percent"].mean()
            grade, grade_description, grade_class = get_grade(overall_score)
            st.markdown(
                f'<div class="grade {grade_class}">Calificación General del Lugar de Trabajo: {grade} ({overall_score:.1f}%)</div>' if LANG == "Español" else
                f'<div class="grade {grade_class}">Overall Workplace Grade: {grade} ({overall_score:.1f}%)</div>',
                unsafe_allow_html=True
            )
            st.markdown(grade_description, unsafe_allow_html=True)

            # Color-coded dataframe
            def color_percent(val):
                color = '#ff4d4d' if val < 50 else '#ffd700' if val < 70 else '#28a745'
                return f'background-color: {color}; color: white;'
            
            st.markdown(
                "Puntuaciones por debajo del 50% (rojo) requieren acción urgente, 50–69% (amarillo) sugieren mejoras, y por encima del 70% (verde) indican fortalezas." if LANG == "Español" else
                "Scores below 50% (red) need urgent action, 50–69% (yellow) suggest improvement, and above 70% (green) indicate strengths.",
                unsafe_allow_html=True
            )
            styled_df = df.style.applymap(color_percent, subset=["Porcentaje" if LANG == "Español" else "Percent"]).format({"Porcentaje" if LANG == "Español" else "Percent": "{:.1f}%"})
            st.dataframe(styled_df, use_container_width=True)

            # Interactive bar chart
            fig = px.bar(
                df.reset_index(),
                y="index",
                x="Porcentaje" if LANG == "Español" else "Percent",
                orientation='h',
                title="Fortalezas y Oportunidades del Lugar de Trabajo" if LANG == "Español" else "Workplace Strengths and Opportunities",
                labels={"index": "Categoría" if LANG == "Español" else "Category", "Porcentaje" if LANG == "Español" else "Percent": "Puntuación (%)" if LANG == "Español" else "Score (%)"},
                color="Porcentaje" if LANG == "Español" else "Percent",
                color_continuous_scale=["#ff4d4d", "#ffd700", "#28a745"],
                range_x=[0, 100]
            )
            fig.add_vline(x=70, line_dash="dash", line_color="blue", annotation_text="Objetivo (70%)" if LANG == "Español" else "Target (70%)", annotation_position="top")
            for i, row in df.iterrows():
                if row["Porcentaje" if LANG == "Español" else "Percent"] < 70:
                    fig.add_annotation(
                        x=row["Porcentaje" if LANG == "Español" else "Percent"], y=i,
                        text="Prioridad" if LANG == "Español" else "Priority", showarrow=True, arrowhead=2, ax=20, ay=-30,
                        font=dict(color="red", size=12)
                    )
            fig.update_layout(
                height=400,
                showlegend=False,
                title_x=0.5,
                xaxis_title="Puntuación (%)" if LANG == "Español" else "Score (%)",
                yaxis_title="Categoría" if LANG == "Español" else "Category",
                coloraxis_showscale=False
            )
            st.plotly_chart(fig, use_container_width=True)

            # Question-level breakdown
            st.markdown(
                "<div class='subheader'>Análisis Detallado: Perspectivas a Nivel de Pregunta</div>" if LANG == "Español" else 
                "<div class='subheader'>Drill Down: Question-Level Insights</div>",
                unsafe_allow_html=True
            )
            selected_category = st.selectbox(
                "Seleccionar Categoría para Explorar" if LANG == "Español" else "Select Category to Explore",
                categories,
                key="category_explore"
            )
            question_scores = pd.DataFrame({
                "Pregunta" if LANG == "Español" else "Question": questions[selected_category][LANG],
                "Puntuación" if LANG == "Español" else "Score": [score / 5 * 100 for score in st.session_state.responses[selected_category]]
            })
            fig_questions = px.bar(
                question_scores,
                x="Puntuación" if LANG == "Español" else "Score",
                y="Pregunta" if LANG == "Español" else "Question",
                orientation='h',
                title=f"Puntuaciones de Preguntas para {selected_category}" if LANG == "Español" else f"Question Scores for {selected_category}",
                labels={"Puntuación" if LANG == "Español" else "Score": "Puntuación (%)" if LANG == "Español" else "Score (%)", "Pregunta" if LANG == "Español" else "Question": "Pregunta" if LANG == "Español" else "Question"},
                color="Puntuación" if LANG == "Español" else "Score",
                color_continuous_scale=["#ff4d4d", "#ffd700", "#28a745"],
                range_x=[0, 100]
            )
            fig_questions.update_layout(
                height=300 + len(question_scores) * 50,
                showlegend=False,
                title_x=0.5,
                xaxis_title="Puntuación (%)" if LANG == "Español" else "Score (%)",
                yaxis_title="Pregunta" if LANG == "Español" else "Question",
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_questions, use_container_width=True)

            # Actionable insights
            st.markdown(
                "<div class='subheader'>Perspectivas Accionables</div>" if LANG == "Español" else "<div class='subheader'>Actionable Insights</div>",
                unsafe_allow_html=True
            )
            insights = []
            recommendations = {
                "Empoderamiento de Empleados": {
                    0: "Implementa un buzón de sugerencias o sesiones regulares de retroalimentación para actuar sobre las ideas de los empleados.",
                    1: "Programa talleres o capacitaciones mensuales para mejorar las habilidades de los empleados.",
                    2: "Crea oportunidades para que los empleados lideren proyectos o iniciativas pequeñas.",
                    3: "Organiza foros abiertos o reuniones regulares para fomentar el diálogo con la gerencia."
                },
                "Liderazgo Ético": {
                    0: "Establece actualizaciones mensuales o boletines para compartir decisiones de liderazgo.",
                    1: "Forma un grupo asesor de empleados para moldear las políticas laborales.",
                    2: "Introduce un programa de reconocimiento para comportamientos éticos y contribuciones al bienestar del equipo."
                },
                "Operaciones Centradas en las Personas": {
                    0: "Incorpora retroalimentación de empleados en revisiones de procesos lean para reducir la carga de trabajo.",
                    1: "Realiza auditorías trimestrales de prácticas operativas para evaluar su impacto en el bienestar.",
                    2: "Capacita a los empleados en herramientas lean que promuevan colaboración y respeto."
                },
                "Prácticas Sostenibles y Éticas": {
                    0: "Lanza una iniciativa de reducción de desperdicios con roles claros para los empleados.",
                    1: "Audita a los proveedores para asegurar estándares laborales justos y ambientales.",
                    2: "Involucra a los empleados en proyectos de sostenibilidad, como reciclaje o alcance comunitario."
                },
                "Bienestar y Equilibrio": {
                    0: "Ofrece servicios de asesoramiento o horarios flexibles para gestionar el estrés.",
                    1: "Implementa revisiones mensuales para monitorear el agotamiento y la fatiga.",
                    2: "Capacita a los gerentes para fomentar una cultura de seguridad psicológica."
                }
            } if LANG == "Español" else {
                "Empowering Employees": {
                    0: "Implement a suggestion box or regular feedback sessions to act on employee ideas.",
                    1: "Schedule monthly workshops or training programs to boost employee skills.",
                    2: "Create opportunities for employees to lead small projects or initiatives.",
                    3: "Host regular town halls or open forums to foster dialogue with management."
                },
                "Ethical Leadership": {
                    0: "Establish monthly updates or newsletters to share leadership decisions.",
                    1: "Form an employee advisory group to shape workplace policies.",
                    2: "Introduce a recognition program for ethical behavior and well-being contributions."
                },
                "Human-Centered Operations": {
                    0: "Incorporate employee feedback into lean process reviews to reduce workload.",
                    1: "Conduct quarterly audits of operational practices for well-being impact.",
                    2: "Provide training on lean tools emphasizing collaboration and respect."
                },
                "Sustainable and Ethical Practices": {
                    0: "Launch a waste reduction initiative with clear employee roles.",
                    1: "Audit suppliers for fair labor and environmental standards.",
                    2: "Engage employees in sustainability projects, like recycling or community outreach."
                },
                "Well-Being and Balance": {
                    0: "Offer counseling services or flexible schedules to manage stress.",
                    1: "Implement monthly check-ins to monitor burnout and fatigue.",
                    2: "Train managers to foster a culture of psychological safety."
                }
            }
            for cat in categories:
                if df.loc[cat, "Porcentaje" if LANG == "Español" else "Percent"] < 50:
                    insights.append(
                        f"**{cat}** obtuvo {df.loc[cat, 'Porcentaje' if LANG == 'Español' else 'Percent']:.1f}% (Alta Prioridad). Enfócate en mejoras inmediatas." if LANG == "Español" else
                        f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (High Priority). Focus on immediate improvements."
                    )
                elif df.loc[cat, "Porcentaje" if LANG == "Español" else "Percent"] < 70:
                    insights.append(
                        f"**{cat}** obtuvo {df.loc[cat, 'Porcentaje' if LANG == 'Español' else 'Percent']:.1f}% (Prioridad Media). Considera acciones específicas." if LANG == "Español" else
                        f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (Medium Priority). Consider targeted actions."
                    )
            if insights:
                st.markdown(
                    "<div class='insights'>" + "<br>".join(insights) + "</div>",
                    unsafe_allow_html=True
                )
            else:
                st.markdown(
                    "<div class='insights'>¡Todas las categorías obtuvieron más del 70%! Continúa manteniendo estas fortalezas.</div>" if LANG == "Español" else 
                    "<div class='insights'>All categories scored above 70%! Continue maintaining these strengths.</div>",
                    unsafe_allow_html=True
                )

            # LEAN 2.0 Institute Advertisement
            st.markdown(
                "<div class='subheader'>Optimiza tu Lugar de Trabajo con LEAN 2.0 Institute</div>" if LANG == "Español" else 
                "<div class='subheader'>Optimize Your Workplace with LEAN 2.0 Institute</div>",
                unsafe_allow_html=True
            )
            ad_text = []
            if overall_score < 85:
                ad_text.append(
                    "Los resultados de tu auditoría indican oportunidades para crear un lugar de trabajo óptimo. LEAN 2.0 Institute ofrece servicios de consultoría personalizados para transformar tu lugar de trabajo en un entorno ético, lean y centrado en las personas." if LANG == "Español" else
                    "Your audit results indicate opportunities to create an optimal workplace. LEAN 2.0 Institute offers tailored consulting services to transform your workplace into an ethical, lean, and human-centered environment."
                )
                if df["Porcentaje" if LANG == "Español" else "Percent"].min() < 70:
                    low_categories = df[df["Porcentaje" if LANG == "Español" else "Percent"] < 70].index.tolist()
                    services = {
                        "Empoderamiento de Empleados": "Programas de Compromiso y Empoderamiento de Empleados",
                        "Liderazgo Ético": "Capacitación en Liderazgo y Talleres de Toma de Decisiones Éticas",
                        "Operaciones Centradas en las Personas": "Optimización de Procesos Lean con Diseño Centrado en las Personas",
                        "Prácticas Sostenibles y Éticas": "Estrategia de Sostenibilidad y Consultoría de Cadena de Suministro Ética",
                        "Bienestar y Equilibrio": "Programas de Bienestar y Resiliencia de Empleados"
                    } if LANG == "Español" else {
                        "Empowering Employees": "Employee Engagement and Empowerment Programs",
                        "Ethical Leadership": "Leadership Coaching and Ethical Decision-Making Workshops",
                        "Human-Centered Operations": "Lean Process Optimization with Human-Centered Design",
                        "Sustainable and Ethical Practices": "Sustainability Strategy and Ethical Supply Chain Consulting",
                        "Well-Being and Balance": "Employee Well-Being and Resilience Programs"
                    }
                    ad_text.append(
                        f"Las áreas clave para mejorar incluyen {', '.join(low_categories)}. LEAN 2.0 Institute se especializa en: {', '.join([services[cat] for cat in low_categories])}." if LANG == "Español" else
                        f"Key areas for improvement include {', '.join(low_categories)}. LEAN 2.0 Institute specializes in: {', '.join([services[cat] for cat in low_categories])}."
                    )
            else:
                ad_text.append(
                    "¡Felicidades por tu excelente lugar de trabajo! Asóciate con LEAN 2.0 Institute para mantener y mejorar tus fortalezas mediante estrategias lean avanzadas y desarrollo de liderazgo." if LANG == "Español" else
                    "Congratulations on your excellent workplace! Partner with LEAN 2.0 Institute to sustain and enhance your strengths through advanced lean strategies and leadership development."
                )
            ad_text.append(
                "¡Contáctanos en www.lean2institute.com o info@lean2institute.com para una consulta y lleva tu lugar de trabajo al siguiente nivel!" if LANG == "Español" else
                "Contact us at www.lean2institute.com or info@lean2institute.com for a consultation to elevate your workplace to the next level!"
            )
            st.markdown("<div class='insights'>" + "<br>".join(ad_text) + "</div>", unsafe_allow_html=True)

            # Download buttons
            st.markdown(
                '<div class="subheader">Descarga tu Informe</div>' if LANG == "Español" else
                '<div class="subheader">Download Your Report</div>',
                unsafe_allow_html=True
            )
            col1, col2 = st.columns(2)
            
            # PDF Report
            font_path = "DejaVuSans.ttf"
            if not os.path.exists(font_path):
                st.error("El archivo de fuente 'DejaVuSans.ttf' no se encuentra. Por favor, descárgalo desde https://dejavu-fonts.github.io/ y colócalo en el directorio del proyecto." if LANG == "Español" else
                         "Font file 'DejaVuSans.ttf' not found. Please download it from https://dejavu-fonts.github.io/ and place it in the project directory.")
                st.stop()

            with col1:
                try:
                    pdf = FPDF()
                    pdf.set_margins(15, 15, 15)
                    pdf.add_font('DejaVu', '', font_path, uni=True)
                    pdf.add_page()
                    
                    # Overall Grade
                    pdf.set_font('DejaVu', 'B', 12)
                    pdf.set_text_color(51)
                    pdf.multi_cell(0, 10, f"Calificación General del Lugar de Trabajo: {grade} ({overall_score:.1f}%)" if LANG == "Español" else
                                        f"Overall Workplace Grade: {grade} ({overall_score:.1f}%)")
                    pdf.set_font('DejaVu', '', 12)
                    pdf.multi_cell(0, 10, grade_description)
                    pdf.ln(5)
                    
                    # Audit Results
                    pdf.set_font('DejaVu', 'B', 12)
                    pdf.multi_cell(0, 10, "Resultados de la Auditoría" if LANG == "Español" else "Audit Results")
                    pdf.set_font('DejaVu', '', 12)
                    pdf.ln(5)
                    for cat, row in df.iterrows():
                        pdf.multi_cell(0, 10, f"{cat}: {row['Porcentaje' if LANG == 'Español' else 'Percent']:.1f}% (Prioridad: {row['Prioridad' if LANG == 'Español' else 'Priority']})")
                    
                    # Action Plan
                    pdf.add_page()
                    pdf.set_font('DejaVu', 'B', 12)
                    pdf.multi_cell(0, 10, "Plan de Acción" if LANG == "Español" else "Action Plan")
                    pdf.set_font('DejaVu', '', 12)
                    pdf.ln(5)
                    for cat in categories:
                        if df.loc[cat, "Porcentaje" if LANG == "Español" else "Percent"] < 70:
                            pdf.set_font('DejaVu', 'B', 12)
                            pdf.multi_cell(0, 10, cat)
                            pdf.set_font('DejaVu', '', 12)
                            for idx, score in enumerate(st.session_state.responses[cat]):
                                if score / 5 * 100 < 70:
                                    question = questions[cat][LANG][idx]
                                    rec = recommendations[cat][idx]
                                    pdf.multi_cell(0, 10, f"- {question}: {rec}")
                            pdf.ln(5)
                    
                    # LEAN 2.0 Institute Advertisement
                    pdf.add_page()
                    pdf.set_font('DejaVu', 'B', 12)
                    pdf.multi_cell(0, 10, "Asóciate con LEAN 2.0 Institute" if LANG == "Español" else "Partner with LEAN 2.0 Institute")
                    pdf.set_font('DejaVu', '', 12)
                    pdf.ln(5)
                    for text in ad_text:
                        pdf.multi_cell(0, 10, text)
                    
                    # Certificate
                    pdf.add_page()
                    pdf.set_font('DejaVu', 'B', 16)
                    pdf.set_text_color(40, 167, 69)
                    pdf.multi_cell(0, 10, "Certificado de Finalización" if LANG == "Español" else "Certificate of Completion", align="C")
                    pdf.ln(10)
                    pdf.set_font('DejaVu', '', 12)
                    pdf.set_text_color(51)
                    pdf.multi_cell(
                        0, 10, 
                        "¡Felicidades por completar la Auditoría del Lugar de Trabajo Ético! Tus aportes están ayudando a crear un lugar de trabajo más ético, centrado en las personas y sostenible." if LANG == "Español" else 
                        "Congratulations on completing the Ethical Workplace Audit! Your insights are helping to create a more ethical, human-centered, and sustainable workplace."
                    )
                    
                    pdf_output = io.BytesIO()
                    pdf.output(pdf_output)
                    pdf_output.seek(0)
                    b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
                    href_pdf = (
                        f'<a href="data:application/pdf;base64,{b64_pdf}" download="informe_auditoria_lugar_trabajo_etico.pdf" class="download-link">Descargar Informe PDF y Plan de Acción</a>' if LANG == "Español" else 
                        f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_workplace_audit_report.pdf" class="download-link">Download PDF Report & Action Plan</a>'
                    )
                    st.markdown(href_pdf, unsafe_allow_html=True)
                    pdf_output.close()
                except Exception as e:
                    st.error(f"No se pudo generar el PDF: {str(e)}" if LANG == "Español" else f"Failed to generate PDF: {str(e)}")
                    st.markdown(
                        "Por favor, asegura que el archivo de fuente 'DejaVuSans.ttf' esté en el directorio del proyecto. Descárgalo desde https://dejavu-fonts.github.io/ si falta." if LANG == "Español" else
                        "Please ensure the 'DejaVuSans.ttf' font file is in the project directory. Download it from https://dejavu-fonts.github.io/ if missing.",
                        unsafe_allow_html=True
                    )

            # Excel export
            with col2:
                try:
                    excel_output = io.BytesIO()
                    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                        df.to_excel(
                            writer, 
                            sheet_name='Resultados de la Auditoría' if LANG == "Español" else 'Audit Results', 
                            float_format="%.1f"
                        )
                        pd.DataFrame({"Puntuación General" if LANG == "Español" else "Overall Score": [overall_score], "Calificación" if LANG == "Español" else "Grade": [grade]}).to_excel(
                            writer, 
                            sheet_name='Resumen' if LANG == "Español" else 'Summary', 
                            index=False
                        )
                    excel_output.seek(0)
                    b64_excel = base64.b64encode(excel_output.getvalue()).decode()
                    href_excel = (
                        f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="resultados_auditoria_lugar_trabajo_etico.xlsx" class="download-link">Descargar Informe Excel</a>' if LANG == "Español" else 
                        f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_workplace_audit_results.xlsx" class="download-link">Download Excel Report</a>'
                    )
                    st.markdown(href_excel, unsafe_allow_html=True)
                    excel_output.close()
                except ImportError:
                    st.error("La exportación a Excel requiere 'xlsxwriter'. Por favor, instálalo usando `pip install xlsxwriter`." if LANG == "Español" else
                             "Excel export requires 'xlsxwriter'. Please install it using `pip install xlsxwriter`.")
                except Exception as e:
                    st.error(f"No se pudo generar el archivo Excel: {str(e)}" if LANG == "Español" else f"Failed to generate Excel file: {str(e)}")

            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

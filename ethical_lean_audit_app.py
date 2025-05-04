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
            content: '📞';
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
            cursor: pointer;
        }
        .step.active {
            background-color: #007bff;
        }
        .step.completed {
            background-color: #28a745;
        }
        .step:hover {
            background-color: #0056b3;
        }
        .card {
            background-color: #f9f9f9;
            padding: 20px;
            border-radius: 10px;
            margin: 15px 0;
            box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            transition: opacity 0.3s ease, transform 0.3s ease;
        }
        .sticky-nav {
            position: sticky;
            bottom: 20px;
            background-color: #ffffff;
            padding: 10px;
            border-radius: 8px;
            box-shadow: 0 -2px 6px rgba(0, 0, 0, 0.1);
            z-index: 1000;
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
            Esta auditoría está diseñada para directivos y profesionales de Recursos Humanos para evaluar de manera objetiva el entorno laboral. Responde preguntas en 5 categorías (5–10 minutos) con datos específicos y ejemplos verificables. Tus respuestas son confidenciales y generarán un informe detallado con recomendaciones accionables. Al completar la auditoría, contacta a LEAN 2.0 Institute en <a href="https://lean2institute.mystrikingly.com/" target="_blank">https://lean2institute.mystrikingly.com/</a> para consultas personalizadas.
            
            **Pasos**:
            1. Responde las preguntas de cada categoría.
            2. Genera y descarga tu informe.
            
            ¡Empecemos!
            """
            if LANG == "Español" else
            """
            This audit is designed for directors and HR professionals to objectively assess the workplace environment. Answer questions across 5 categories (5–10 minutes) with specific data and verifiable examples. Your responses are confidential and will generate a detailed report with actionable recommendations. Upon completion, contact LEAN 2.0 Institute at <a href="https://lean2institute.mystrikingly.com/" target="_blank">https://lean2institute.mystrikingly.com/</a> for personalized consultation.
            
            **Steps**:
            1. Answer questions for each category.
            2. Generate and download your report.
            
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
            '<div class="header">¡Evalúa y Mejora tu Lugar de Trabajo!</div>' if LANG == "Español" else 
            '<div class="header">Assess and Enhance Your Workplace!</div>', 
            unsafe_allow_html=True
        )
        
        # Likert scale labels for different question types
        labels = {
            "percentage": {
                "Español": ["0%", "25%", "50%", "75%", "100%"],
                "English": ["0%", "25%", "50%", "75%", "100%"]
            },
            "frequency": {
                "Español": ["Nunca", "Anualmente", "Semestralmente", "Trimestralmente", "Mensualmente"],
                "English": ["Never", "Annually", "Semi-Annually", "Quarterly", "Monthly"]
            },
            "count": {
                "Español": ["Ninguno", "1–10% de la fuerza laboral", "11–25%", "26–50%", ">50%"],
                "English": ["None", "1–10% of workforce", "11–25%", "26–50%", ">50%"]
            }
        }

        # Audit categories and questions
        questions = {
            "Empoderamiento de Empleados": {
                "Español": [
                    ("¿Qué porcentaje de sugerencias de empleados presentadas en los últimos 12 meses fueron implementadas con resultados documentados?", "percentage"),
                    ("¿Cuántas horas de capacitación en habilidades profesionales se ofrecieron por empleado en el último año?", "count"),
                    ("En los últimos 12 meses, ¿cuántos empleados lideraron proyectos o iniciativas con presupuesto asignado?", "count"),
                    ("¿Con qué frecuencia (en meses) se realizan foros o reuniones formales para que los empleados compartan retroalimentación con la gerencia?", "frequency")
                ],
                "English": [
                    ("What percentage of employee suggestions submitted in the past 12 months were implemented with documented outcomes?", "percentage"),
                    ("How many hours of professional skills training were provided per employee in the past year?", "count"),
                    ("In the past 12 months, how many employees led projects or initiatives with allocated budgets?", "count"),
                    ("How frequently (in months) are formal forums or meetings held for employees to share feedback with management?", "frequency")
                ]
            },
            "Liderazgo Ético": {
                "Español": [
                    ("¿En cuántas ocasiones en los últimos 12 meses los líderes compartieron actualizaciones escritas sobre decisiones que afectan a los empleados?", "count"),
                    ("¿Qué porcentaje de políticas laborales nuevas o revisadas en el último año incluyó consulta formal con empleados?", "percentage"),
                    ("¿Cuántos casos de comportamiento ético destacado fueron reconocidos formalmente (por ejemplo, con premios o bonos) en los últimos 12 meses?", "count")
                ],
                "English": [
                    ("How many times in the past 12 months have leaders shared written updates on decisions affecting employees?", "count"),
                    ("What percentage of new or revised workplace policies in the past year included formal employee consultation?", "percentage"),
                    ("How many instances of exemplary ethical behavior were formally recognized (e.g., with awards or bonuses) in the past 12 months?", "count")
                ]
            },
            "Operaciones Centradas en las Personas": {
                "Español": [
                    ("¿Qué porcentaje de procesos lean revisados en los últimos 12 meses incorporó retroalimentación de empleados para reducir tareas redundantes?", "percentage"),
                    ("¿Con qué frecuencia (en meses) se auditan las prácticas operativas para evaluar su impacto en el bienestar de los empleados?", "frequency"),
                    ("¿Cuántos empleados recibieron capacitación en herramientas lean con énfasis en colaboración en el último año?", "count")
                ],
                "English": [
                    ("What percentage of lean processes revised in the past 12 months incorporated employee feedback to reduce redundant tasks?", "percentage"),
                    ("How frequently (in months) are operational practices audited to assess their impact on employee well-being?", "frequency"),
                    ("How many employees received training on lean tools emphasizing collaboration in the past year?", "count")
                ]
            },
            "Prácticas Sostenibles y Éticas": {
                "Español": [
                    ("¿Qué porcentaje de iniciativas lean implementadas en los últimos 12 meses redujo el consumo de recursos (por ejemplo, energía, materiales)?", "percentage"),
                    ("¿Qué porcentaje de proveedores principales fueron auditados en el último año para verificar estándares laborales y ambientales?", "percentage"),
                    ("¿Cuántos empleados participaron en proyectos de sostenibilidad con impacto comunitario o laboral en los últimos 12 meses?", "count")
                ],
                "English": [
                    ("What percentage of lean initiatives implemented in the past 12 months reduced resource consumption (e.g., energy, materials)?", "percentage"),
                    ("What percentage of primary suppliers were audited in the past year to verify labor and environmental standards?", "percentage"),
                    ("How many employees participated in sustainability projects with community or workplace impact in the past 12 months?", "count")
                ]
            },
            "Bienestar y Equilibrio": {
                "Español": [
                    ("¿Qué porcentaje de empleados accedió a recursos de bienestar (por ejemplo, asesoramiento, horarios flexibles) en los últimos 12 meses?", "percentage"),
                    ("¿Con qué frecuencia (en meses) se realizan encuestas o revisiones para evaluar el agotamiento o la fatiga de los empleados?", "frequency"),
                    ("¿Cuántos casos de desafíos personales o profesionales reportados por empleados fueron abordados con planes de acción documentados en el último año?", "count")
                ],
                "English": [
                    ("What percentage of employees accessed well-being resources (e.g., counseling, flexible schedules) in the past 12 months?", "percentage"),
                    ("How frequently (in months) are surveys or check-ins conducted to assess employee burnout or fatigue?", "frequency"),
                    ("How many reported employee personal or professional challenges were addressed with documented action plans in the past year?", "count")
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
        st.markdown('<div class="stepper">', unsafe_allow_html=True)
        for i, cat in enumerate(categories):
            status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
            if st.button(f"{i+1}", key=f"step_{i}", help=f"Ir a {cat} / Go to {cat}"):
                st.session_state.current_category = i
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

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
            for idx, (q, q_type) in enumerate(questions[category][LANG]):
                with st.container():
                    st.markdown(f"**{q}**")
                    options = [0, 25, 50, 75, 100]
                    score = st.radio(
                        "",
                        options,
                        format_func=lambda x: f"{x} - {labels[q_type][LANG][options.index(x)]}",
                        key=f"{category}_{idx}",
                        horizontal=True,
                        help="Selecciona una respuesta basada en datos verificables." if LANG == "Español" else
                             "Select a response based on verifiable data."
                    )
                    st.session_state.responses[category][idx] = score
            st.markdown('</div>', unsafe_allow_html=True)

        # Sticky navigation buttons
        with st.container():
            st.markdown('<div class="sticky-nav">', unsafe_allow_html=True)
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
                if category_index < len(categories) - 1:
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
                else:
                    if st.button("Enviar Auditoría / Submit Audit"):
                        if all(all(score != 0 for score in scores) for scores in st.session_state.responses.values()):
                            with st.spinner("Generando tu informe... / Generating your report..."):
                                st.session_state.current_category = len(categories)
                                st.rerun()
                        else:
                            st.error(
                                "Por favor, responde todas las preguntas en todas las categorías." if LANG == "Español" else
                                "Please answer all questions in all categories."
                            )
            st.markdown('</div>', unsafe_allow_html=True)

        # Grading matrix function
        def get_grade(score):
            if score >= 85:
                return (
                    "Excelente" if LANG == "Español" else "Excellent",
                    "Tu lugar de trabajo demuestra prácticas sobresalientes. ¡Continúa fortaleciendo estas áreas!" if LANG == "Español" else
                    "Your workplace demonstrates outstanding practices. Continue strengthening these areas!",
                    "grade-excellent"
                )
            elif score >= 70:
                return (
                    "Bueno" if LANG == "Español" else "Good",
                    "Tu lugar de trabajo tiene fortalezas, pero requiere mejoras específicas para alcanzar la excelencia." if LANG == "Español" else
                    "Your workplace has strengths but requires specific improvements to achieve excellence.",
                    "grade-good"
                )
            elif score >= 50:
                return (
                    "Necesita Mejora" if LANG == "Español" else "Needs Improvement",
                    "Se identificaron debilidades moderadas. Prioriza acciones correctivas en áreas críticas." if LANG == "Español" else
                    "Moderate weaknesses identified. Prioritize corrective actions in critical areas.",
                    "grade-needs-improvement"
                )
            else:
                return (
                    "Crítico" if LANG == "Español" else "Critical",
                    "Existen problemas significativos que requieren intervención urgente. Considera apoyo externo." if LANG == "Español" else
                    "Significant issues exist requiring urgent intervention. Consider external support.",
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
                '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por tu compromiso con un lugar de trabajo ético!</div>' if LANG == "Español" else
                '<div class="badge">🏆 Audit Completed! Thank you for your commitment to an ethical workplace!</div>', 
                unsafe_allow_html=True
            )
            
            # Calculate scores
            results = {cat: sum(scores) / len(scores) for cat, scores in st.session_state.responses.items()}
            df = pd.DataFrame.from_dict(results, orient="index", columns=["Puntuación" if LANG == "Español" else "Score"])
            df["Porcentaje" if LANG == "Español" else "Percent"] = df["Puntuación" if LANG == "Español" else "Score"]
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
                "Pregunta" if LANG == "Español" else "Question": [q for q, _ in questions[selected_category][LANG]],
                "Puntuación" if LANG == "Español" else "Score": st.session_state.responses[selected_category]
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
                    0: "Establece un sistema formal para rastrear e implementar sugerencias de empleados con métricas claras.",
                    1: "Aumenta las horas de capacitación profesional, asegurando acceso equitativo para todos los empleados.",
                    2: "Asigna presupuestos a más iniciativas lideradas por empleados para fomentar la innovación.",
                    3: "Programa foros mensuales para retroalimentación directa entre empleados y gerencia."
                },
                "Liderazgo Ético": {
                    0: "Implementa boletines mensuales para comunicar decisiones de liderazgo de manera transparente.",
                    1: "Incluye a representantes de empleados en la revisión de todas las políticas laborales nuevas.",
                    2: "Crea un programa formal de reconocimiento para comportamientos éticos, con incentivos claros."
                },
                "Operaciones Centradas en las Personas": {
                    0: "Integra retroalimentación de empleados en cada revisión de procesos lean para eliminar redundancias.",
                    1: "Realiza auditorías trimestrales de prácticas operativas con enfoque en el bienestar.",
                    2: "Capacita a todos los empleados en herramientas lean, priorizando la colaboración."
                },
                "Prácticas Sostenibles y Éticas": {
                    0: "Lanza iniciativas lean específicas para reducir el consumo de recursos, con metas medibles.",
                    1: "Audita anualmente a todos los proveedores principales para garantizar estándares éticos.",
                    2: "Involucra a más empleados en proyectos de sostenibilidad con impacto comunitario."
                },
                "Bienestar y Equilibrio": {
                    0: "Amplía el acceso a recursos de bienestar, como asesoramiento y horarios flexibles.",
                    1: "Implementa encuestas mensuales para monitorear el agotamiento y actuar rápidamente.",
                    2: "Establece procesos formales para abordar desafíos reportados con planes de acción."
                }
            } if LANG == "Español" else {
                "Empowering Employees": {
                    0: "Establish a formal system to track and implement employee suggestions with clear metrics.",
                    1: "Increase professional training hours, ensuring equitable access for all employees.",
                    2: "Allocate budgets to more employee-led initiatives to foster innovation.",
                    3: "Schedule monthly forums for direct employee-management feedback."
                },
                "Ethical Leadership": {
                    0: "Implement monthly newsletters to transparently communicate leadership decisions.",
                    1: "Include employee representatives in reviewing all new workplace policies.",
                    2: "Create a formal recognition program for ethical behavior with clear incentives."
                },
                "Human-Centered Operations": {
                    0: "Integrate employee feedback into every lean process review to eliminate redundancies.",
                    1: "Conduct quarterly audits of operational practices focusing on well-being.",
                    2: "Train all employees on lean tools, prioritizing collaboration."
                },
                "Sustainable and Ethical Practices": {
                    0: "Launch specific lean initiatives to reduce resource consumption with measurable goals.",
                    1: "Audit all primary suppliers annually to ensure ethical standards.",
                    2: "Engage more employees in sustainability projects with community impact."
                },
                "Well-Being and Balance": {
                    0: "Expand access to well-being resources, such as counseling and flexible schedules.",
                    1: "Implement monthly surveys to monitor burnout and act swiftly.",
                    2: "Establish formal processes to address reported challenges with action plans."
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
                    "Los resultados de tu auditoría indican oportunidades para optimizar el lugar de trabajo. LEAN 2.0 Institute ofrece consultoría especializada para directivos y HR, transformando tu entorno laboral en uno ético y eficiente." if LANG == "Español" else
                    "Your audit results indicate opportunities to optimize the workplace. LEAN 2.0 Institute offers specialized consulting for directors and HR, transforming your workplace into an ethical and efficient environment."
                )
                if df["Porcentaje" if LANG == "Español" else "Percent"].min() < 70:
                    low_categories = df[df["Porcentaje" if LANG == "Español" else "Percent"] < 70].index.tolist()
                    services = {
                        "Empoderamiento de Empleados": "Programas de Compromiso y Liderazgo de Empleados",
                        "Liderazgo Ético": "Capacitación en Liderazgo Ético y Gobernanza",
                        "Operaciones Centradas en las Personas": "Optimización de Procesos con Enfoque Humano",
                        "Prácticas Sostenibles y Éticas": "Consultoría en Sostenibilidad y Ética Empresarial",
                        "Bienestar y Equilibrio": "Estrategias de Bienestar Organizacional"
                    } if LANG == "Español" else {
                        "Empowering Employees": "Employee Engagement and Leadership Programs",
                        "Ethical Leadership": "Ethical Leadership and Governance Training",
                        "Human-Centered Operations": "Process Optimization with Human Focus",
                        "Sustainable and Ethical Practices": "Sustainability and Business Ethics Consulting",
                        "Well-Being and Balance": "Organizational Well-Being Strategies"
                    }
                    ad_text.append(
                        f"Las áreas clave para mejorar incluyen {', '.join(low_categories)}. LEAN 2.0 Institute se especializa en: {', '.join([services[cat] for cat in low_categories])}." if LANG == "Español" else
                        f"Key areas for improvement include {', '.join(low_categories)}. LEAN 2.0 Institute specializes in: {', '.join([services[cat] for cat in low_categories])}."
                    )
            else:
                ad_text.append(
                    "¡Felicidades por un lugar de trabajo sobresaliente! Colabora con LEAN 2.0 Institute para mantener estas fortalezas y liderar con innovación." if LANG == "Español" else
                    "Congratulations on an outstanding workplace! Partner with LEAN 2.0 Institute to sustain these strengths and lead with innovation."
                )
            ad_text.append(
                'Contáctanos en <a href="https://lean2institute.mystrikingly.com/" target="_blank" class="download-link">https://lean2institute.mystrikingly.com/</a> o envíanos un correo a info@lean2institute.com para una consulta estratégica.' if LANG == "Español" else
                'Contact us at <a href="https://lean2institute.mystrikingly.com/" target="_blank" class="download-link">https://lean2institute.mystrikingly.com/</a> or email us at info@lean2institute.com for a strategic consultation.'
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
            with col1:
                try:
                    pdf = FPDF()
                    pdf.set_margins(15, 15, 15)
                    if os.path.exists(font_path):
                        pdf.add_font('DejaVu', '', font_path, uni=True)
                        font_name = 'DejaVu'
                    else:
                        font_name = 'Arial'
                        st.warning(
                            "Fuente 'DejaVuSans.ttf' no encontrada. Usando Arial como respaldo. Para una mejor renderización, descárgala desde https://dejavu-fonts.github.io/ y colócala en el directorio del proyecto." if LANG == "Español" else
                            "Font 'DejaVuSans.ttf' not found. Using Arial as fallback. For better rendering, download it from https://dejavu-fonts.github.io/ and place it in the project directory."
                        )
                    pdf.add_page()
                    
                    # Overall Grade
                    pdf.set_font(font_name, 'B', 12)
                    pdf.set_text_color(51)
                    pdf.multi_cell(0, 10, f"Calificación General del Lugar de Trabajo: {grade} ({overall_score:.1f}%)" if LANG == "Español" else
                                        f"Overall Workplace Grade: {grade} ({overall_score:.1f}%)")
                    pdf.set_font(font_name, '', 12)
                    pdf.multi_cell(0, 10, grade_description)
                    pdf.ln(5)
                    
                    # Audit Results
                    pdf.set_font(font_name, 'B', 12)
                    pdf.multi_cell(0, 10, "Resultados de la Auditoría" if LANG == "Español" else "Audit Results")
                    pdf.set_font(font_name, '', 12)
                    pdf.ln(5)
                    for cat, row in df.iterrows():
                        pdf.multi_cell(0, 10, f"{cat}: {row['Porcentaje' if LANG == 'Español' else 'Percent']:.1f}% (Prioridad: {row['Prioridad' if LANG == 'Español' else 'Priority']})")
                    
                    # Action Plan
                    pdf.add_page()
                    pdf.set_font(font_name, 'B', 12)
                    pdf.multi_cell(0, 10, "Plan de Acción" if LANG == "Español" else "Action Plan")
                    pdf.set_font(font_name, '', 12)
                    pdf.ln(5)
                    for cat in categories:
                        if df.loc[cat, "Porcentaje" if LANG == "Español" else "Percent"] < 70:
                            pdf.set_font(font_name, 'B', 12)
                            pdf.multi_cell(0, 10, cat)
                            pdf.set_font(font_name, '', 12)
                            for idx, score in enumerate(st.session_state.responses[cat]):
                                if score < 70:
                                    question = questions[cat][LANG][idx][0]
                                    rec = recommendations[cat][idx]
                                    pdf.multi_cell(0, 10, f"- {question}: {rec}")
                            pdf.ln(5)
                    
                    # LEAN 2.0 Institute Advertisement
                    pdf.add_page()
                    pdf.set_font(font_name, 'B', 12)
                    pdf.multi_cell(0, 10, "Asóciate con LEAN 2.0 Institute" if LANG == "Español" else "Partner with LEAN 2.0 Institute")
                    pdf.set_font(font_name, '', 12)
                    pdf.ln(5)
                    for text in ad_text:
                        pdf.multi_cell(0, 10, text.replace('<a href="https://lean2institute.mystrikingly.com/" target="_blank" class="download-link">https://lean2institute.mystrikingly.com/</a>', 'https://lean2institute.mystrikingly.com/'))
                    
                    # Certificate
                    pdf.add_page()
                    pdf.set_font(font_name, 'B', 16)
                    pdf.set_text_color(40, 167, 69)
                    pdf.multi_cell(0, 10, "Certificado de Finalización" if LANG == "Español" else "Certificate of Completion", align="C")
                    pdf.ln(10)
                    pdf.set_font(font_name, '', 12)
                    pdf.set_text_color(51)
                    pdf.multi_cell(
                        0, 10, 
                        "¡Felicidades por completar la Auditoría del Lugar de Trabajo Ético! Tus respuestas están ayudando a construir un entorno laboral ético y sostenible. Contáctanos en https://lean2institute.mystrikingly.com/ para apoyo estratégico." if LANG == "Español" else 
                        "Congratulations on completing the Ethical Workplace Audit! Your responses are helping build an ethical and sustainable workplace. Contact us at https://lean2institute.mystrikingly.com/ for strategic support."
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
                        "Por favor, asegura que el archivo de fuente 'DejaVuSans.ttf' esté en el directorio del proyecto o usa Arial como respaldo. Descárgala desde https://dejavu-fonts.github.io/ si es necesario." if LANG == "Español" else
                        "Please ensure the 'DejaVuSans.ttf' font file is in the project directory or use Arial as fallback. Download it from https://dejavu-fonts.github.io/ if needed.",
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

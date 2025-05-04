import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import base64
import io
import numpy as np
import textwrap

# Set page configuration
st.set_page_config(page_title="Auditoría Ética de Lugar de Trabajo Lean", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for professional, accessible, and responsive UI
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
        :root {
            --primary: #1E88E5;
            --secondary: #43A047;
            --error: #D32F2F;
            --background: #F5F7FA;
            --surface: #FFFFFF;
            --text: #212121;
            --text-secondary: #757575;
        }
        body {
            font-family: 'Inter', sans-serif;
            background-color: var(--background);
            color: var(--text);
            line-height: 1.6;
        }
        .main-container {
            background-color: var(--surface);
            padding: 2rem;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
            max-width: 1200px;
            margin: 2rem auto;
            transition: transform 0.3s ease;
        }
        .stButton>button {
            background-color: var(--primary);
            color: white;
            border-radius: 8px;
            padding: 0.75rem 1.5rem;
            font-weight: 600;
            border: none;
            transition: background-color 0.3s ease, transform 0.2s ease, box-shadow 0.2s ease;
        }
        .stButton>button:hover {
            background-color: #1565C0;
            transform: translateY(-2px);
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
        }
        .stButton>button:focus {
            outline: 3px solid #90CAF9;
            outline-offset: 2px;
        }
        .stButton>button:disabled {
            background-color: #B0BEC5;
            cursor: not-allowed;
        }
        .stRadio>label {
            background-color: #E3F2FD;
            padding: 0.75rem;
            border-radius: 8px;
            margin: 0.5rem 0;
            font-size: 1rem;
            transition: background-color 0.3s ease;
        }
        .stRadio>label:hover {
            background-color: #BBDEFB;
        }
        .required {
            color: var(--error);
            font-weight: 600;
            margin-left: 0.25rem;
        }
        .header {
            color: var(--primary);
            font-size: 2.25rem;
            text-align: center;
            margin-bottom: 1.5rem;
            font-weight: 700;
        }
        .subheader {
            color: var(--text);
            font-size: 1.5rem;
            margin: 1.5rem 0;
            text-align: center;
            font-weight: 600;
        }
        .sidebar .sidebar-content {
            background-color: var(--surface);
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
        }
        .download-link {
            color: var(--primary);
            font-weight: 600;
            text-decoration: none;
            font-size: 1rem;
            display: inline-flex;
            align-items: center;
            transition: color 0.3s ease;
        }
        .download-link:hover {
            color: #1565C0;
            text-decoration: underline;
        }
        .download-link::before {
            content: '📄';
            margin-right: 0.5rem;
        }
        .motivation {
            color: var(--secondary);
            font-size: 1.1rem;
            text-align: center;
            margin: 1.5rem 0;
            font-style: italic;
            background-color: #E8F5E9;
            padding: 0.75rem;
            border-radius: 8px;
        }
        .badge {
            background-color: var(--secondary);
            color: white;
            padding: 0.75rem 1.5rem;
            border-radius: 24px;
            font-size: 1.1rem;
            text-align: center;
            margin: 1.5rem auto;
            display: block;
            width: fit-content;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        }
        .insights {
            background-color: #E3F2FD;
            padding: 1.5rem;
            border-radius: 12px;
            margin: 1.5rem 0;
            border-left: 4px solid var(--primary);
        }
        .grade {
            font-size: 1.5rem;
            font-weight: 700;
            text-align: center;
            margin: 1.5rem 0;
            padding: 0.75rem;
            border-radius: 8px;
        }
        .grade-excellent { background-color: var(--secondary); color: white; }
        .grade-good { background-color: #FFD54F; color: #212121; }
        .grade-needs-improvement { background-color: var(--error); color: white; }
        .grade-critical { background-color: #B71C1C; color: white; }
        .progress-bar {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin: 1.5rem 0;
            background-color: #ECEFF1;
            padding: 0.5rem;
            border-radius: 12px;
        }
        .progress-step {
            flex: 1;
            text-align: center;
            padding: 0.5rem;
            border-radius: 8px;
            cursor: pointer;
            transition: background-color 0.3s ease, color 0.3s ease;
            font-size: 0.9rem;
            font-weight: 600;
        }
        .progress-step.active {
            background-color: var(--primary);
            color: white;
        }
        .progress-step.completed {
            background-color: var(--secondary);
            color: white;
        }
        .progress-step:hover {
            background-color: #BBDEFB;
        }
        .progress-completion {
            background-color: var(--secondary);
            height: 8px;
            border-radius: 4px;
            transition: width 0.3s ease;
        }
        .card {
            background-color: #FAFAFA;
            padding: 1.5rem;
            border-radius: 12px;
            margin: 1rem 0;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.08);
            transition: transform 0.3s ease;
        }
        .card:hover {
            transform: translateY(-4px);
        }
        .sticky-nav {
            position: sticky;
            bottom: 1rem;
            background-color: var(--surface);
            padding: 1rem;
            border-radius: 12px;
            box-shadow: 0 -4px 16px rgba(0, 0, 0, 0.1);
            z-index: 1000;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .tooltip {
            position: relative;
            display: inline-block;
        }
        .tooltip .tooltiptext {
            visibility: hidden;
            width: 200px;
            background-color: #424242;
            color: white;
            text-align: center;
            border-radius: 6px;
            padding: 0.5rem;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -100px;
            opacity: 0;
            transition: opacity 0.3s;
        }
        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
        @media (max-width: 768px) {
            .main-container {
                padding: 1rem;
                margin: 1rem;
            }
            .header {
                font-size: 1.75rem;
            }
            .subheader {
                font-size: 1.25rem;
            }
            .stButton>button {
                padding: 0.5rem 1rem;
                font-size: 0.9rem;
            }
            .stRadio>label {
                font-size: 0.9rem;
                padding: 0.5rem;
            }
            .progress-bar {
                flex-direction: column;
                gap: 0.5rem;
            }
            .progress-step {
                font-size: 0.8rem;
                padding: 0.5rem;
            }
            .sticky-nav {
                flex-direction: column;
                gap: 0.5rem;
            }
        }
        [role="radiogroup"] {
            margin: 0.5rem 0;
        }
        [role="radio"] {
            cursor: pointer;
        }
        [role="radio"]:focus {
            outline: 3px solid #90CAF9;
            outline-offset: 2px;
        }
        .sr-only {
            position: absolute;
            width: 1px;
            height: 1px;
            padding: 0;
            margin: -1px;
            overflow: hidden;
            clip: rect(0, 0, 0, 0);
            border: 0;
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

# Sidebar navigation
with st.sidebar:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.selectbox(
        "Idioma / Language",
        ["Español", "English"],
        key="language_select",
        help="Selecciona tu idioma preferido / Select your preferred language"
    )
    st.markdown('<div class="subheader">Progreso</div>', unsafe_allow_html=True)
    categories = [
        "Empoderamiento de Empleados", "Liderazgo Ético", "Operaciones Centradas en las Personas",
        "Prácticas Sostenibles y Éticas", "Bienestar y Equilibrio"
    ] if st.session_state.language == "Español" else [
        "Empowering Employees", "Ethical Leadership", "Human-Centered Operations",
        "Sustainable and Ethical Practices", "Well-Being and Balance"
    ]
    for i, cat in enumerate(categories):
        status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
        if st.button(f"{cat}", key=f"nav_{i}", help=f"Ir a {cat} / Go to {cat}"):
            st.session_state.current_category = i
            st.session_state.show_intro = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# Reset session state on language change
if st.session_state.language != st.session_state.prev_language:
    st.session_state.current_category = 0
    st.session_state.responses = {}
    st.session_state.prev_language = st.session_state.language
    st.session_state.show_intro = True

# Introductory modal
if st.session_state.show_intro:
    with st.container():
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        st.markdown('<div class="header">¡Bienvenido a la Auditoría! / Welcome to the Audit!</div>', unsafe_allow_html=True)
        with st.expander("", expanded=True):
            st.markdown(
                """
                Esta auditoría está diseñada para directivos y profesionales de Recursos Humanos para evaluar de manera objetiva el entorno laboral. Responde preguntas en 5 categorías (5–10 minutos) con datos específicos y ejemplos verificables. Tus respuestas son confidenciales y generarán un informe detallado con recomendaciones accionables. Al completar la auditoría, contacta a LEAN 2.0 Institute en <a href="https://lean2institute.mystrikingly.com/" target="_blank">https://lean2institute.mystrikingly.com/</a> para consultas personalizadas.
                
                **Pasos**:
                1. Responde las preguntas de cada categoría.
                2. Revisa y descarga tu informe.
                
                ¡Empecemos!
                """
                if st.session_state.language == "Español" else
                """
                This audit is designed for directors and HR professionals to objectively assess the workplace environment. Answer questions across 5 categories (5–10 minutes) with specific data and verifiable examples. Your responses are confidential and will generate a detailed report with actionable recommendations. Upon completion, contact LEAN 2.0 Institute at <a href="https://lean2institute.mystrikingly.com/" target="_blank">https://lean2institute.mystrikingly.com/</a> for personalized consultation.
                
                **Steps**:
                1. Answer questions for each category.
                2. Review and download your report.
                
                Let’s get started!
                """
            )
            if st.button("Iniciar Auditoría / Start Audit", use_container_width=True):
                st.session_state.show_intro = False
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# Main content
if not st.session_state.show_intro:
    with st.container():
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        st.markdown(
            '<div class="header">¡Evalúa y Mejora tu Lugar de Trabajo!</div>' if st.session_state.language == "Español" else 
            '<div class="header">Assess and Enhance Your Workplace!</div>', 
            unsafe_allow_html=True
        )

        # Likert scale labels
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

        # Audit questions
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

        # Initialize responses
        if not st.session_state.responses or len(st.session_state.responses) != len(questions):
            st.session_state.responses = {cat: [None] * len(questions[cat][st.session_state.language]) for cat in questions}

        # Progress bar
        st.markdown('<div class="progress-bar">', unsafe_allow_html=True)
        for i, cat in enumerate(categories):
            status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
            st.markdown(
                f'<div class="progress-step {status}" onclick="Streamlit.setComponentValue({i})">{cat}</div>',
                unsafe_allow_html=True
            )
        st.markdown('</div>', unsafe_allow_html=True)

        # Progress completion bar
        completed_questions = sum(len([s for s in scores if s is not None]) for scores in st.session_state.responses.values())
        total_questions = sum(len(scores) for scores in st.session_state.responses.values())
        completion_percentage = (completed_questions / total_questions) * 100 if total_questions > 0 else 0
        st.markdown(
            f"""
            <div style='background-color: #ECEFF1; border-radius: 4px; margin: 1rem 0;'>
                <div class='progress-completion' style='width: {completion_percentage}%;'></div>
            </div>
            <div class='motivation'>{completed_questions}/{total_questions} preguntas completadas ({completion_percentage:.1f}%)</div>
            """ if st.session_state.language == "Español" else
            f"""
            <div style='background-color: #ECEFF1; border-radius: 4px; margin: 1rem 0;'>
                <div class='progress-completion' style='width: {completion_percentage}%;'></div>
            </div>
            <div class='motivation'>{completed_questions}/{total_questions} questions completed ({completion_percentage:.1f}%)</div>
            """,
            unsafe_allow_html=True
        )

        # Check if audit is complete
        audit_complete = all(all(score is not None for score in scores) for scores in st.session_state.responses.values())

        # Category questions
        if not audit_complete:
            category_index = min(st.session_state.current_category, len(categories) - 1)
            category = categories[category_index]
            
            with st.container():
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
                for idx, (q, q_type) in enumerate(questions[category][st.session_state.language]):
                    with st.container():
                        is_unanswered = st.session_state.responses[category][idx] is None
                        st.markdown(
                            f"""
                            <div class="tooltip">
                                <strong>{q}</strong> {'<span class="required">*</span>' if is_unanswered else ''}
                                <span class="tooltiptext">Proporciona datos verificables para una evaluación precisa.</span>
                            </div>
                            """ if st.session_state.language == "Español" else
                            f"""
                            <div class="tooltip">
                                <strong>{q}</strong> {'<span class="required">*</span>' if is_unanswered else ''}
                                <span class="tooltiptext">Provide verifiable data for an accurate assessment.</span>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                        options = [0, 25, 50, 75, 100]
                        score = st.radio(
                            "",
                            options,
                            format_func=lambda x: f"{x}% - {labels[q_type][st.session_state.language][options.index(x)]}",
                            key=f"{category}_{idx}",
                            horizontal=True,
                            help="Selecciona una respuesta basada en datos verificables." if st.session_state.language == "Español" else
                                 "Select a response based on verifiable data."
                        )
                        st.session_state.responses[category][idx] = score
                st.markdown('</div>', unsafe_allow_html=True)

            # Sticky navigation
            with st.container():
                st.markdown('<div class="sticky-nav">', unsafe_allow_html=True)
                col1, col2 = st.columns([1, 1], gap="small")
                with col1:
                    if st.button(
                        "⬅ Anterior" if st.session_state.language == "Español" else "⬅ Previous",
                        disabled=category_index == 0,
                        use_container_width=True,
                        help="Volver a la categoría anterior" if st.session_state.language == "Español" else "Go to previous category"
                    ):
                        st.session_state.current_category = max(category_index - 1, 0)
                        st.rerun()
                with col2:
                    if category_index < len(categories) - 1:
                        if st.button(
                            "Siguiente ➡" if st.session_state.language == "Español" else "Next ➡",
                            disabled=category_index == len(categories) - 1,
                            use_container_width=True,
                            help="Avanzar a la siguiente categoría" if st.session_state.language == "Español" else "Go to next category"
                        ):
                            if all(score is not None for score in st.session_state.responses[category]):
                                st.session_state.current_category = min(category_index + 1, len(categories) - 1)
                                st.rerun()
                            else:
                                unanswered = [q for i, (q, _) in enumerate(questions[category][st.session_state.language]) if st.session_state.responses[category][i] is None]
                                st.error(
                                    f"Por favor, responde las siguientes preguntas: {', '.join([q[:50] + '...' if len(q) > 50 else q for q in unanswered])}" if st.session_state.language == "Español" else
                                    f"Please answer the following questions: {', '.join([q[:50] + '...' if len(q) > 50 else q for q in unanswered])}"
                                )
                                first_unanswered_idx = next((i for i, score in enumerate(st.session_state.responses[category]) if score is None), None)
                                if first_unanswered_idx is not None:
                                    st.markdown(
                                        f"""
                                        <script>
                                            document.getElementById('{category}_{first_unanswered_idx}').scrollIntoView({{behavior: 'smooth', block: 'center'}});
                                            document.getElementById('{category}_{first_unanswered_idx}').focus();
                                        </script>
                                        """,
                                        unsafe_allow_html=True
                                    )
                    else:
                        if st.button(
                            "Ver Resultados" if st.session_state.language == "Español" else "View Results",
                            use_container_width=True,
                            help="Ver el informe de la auditoría" if st.session_state.language == "Español" else "View audit report"
                        ):
                            if all(all(score is not None for score in scores) for scores in st.session_state.responses.values()):
                                st.session_state.current_category = len(categories)
                                st.rerun()
                            else:
                                unanswered_questions = []
                                for cat in categories:
                                    for i, (q, _) in enumerate(questions[cat][st.session_state.language]):
                                        if st.session_state.responses[cat][i] is None:
                                            unanswered_questions.append(f"{cat}: {q[:50] + '...' if len(q) > 50 else q}")
                                st.error(
                                    f"Por favor, responde todas las preguntas en todas las categorías. Preguntas faltantes: {', '.join(unanswered_questions)}" if st.session_state.language == "Español" else
                                    f"Please answer all questions in all categories. Missing questions: {', '.join(unanswered_questions)}"
                                )
                st.markdown('</div>', unsafe_allow_html=True)

        # Grading matrix
        def get_grade(score):
            if score >= 85:
                return (
                    "Excelente" if st.session_state.language == "Español" else "Excellent",
                    "Tu lugar de trabajo demuestra prácticas sobresalientes. ¡Continúa fortaleciendo estas áreas!" if st.session_state.language == "Español" else
                    "Your workplace demonstrates outstanding practices. Continue strengthening these areas!",
                    "grade-excellent"
                )
            elif score >= 70:
                return (
                    "Bueno" if st.session_state.language == "Español" else "Good",
                    "Tu lugar de trabajo tiene fortalezas, pero requiere mejoras específicas para alcanzar la excelencia." if st.session_state.language == "Español" else
                    "Your workplace has strengths but requires specific improvements to achieve excellence.",
                    "grade-good"
                )
            elif score >= 50:
                return (
                    "Necesita Mejora" if st.session_state.language == "Español" else "Needs Improvement",
                    "Se identificaron debilidades moderadas. Prioriza acciones correctivas en áreas críticas." if st.session_state.language == "Español" else
                    "Moderate weaknesses identified. Prioritize corrective actions in critical areas.",
                    "grade-needs-improvement"
                )
            else:
                return (
                    "Crítico" if st.session_state.language == "Español" else "Critical",
                    "Existen problemas significativos que requieren intervención urgente. Considera apoyo externo." if st.session_state.language == "Español" else
                    "Significant issues exist requiring urgent intervention. Consider external support.",
                    "grade-critical"
                )

        # Generate report
        if audit_complete:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(
                f'<div class="subheader">{"Tu Informe de Impacto en el Lugar de Trabajo" if st.session_state.language == "Español" else "Your Workplace Impact Report"}</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por tu compromiso con un lugar de trabajo ético!</div>' if st.session_state.language == "Español" else
                '<div class="badge">🏆 Audit Completed! Thank you for your commitment to an ethical workplace!</div>', 
                unsafe_allow_html=True
            )

            # Calculate scores
            results = {cat: sum(scores) / len(scores) for cat, scores in st.session_state.responses.items()}
            df = pd.DataFrame.from_dict(results, orient="index", columns=["Puntuación" if st.session_state.language == "Español" else "Score"])
            df["Porcentaje" if st.session_state.language == "Español" else "Percent"] = df["Puntuación" if st.session_state.language == "Español" else "Score"]
            df["Prioridad" if st.session_state.language == "Español" else "Priority"] = df["Porcentaje" if st.session_state.language == "Español" else "Percent"].apply(
                lambda x: "Alta" if x < 50 else "Media" if x < 70 else "Baja" if st.session_state.language == "Español" else
                "High" if x < 50 else "Medium" if x < 70 else "Low"
            )

            # Summary dashboard
            st.markdown('<div class="subheader">Resumen Ejecutivo</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                overall_score = df["Porcentaje" if st.session_state.language == "Español" else "Percent"].mean()
                grade, grade_description, grade_class = get_grade(overall_score)
                st.markdown(
                    f'<div class="grade {grade_class}">Calificación General: {grade} ({overall_score:.1f}%)</div>' if st.session_state.language == "Español" else
                    f'<div class="grade {grade_class}">Overall Grade: {grade} ({overall_score:.1f}%)</div>',
                    unsafe_allow_html=True
                )
                st.markdown(grade_description, unsafe_allow_html=True)
            with col2:
                st.metric(
                    "Categorías con Alta Prioridad" if st.session_state.language == "Español" else "High Priority Categories",
                    len(df[df["Porcentaje" if st.session_state.language == "Español" else "Percent"] < 50])
                )
            with col3:
                st.metric(
                    "Puntuación Promedio" if st.session_state.language == "Español" else "Average Score",
                    f"{overall_score:.1f}%"
                )

            # Color-coded dataframe
            def color_percent(val):
                color = '#D32F2F' if val < 50 else '#FFD54F' if val < 70 else '#43A047'
                return f'background-color: {color}; color: white;'
            
            st.markdown(
                "Puntuaciones por debajo del 50% (rojo) requieren acción urgente, 50–69% (amarillo) sugieren mejoras, y por encima del 70% (verde) indican fortalezas." if st.session_state.language == "Español" else
                "Scores below 50% (red) need urgent action, 50–69% (yellow) suggest improvement, and above 70% (green) indicate strengths.",
                unsafe_allow_html=True
            )
            styled_df = df.style.applymap(color_percent, subset=["Porcentaje" if st.session_state.language == "Español" else "Percent"]).format({"Porcentaje" if st.session_state.language == "Español" else "Percent": "{:.1f}%"})
            st.dataframe(styled_df, use_container_width=True)

            # Interactive bar chart
            fig = px.bar(
                df.reset_index(),
                y="index",
                x="Porcentaje" if st.session_state.language == "Español" else "Percent",
                orientation='h',
                title="Fortalezas y Oportunidades del Lugar de Trabajo" if st.session_state.language == "Español" else "Workplace Strengths and Opportunities",
                labels={"index": "Categoría" if st.session_state.language == "Español" else "Category", "Porcentaje" if st.session_state.language == "Español" else "Percent": "Puntuación (%)" if st.session_state.language == "Español" else "Score (%)"},
                color="Porcentaje" if st.session_state.language == "Español" else "Percent",
                color_continuous_scale=["#D32F2F", "#FFD54F", "#43A047"],
                range_x=[0, 100],
                height=400
            )
            fig.add_vline(x=70, line_dash="dash", line_color="blue", annotation_text="Objetivo (70%)" if st.session_state.language == "Español" else "Target (70%)", annotation_position="top")
            for i, row in df.iterrows():
                if row["Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70:
                    fig.add_annotation(
                        x=row["Porcentaje" if st.session_state.language == "Español" else "Percent"], y=i,
                        text="Prioridad" if st.session_state.language == "Español" else "Priority", showarrow=True, arrowhead=2, ax=20, ay=-30,
                        font=dict(color="red", size=12)
                    )
            fig.update_layout(
                showlegend=False,
                title_x=0.5,
                xaxis_title="Puntuación (%)" if st.session_state.language == "Español" else "Score (%)",
                yaxis_title="Categoría" if st.session_state.language == "Español" else "Category",
                coloraxis_showscale=False,
                margin=dict(l=150, r=50, t=100, b=50)
            )
            st.plotly_chart(fig, use_container_width=True)

            # Question-level breakdown (collapsible)
            with st.expander("Análisis Detallado: Perspectivas a Nivel de Pregunta" if st.session_state.language == "Español" else "Drill Down: Question-Level Insights"):
                selected_category = st.selectbox(
                    "Seleccionar Categoría para Explorar" if st.session_state.language == "Español" else "Select Category to Explore",
                    categories,
                    key="category_explore"
                )
                question_scores = pd.DataFrame({
                    "Pregunta" if st.session_state.language == "Español" else "Question": [q for q, _ in questions[selected_category][st.session_state.language]],
                    "Puntuación" if st.session_state.language == "Español" else "Score": st.session_state.responses[selected_category]
                })
                fig_questions = px.bar(
                    question_scores,
                    x="Puntuación" if st.session_state.language == "Español" else "Score",
                    y="Pregunta" if st.session_state.language == "Español" else "Question",
                    orientation='h',
                    title=f"Puntuaciones de Preguntas para {selected_category}" if st.session_state.language == "Español" else f"Question Scores for {selected_category}",
                    labels={"Puntuación" if st.session_state.language == "Español" else "Score": "Puntuación (%)" if st.session_state.language == "Español" else "Score (%)", "Pregunta" if st.session_state.language == "Español" else "Question": "Pregunta" if st.session_state.language == "Español" else "Question"},
                    color="Puntuación" if st.session_state.language == "Español" else "Score",
                    color_continuous_scale=["#D32F2F", "#FFD54F", "#43A047"],
                    range_x=[0, 100],
                    height=300 + len(question_scores) * 50
                )
                fig_questions.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    xaxis_title="Puntuación (%)" if st.session_state.language == "Español" else "Score (%)",
                    yaxis_title="Pregunta" if st.session_state.language == "Español" else "Question",
                    coloraxis_showscale=False,
                    margin=dict(l=150, r=50, t=100, b=50)
                )
                st.plotly_chart(fig_questions, use_container_width=True)

            # Actionable insights
            with st.expander("Perspectivas Accionables" if st.session_state.language == "Español" else "Actionable Insights"):
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
                } if st.session_state.language == "Español" else {
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
                    if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 50:
                        insights.append(
                            f"**{cat}** obtuvo {df.loc[cat, 'Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}% (Alta Prioridad). Enfócate en mejoras inmediatas." if st.session_state.language == "Español" else
                            f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (High Priority). Focus on immediate improvements."
                        )
                    elif df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70:
                        insights.append(
                            f"**{cat}** obtuvo {df.loc[cat, 'Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}% (Prioridad Media). Considera acciones específicas." if st.session_state.language == "Español" else
                            f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (Medium Priority). Consider targeted actions."
                        )
                if insights:
                    st.markdown(
                        "<div class='insights'>" + "<br>".join(insights) + "</div>",
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        "<div class='insights'>¡Todas las categorías obtuvieron más del 70%! Continúa manteniendo estas fortalezas.</div>" if st.session_state.language == "Español" else 
                        "<div class='insights'>All categories scored above 70%! Continue maintaining these strengths.</div>",
                        unsafe_allow_html=True
                    )

            # LEAN 2.0 Institute Advertisement
            st.markdown(
                "<div class='subheader'>Optimiza tu Lugar de Trabajo con LEAN 2.0 Institute</div>" if st.session_state.language == "Español" else 
                "<div class='subheader'>Optimize Your Workplace with LEAN 2.0 Institute</div>",
                unsafe_allow_html=True
            )
            ad_text = []
            if overall_score < 85:
                ad_text.append(
                    "Los resultados de tu auditoría indican oportunidades para optimizar el lugar de trabajo. LEAN 2.0 Institute ofrece consultoría especializada para directivos y HR, transformando tu entorno laboral en uno ético y eficiente." if st.session_state.language == "Español" else
                    "Your audit results indicate opportunities to optimize the workplace. LEAN 2.0 Institute offers specialized consulting for directors and HR, transforming your workplace into an ethical and efficient environment."
                )
                if df["Porcentaje" if st.session_state.language == "Español" else "Percent"].min() < 70:
                    low_categories = df[df["Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70].index.tolist()
                    services = {
                        "Empoderamiento de Empleados": "Programas de Compromiso y Liderazgo de Empleados",
                        "Liderazgo Ético": "Capacitación en Liderazgo Ético y Gobernanza",
                        "Operaciones Centradas en las Personas": "Optimización de Procesos con Enfoque Humano",
                        "Prácticas Sostenibles y Éticas": "Consultoría en Sostenibilidad y Ética Empresarial",
                        "Bienestar y Equilibrio": "Estrategias de Bienestar Organizacional"
                    } if st.session_state.language == "Español" else {
                        "Empowering Employees": "Employee Engagement and Leadership Programs",
                        "Ethical Leadership": "Ethical Leadership and Governance Training",
                        "Human-Centered Operations": "Process Optimization with Human Focus",
                        "Sustainable and Ethical Practices": "Sustainability and Business Ethics Consulting",
                        "Well-Being and Balance": "Organizational Well-Being Strategies"
                    }
                    ad_text.append(
                        f"Las áreas clave para mejorar incluyen {', '.join(low_categories)}. LEAN 2.0 Institute se especializa en: {', '.join([services[cat] for cat in low_categories])}." if st.session_state.language == "Español" else
                        f"Key areas for improvement include {', '.join(low_categories)}. LEAN 2.0 Institute specializes in: {', '.join([services[cat] for cat in low_categories])}."
                    )
            else:
                ad_text.append(
                    "¡Felicidades por un lugar de trabajo sobresaliente! Colabora con LEAN 2.0 Institute para mantener estas fortalezas y liderar con innovación." if st.session_state.language == "Español" else
                    "Congratulations on an outstanding workplace! Partner with LEAN 2.0 Institute to sustain these strengths and lead with innovation."
                )
            ad_text.append(
                "Contáctanos en https://lean2institute.mystrikingly.com/ o envíanos un correo a info@lean2institute.com para una consulta estratégica." if st.session_state.language == "Español" else
                "Contact us at https://lean2institute.mystrikingly.com/ or email us at info@lean2institute.com for a strategic consultation."
            )
            st.markdown("<div class='insights'>" + "<br>".join(ad_text) + "</div>", unsafe_allow_html=True)

            # Download buttons
            st.markdown(
                '<div class="subheader">Descarga tu Informe</div>' if st.session_state.language == "Español" else
                '<div class="subheader">Download Your Report</div>',
                unsafe_allow_html=True
            )
            col1, col2 = st.columns(2)

            # PDF Report
            with col1:
                with st.spinner("Generando PDF..." if st.session_state.language == "Español" else "Generating PDF..."):
                    try:
                        pdf = FPDF()
                        pdf.set_margins(20, 20, 20)
                        pdf.add_page()
                        font_name = 'Helvetica'
                        pdf.set_font(font_name, 'B', 14)
                        pdf.set_text_color(30, 136, 229)
                        pdf.cell(0, 10, "Informe de Auditoría del Lugar de Trabajo Ético" if st.session_state.language == "Español" else "Ethical Workplace Audit Report", ln=True, align="C")
                        pdf.ln(10)

                        # Overall Grade
                        pdf.set_font(font_name, 'B', 11)
                        pdf.set_text_color(51)
                        wrapped_grade = textwrap.wrap(f"Calificación General del Lugar de Trabajo: {grade} ({overall_score:.1f}%)" if st.session_state.language == "Español" else
                                                     f"Overall Workplace Grade: {grade} ({overall_score:.1f}%)", width=80)
                        for line in wrapped_grade:
                            pdf.multi_cell(0, 8, line)
                        pdf.set_font(font_name, '', 10)
                        wrapped_description = textwrap.wrap(grade_description, width=80)
                        for line in wrapped_description:
                            pdf.multi_cell(0, 8, line)
                        pdf.ln(8)

                        # Audit Results
                        pdf.set_font(font_name, 'B', 11)
                        pdf.multi_cell(0, 8, "Resultados de la Auditoría" if st.session_state.language == "Español" else "Audit Results")
                        pdf.set_font(font_name, '', 10)
                        pdf.ln(5)
                        for cat, row in df.iterrows():
                            result_text = f"{cat}: {row['Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}% (Prioridad: {row['Prioridad' if st.session_state.language == 'Español' else 'Priority']})"
                            wrapped_result = textwrap.wrap(result_text, width=80)
                            for line in wrapped_result:
                                pdf.multi_cell(0, 8, line)
                        pdf.ln(8)

                        # Action Plan
                        pdf.add_page()
                        pdf.set_font(font_name, 'B', 11)
                        pdf.multi_cell(0, 8, "Plan de Acción" if st.session_state.language == "Español" else "Action Plan")
                        pdf.set_font(font_name, '', 10)
                        pdf.ln(5)
                        for cat in categories:
                            if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70:
                                pdf.set_font(font_name, 'B', 11)
                                wrapped_cat = textwrap.wrap(cat, width=80)
                                for line in wrapped_cat:
                                    pdf.multi_cell(0, 8, line)
                                pdf.set_font(font_name, '', 10)
                                for idx, score in enumerate(st.session_state.responses[cat]):
                                    if score < 70:
                                        question = questions[cat][st.session_state.language][idx][0]
                                        rec = recommendations[cat][idx]
                                        action_text = f"- {question}: {rec}"
                                        wrapped_action = textwrap.wrap(action_text, width=75)
                                        for line in wrapped_action:
                                            pdf.multi_cell(0, 8, line)
                                pdf.ln(5)

                        # LEAN 2.0 Institute Advertisement
                        pdf.add_page()
                        pdf.set_font(font_name, 'B', 11)
                        pdf.multi_cell(0, 8, "Asóciate con LEAN 2.0 Institute" if st.session_state.language == "Español" else "Partner with LEAN 2.0 Institute")
                        pdf.set_font(font_name, '', 10)
                        pdf.ln(5)
                        for text in ad_text:
                            wrapped_text = textwrap.wrap(text, width=80)
                            for line in wrapped_text:
                                pdf.multi_cell(0, 8, line)

                        # Certificate
                        pdf.add_page()
                        pdf.set_font(font_name, 'B', 14)
                        pdf.set_text_color(67, 160, 71)
                        pdf.multi_cell(0, 10, "Certificado de Finalización" if st.session_state.language == "Español" else "Certificate of Completion", align="C")
                        pdf.ln(10)
                        pdf.set_font(font_name, '', 10)
                        pdf.set_text_color(51)
                        cert_text = (
                            "¡Felicidades por completar la Auditoría del Lugar de Trabajo Ético! Tus respuestas están ayudando a construir un entorno laboral ético y sostenible. Contáctanos en https://lean2institute.mystrikingly.com/ para apoyo estratégico." if st.session_state.language == "Español" else 
                            "Congratulations on completing the Ethical Workplace Audit! Your responses are helping build an ethical and sustainable workplace. Contact us at https://lean2institute.mystrikingly.com/ for strategic support."
                        )
                        wrapped_cert = textwrap.wrap(cert_text, width=80)
                        for line in wrapped_cert:
                            pdf.multi_cell(0, 8, line)

                        pdf_output = io.BytesIO()
                        pdf.output(pdf_output)
                        pdf_output.seek(0)
                        b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
                        href_pdf = (
                            f'<a href="data:application/pdf;base64,{b64_pdf}" download="informe_auditoria_lugar_trabajo_etico.pdf" class="download-link" role="button" aria-label="Descargar Informe PDF">Descargar Informe PDF y Plan de Acción</a>' if st.session_state.language == "Español" else 
                            f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_workplace_audit_report.pdf" class="download-link" role="button" aria-label="Download PDF Report">Download PDF Report & Action Plan</a>'
                        )
                        st.markdown(href_pdf, unsafe_allow_html=True)
                        pdf_output.close()
                    except Exception as e:
                        st.error(f"Error al generar el PDF: {str(e)}. Por favor, intenta de nuevo o contacta a soporte." if st.session_state.language == "Español" else
                                 f"Error generating PDF: {str(e)}. Please try again or contact support.")

            # Excel export
            with col2:
                with st.spinner("Generando Excel..." if st.session_state.language == "Español" else "Generating Excel..."):
                    try:
                        excel_output = io.BytesIO()
                        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                            df.to_excel(
                                writer, 
                                sheet_name='Resultados de la Auditoría' if st.session_state.language == "Español" else 'Audit Results', 
                                float_format="%.1f"
                            )
                            pd.DataFrame({"Puntuación General" if st.session_state.language == "Español" else "Overall Score": [overall_score], "Calificación" if st.session_state.language == "Español" else "Grade": [grade]}).to_excel(
                                writer, 
                                sheet_name='Resumen' if st.session_state.language == "Español" else 'Summary', 
                                index=False
                            )
                        excel_output.seek(0)
                        b64_excel = base64.b64encode(excel_output.getvalue()).decode()
                        href_excel = (
                            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="resultados_auditoria_lugar_trabajo_etico.xlsx" class="download-link" role="button" aria-label="Descargar Informe Excel">Descargar Informe Excel</a>' if st.session_state.language == "Español" else 
                            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_workplace_audit_results.xlsx" class="download-link" role="button" aria-label="Download Excel Report">Download Excel Report</a>'
                        )
                        st.markdown(href_excel, unsafe_allow_html=True)
                        excel_output.close()
                    except ImportError:
                        st.error("La exportación a Excel requiere 'xlsxwriter'. Por favor, instálalo usando `pip install xlsxwriter`." if st.session_state.language == "Español" else
                                 "Excel export requires 'xlsxwriter'. Please install it using `pip install xlsxwriter`.")
                    except Exception as e:
                        st.error(f"No se pudo generar el archivo Excel: {str(e)}" if st.session_state.language == "Español" else f"Failed to generate Excel file: {str(e)}")

            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

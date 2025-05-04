import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
import io
import numpy as np
import xlsxwriter

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
        .stSelectbox select {
            background-color: #E3F2FD;
            border-radius: 8px;
            padding: 0.5rem;
            font-size: 1rem;
            font-weight: 600;
            transition: background-color 0.3s ease;
        }
        .stSelectbox select:hover {
            background-color: #BBDEFB;
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
if 'language_changed' not in st.session_state:
    st.session_state.language_changed = False
if 'show_results' not in st.session_state:
    st.session_state.show_results = False

# Category mapping for bilingual support
category_mapping = {
    "Español": {
        "Empoderamiento de Empleados": "Empoderamiento de Empleados",
        "Liderazgo Ético": "Liderazgo Ético",
        "Operaciones Centradas en las Personas": "Operaciones Centradas en las Personas",
        "Prácticas Sostenibles y Éticas": "Prácticas Sostenibles y Éticas",
        "Bienestar y Equilibrio": "Bienestar y Equilibrio"
    },
    "English": {
        "Empowering Employees": "Empoderamiento de Empleados",
        "Ethical Leadership": "Liderazgo Ético",
        "Human-Centered Operations": "Operaciones Centradas en las Personas",
        "Sustainable and Ethical Practices": "Prácticas Sostenibles y Éticas",
        "Well-Being and Balance": "Bienestar y Equilibrio"
    }
}

# Audit questions
questions = {
    "Empoderamiento de Empleados": {
        "Español": [
            ("¿Qué porcentaje de sugerencias de empleados presentadas en los últimos 12 meses fueron implementadas con resultados documentados?", "percentage"),
            ("¿Cuántos empleados recibieron capacitación en habilidades profesionales en el último año?", "count"),
            ("En los últimos 12 meses, ¿cuántos empleados lideraron proyectos o iniciativas con presupuesto asignado?", "count"),
            ("¿Con qué frecuencia se realizan foros o reuniones formales para que los empleados compartan retroalimentación con la gerencia?", "frequency")
        ],
        "English": [
            ("What percentage of employee suggestions submitted in the past 12 months were implemented with documented outcomes?", "percentage"),
            ("How many employees received professional skills training in the past year?", "count"),
            ("In the past 12 months, how many employees led projects or initiatives with allocated budgets?", "count"),
            ("How frequently are formal forums or meetings held for employees to share feedback with management?", "frequency")
        ]
    },
    "Liderazgo Ético": {
        "Español": [
            ("¿Con qué frecuencia los líderes compartieron actualizaciones escritas sobre decisiones que afectan a los empleados en los últimos 12 meses?", "frequency"),
            ("¿Qué porcentaje de políticas laborales nuevas o revisadas en el último año incluyó consulta formal con empleados?", "percentage"),
            ("¿Cuántos casos de comportamiento ético destacado fueron reconocidos formalmente en los últimos 12 meses?", "count")
        ],
        "English": [
            ("How frequently did leaders share written updates on decisions affecting employees in the past 12 months?", "frequency"),
            ("What percentage of new or revised workplace policies in the past year included formal employee consultation?", "percentage"),
            ("How many instances of exemplary ethical behavior were formally recognized in the past 12 months?", "count")
        ]
    },
    "Operaciones Centradas en las Personas": {
        "Español": [
            ("¿Qué porcentaje de procesos lean revisados en los últimos 12 meses incorporó retroalimentación de empleados para reducir tareas redundantes?", "percentage"),
            ("¿Con qué frecuencia se auditan las prácticas operativas para evaluar su impacto en el bienestar de los empleados?", "frequency"),
            ("¿Cuántos empleados recibieron capacitación en herramientas lean con énfasis en colaboración en el último año?", "count")
        ],
        "English": [
            ("What percentage of lean processes revised in the past 12 months incorporated employee feedback to reduce redundant tasks?", "percentage"),
            ("How frequently are operational practices audited to assess their impact on employee well-being?", "frequency"),
            ("How many employees received training on lean tools emphasizing collaboration in the past year?", "count")
        ]
    },
    "Prácticas Sostenibles y Éticas": {
        "Español": [
            ("¿Qué porcentaje de iniciativas lean implementadas en los últimos 12 meses redujo el consumo de recursos?", "percentage"),
            ("¿Qué porcentaje de proveedores principales fueron auditados en el último año para verificar estándares laborales y ambientales?", "percentage"),
            ("¿Cuántos empleados participaron en proyectos de sostenibilidad con impacto comunitario o laboral en los últimos 12 meses?", "count")
        ],
        "English": [
            ("What percentage of lean initiatives implemented in the past 12 months reduced resource consumption?", "percentage"),
            ("What percentage of primary suppliers were audited in the past year to verify labor and environmental standards?", "percentage"),
            ("How many employees participated in sustainability projects with community or workplace impact in the past 12 months?", "count")
        ]
    },
    "Bienestar y Equilibrio": {
        "Español": [
            ("¿Qué porcentaje de empleados accedió a recursos de bienestar en los últimos 12 meses?", "percentage"),
            ("¿Con qué frecuencia se realizan encuestas o revisiones para evaluar el agotamiento o la fatiga de los empleados?", "frequency"),
            ("¿Cuántos casos de desafíos personales o profesionales reportados por empleados fueron abordados con planes de acción documentados en el último año?", "count")
        ],
        "English": [
            ("What percentage of employees accessed well-being resources in the past 12 months?", "percentage"),
            ("How frequently are surveys or check-ins conducted to assess employee burnout or fatigue?", "frequency"),
            ("How many reported employee personal or professional challenges were addressed with documented action plans in the past year?", "count")
        ]
    }
}

# Sidebar navigation
with st.sidebar:
    st.sidebar.image("assets/FOBO2.png", width=250)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    # Language selection with on_change callback
    def update_language():
        if st.session_state.language_select != st.session_state.language:
            st.session_state.language_changed = True
    
    st.selectbox(
        "Idioma / Language",
        ["Español", "English"],
        key="language_select",
        help="Selecciona tu idioma preferido / Select your preferred language",
        on_change=update_language
    )
    
    # Handle language change in main script
    if st.session_state.language_changed:
        st.session_state.language = st.session_state.language_select
        st.session_state.current_category = 0
        st.session_state.responses = {cat: [None] * len(questions[cat][st.session_state.language]) for cat in questions}
        st.session_state.prev_language = st.session_state.language
        st.session_state.show_intro = True
        st.session_state.show_results = False
        st.session_state.language_changed = False
        st.info("Cargando contenido en el nuevo idioma..." if st.session_state.language == "Español" else "Loading content in the new language...")
        st.rerun()
    
    st.markdown('<div class="subheader">Progreso</div>', unsafe_allow_html=True)
    display_categories = list(category_mapping[st.session_state.language].keys())
    for i, display_cat in enumerate(display_categories):
        status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
        if st.button(f"{display_cat}", key=f"nav_{i}", help=f"Ir a {display_cat} / Go to {display_cat}"):
            st.session_state.current_category = i
            st.session_state.show_intro = False
            st.session_state.show_results = False
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# Initialize responses if empty or mismatched
if not st.session_state.responses or len(st.session_state.responses) != len(questions):
    st.session_state.responses = {cat: [None] * len(questions[cat][st.session_state.language]) for cat in questions}

# Introductory modal
if st.session_state.show_intro:
    with st.container():
        st.markdown('<div class="main-container">', unsafe_allow_html=True)
        st.markdown('<div class="header">¡Bienvenido a LEAN 2.0 Institute! Evaluemos juntos tu entorno laboral 🤝/ Welcome to LEAN 2.0 Institute! Let’s assess together your work environment 🤝</div>', unsafe_allow_html=True)
        with st.expander("", expanded=True):
            st.markdown(
                """
                Esta evaluación está diseñada para ser completada por la gerencia en conjunto con Recursos Humanos, proporcionando una evaluación objetiva de tu entorno laboral. Responde preguntas en 5 categorías (5–10 minutos) con datos específicos y ejemplos verificables. Tus respuestas son confidenciales y generarán un informe detallado con recomendaciones accionables que podemos ayudarte a implementar. Al completar la evaluación, no dudes en ponerte en contacto con nosotros para consultas personalizadas a través de: ✉️Email: contacto@lean2institute.org 🌐 Website: https://lean2institute.mystrikingly.com/
                
                **Pasos**:
                1. Responde las preguntas de cada categoría.
                2. Revisa y descarga tu informe.
                
                ¡Empecemos!
                """
                if st.session_state.language == "Español" else
                """
                This assessment is designed to be completed by management and HR, to provide an objective evaluation of the work environment at your company. Answer questions across 5 categories (5–10 minutes) with specific data and verifiable examples. Your responses are confidential and will generate a detailed report with actionable recommendations that we can help you implement. Once you’ve completed the evaluation, feel free to reach out to us for personalized consultations at: ✉️Email: contacto@lean2institute.org 🌐 Website: https://lean2institute.mystrikingly.com/
                
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

        # Response options with descriptions as selectable options
        response_options = {
            "percentage": {
                "Español": {
                    "descriptions": [
                        "Ninguna sugerencia/proceso fue implementado.",
                        "Aproximadamente una cuarta parte fue implementada.",
                        "La mitad fue implementada.",
                        "Tres cuartas partes fueron implementadas.",
                        "Todas las sugerencias/procesos fueron implementados."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "English": {
                    "descriptions": [
                        "No suggestions/processes were implemented.",
                        "About one-quarter were implemented.",
                        "Half were implemented.",
                        "Three-quarters were implemented.",
                        "All suggestions/processes were implemented."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "tooltip": "Selecciona la descripción que mejor refleje la proporción de casos aplicados." if st.session_state.language == "Español" else
                          "Select the description that best reflects the proportion of cases applied."
            },
            "frequency": {
                "Español": {
                    "descriptions": [
                        "Esto nunca ocurre.",
                        "Ocurre muy pocas veces al año.",
                        "Ocurre varias veces al año.",
                        "Ocurre regularmente, casi siempre.",
                        "Ocurre en cada oportunidad."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "English": {
                    "descriptions": [
                        "This never occurs.",
                        "Occurs very few times a year.",
                        "Occurs several times a year.",
                        "Occurs regularly, almost always.",
                        "Occurs every time."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "tooltip": "Selecciona la descripción que mejor refleje la frecuencia de la práctica." if st.session_state.language == "Español" else
                          "Select the description that best reflects the frequency of the practice."
            },
            "count": {
                "Español": {
                    "descriptions": [
                        "Ningún empleado o caso (0%).",
                        "Menos de un cuarto de los empleados (1-25%).",
                        "Entre un cuarto y la mitad (25-50%).",
                        "Más de la mitad pero no la mayoría (50-75%).",
                        "Más del 75% de los empleados o casos."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "English": {
                    "descriptions": [
                        "No employees or cases (0%).",
                        "Less than a quarter of employees (1-25%).",
                        "Between a quarter and half (25-50%).",
                        "More than half but not most (50-75%).",
                        "Over 75% of employees or cases."
                    ],
                    "scores": [0, 25, 50, 75, 100]
                },
                "tooltip": "Selecciona la descripción que mejor refleje la cantidad de empleados o casos afectados." if st.session_state.language == "Español" else
                          "Select the description that best reflects the number of employees or cases affected."
            }
        }

        # Progress bar
        st.markdown('<div class="progress-bar">', unsafe_allow_html=True)
        for i, display_cat in enumerate(display_categories):
            status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
            st.markdown(
                f'<div class="progress-step {status}" onclick="Streamlit.setComponentValue({i})">{display_cat}</div>',
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
        if not st.session_state.show_results:
            category_index = min(st.session_state.current_category, len(display_categories) - 1)
            display_category = display_categories[category_index]
            category = category_mapping[st.session_state.language][display_category]
            
            with st.container():
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown(f'<div class="subheader">{display_category}</div>', unsafe_allow_html=True)
                
                # Response guide (no table, just instructions)
                with st.expander("Guía de Respuestas" if st.session_state.language == "Español" else "Response Guide", expanded=True):
                    st.markdown(
                        "Selecciona la descripción que mejor represente la situación para cada pregunta. Las opciones describen el grado, frecuencia o cantidad aplicable." if st.session_state.language == "Español" else
                        "Select the description that best represents the situation for each question. The options describe the degree, frequency, or quantity applicable."
                    )

                for idx, (q, q_type) in enumerate(questions[category][st.session_state.language]):
                    with st.container():
                        is_unanswered = st.session_state.responses[category][idx] is None
                        st.markdown(
                            f"""
                            <div class="tooltip">
                                <strong>{q}</strong> {'<span class="required">*</span>' if is_unanswered else ''}
                                <span class="tooltiptext">{response_options[q_type]['tooltip']}</span>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                        descriptions = response_options[q_type][st.session_state.language]["descriptions"]
                        scores = response_options[q_type][st.session_state.language]["scores"]
                        selected_description = st.radio(
                            "",
                            descriptions,
                            format_func=lambda x: x,  # Display descriptions as-is
                            key=f"{category}_{idx}",
                            horizontal=False,  # Vertical layout for readability due to longer descriptions
                            help=response_options[q_type]['tooltip']
                        )
                        # Map selected description to its corresponding score
                        score_idx = descriptions.index(selected_description)
                        st.session_state.responses[category][idx] = scores[score_idx]
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
                        st.session_state.show_results = False
                        st.rerun()
                with col2:
                    if category_index < len(display_categories) - 1:
                        if st.button(
                            "Siguiente ➡" if st.session_state.language == "Español" else "Next ➡",
                            disabled=category_index == len(display_categories) - 1,
                            use_container_width=True,
                            help="Avanzar a la siguiente categoría" if st.session_state.language == "Español" else "Go to next category"
                        ):
                            if all(score is not None for score in st.session_state.responses[category]):
                                st.session_state.current_category = min(category_index + 1, len(display_categories) - 1)
                                st.session_state.show_results = False
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
                            help="Ver el informe de la auditoría" if st.session_state.language == "Español" else "View audit report",
                            disabled=not audit_complete
                        ):
                            if audit_complete:
                                st.session_state.show_results = True
                                st.rerun()
                            else:
                                unanswered_questions = []
                                for cat in questions.keys():
                                    for i, (q, _) in enumerate(questions[cat][st.session_state.language]):
                                        if st.session_state.responses[cat][i] is None:
                                            display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                                            unanswered_questions.append(f"{display_cat}: {q[:50] + '...' if len(q) > 50 else q}")
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
        if st.session_state.show_results:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown(
                f'<div class="subheader">{"Tu Informe de Impacto en el Lugar de Trabajo" if st.session_state.language == "Español" else "Your Workplace Impact Report"}</div>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por tu compromiso con la construcción de un entorno laboral saludable, seguro y respetuoso para todas las personas!</div>' if st.session_state.language == "Español" else
                '<div class="badge">🏆 Audit Completed! Thank you for your commitment to fostering a healthy, safe, and respectful work environment for everyone!</div>', 
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
            df_display = df.copy()
            df_display.index = [next(k for k, v in category_mapping[st.session_state.language].items() if v == idx) for idx in df.index]
            fig = px.bar(
                df_display.reset_index(),
                y="index",
                x="Porcentaje" if st.session_state.language == "Español" else "Percent",
                orientation='h',
                title="Fortalezas y Oportunidades del Lugar de Trabajo" if st.session_state.language == "Español" else "Workplace Strengths and Opportunities",
                labels={
                    "index": "Categoría" if st.session_state.language == "Español" else "Category",
                    "Porcentaje" if st.session_state.language == "Español" else "Percent": "Puntuación (%)" if st.session_state.language == "Español" else "Score (%)"
                },
                color="Porcentaje" if st.session_state.language == "Español" else "Percent",
                color_continuous_scale=["#D32F2F", "#FFD54F", "#43A047"],
                range_x=[0, 100],
                height=400
            )
            fig.add_vline(x=70, line_dash="dash", line_color="blue", annotation_text="Objetivo (70%)" if st.session_state.language == "Español" else "Target (70%)", annotation_position="top")
            for i, row in df_display.iterrows():
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
                selected_display_category = st.selectbox(
                    "Seleccionar Categoría para Explorar" if st.session_state.language == "Español" else "Select Category to Explore",
                    display_categories,
                    key="category_explore"
                )
                selected_category = category_mapping[st.session_state.language][selected_display_category]
                question_scores = pd.DataFrame({
                    "Pregunta" if st.session_state.language == "Español" else "Question": [q for q, _ in questions[selected_category][st.session_state.language]],
                    "Puntuación" if st.session_state.language == "Español" else "Score": st.session_state.responses[selected_category]
                })
                fig_questions = px.bar(
                    question_scores,
                    x="Puntuación" if st.session_state.language == "Español" else "Score",
                    y="Pregunta" if st.session_state.language == "Español" else "Question",
                    orientation='h',
                    title=f"Puntuaciones de Preguntas para {selected_display_category}" if st.session_state.language == "Español" else f"Question Scores for {selected_display_category}",
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
                        1: "Aumenta las oportunidades de capacitación profesional para todos los empleados.",
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
                    "Empoderamiento de Empleados": {
                        0: "Establish a formal system to track and implement employee suggestions with clear metrics.",
                        1: "Increase professional training opportunities for all employees.",
                        2: "Allocate budgets to more employee-led initiatives to foster innovation.",
                        3: "Schedule monthly forums for direct employee-management feedback."
                    },
                    "Liderazgo Ético": {
                        0: "Implement monthly newsletters to transparently communicate leadership decisions.",
                        1: "Include employee representatives in reviewing all new workplace policies.",
                        2: "Create a formal recognition program for ethical behavior with clear incentives."
                    },
                    "Operaciones Centradas en las Personas": {
                        0: "Integrate employee feedback into every lean process review to eliminate redundancies.",
                        1: "Conduct quarterly audits of operational practices focusing on well-being.",
                        2: "Train all employees on lean tools, prioritizing collaboration."
                    },
                    "Prácticas Sostenibles y Éticas": {
                        0: "Launch specific lean initiatives to reduce resource consumption with measurable goals.",
                        1: "Audit all primary suppliers annually to ensure ethical standards.",
                        2: "Engage more employees in sustainability projects with community impact."
                    },
                    "Bienestar y Equilibrio": {
                        0: "Expand access to well-being resources, such as counseling and flexible schedules.",
                        1: "Implement monthly surveys to monitor burnout and act swiftly.",
                        2: "Establish formal processes to address reported challenges with action plans."
                    }
                }
                for cat in questions.keys():
                    display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                    if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 50:
                        insights.append(
                            f"**{display_cat}** obtuvo {df.loc[cat, 'Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}% (Alta Prioridad). Enfócate en mejoras inmediatas." if st.session_state.language == "Español" else
                            f"**{display_cat}** scored {df.loc[cat, 'Percent']:.1f}% (High Priority). Focus on immediate improvements."
                        )
                    elif df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70:
                        insights.append(
                            f"**{display_cat}** obtuvo {df.loc[cat, 'Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}% (Prioridad Media). Considera acciones específicas." if st.session_state.language == "Español" else
                            f"**{display_cat}** scored {df.loc[cat, 'Percent']:.1f}% (Medium Priority). Consider targeted actions."
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
                    "Los resultados de tu auditoría indican oportunidades para optimizar el lugar de trabajo. LEAN 2.0 Institute ofrece consultoría especializada para directivos, gerentes y Recursos Humanos, transformando tu entorno laboral en uno ético y eficiente." if st.session_state.language == "Español" else
                    "Your audit results indicate opportunities to optimize the workplace. LEAN 2.0 Institute offers specialized consulting for directors, managers and HR, transforming your workplace into an ethical and efficient environment."
                )
                if df["Porcentaje" if st.session_state.language == "Español" else "Percent"].min() < 70:
                    low_categories = df[df["Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70].index.tolist()
                    low_display_categories = [next(k for k, v in category_mapping[st.session_state.language].items() if v == cat) for cat in low_categories]
                    services = {
                        "Empoderamiento de Empleados": "Programas de Compromiso y Liderazgo de Empleados",
                        "Liderazgo Ético": "Capacitación en Liderazgo Ético y Gobernanza",
                        "Operaciones Centradas en las Personas": "Optimización de Procesos con Enfoque Humano",
                        "Prácticas Sostenibles y Éticas": "Consultoría en Sostenibilidad y Ética Empresarial",
                        "Bienestar y Equilibrio": "Estrategias de Bienestar Organizacional"
                    } if st.session_state.language == "Español" else {
                        "Empoderamiento de Empleados": "Employee Engagement and Leadership Programs",
                        "Liderazgo Ético": "Ethical Leadership and Governance Training",
                        "Operaciones Centradas en las Personas": "Process Optimization with Human Focus",
                        "Prácticas Sostenibles y Éticas": "Sustainability and Business Ethics Consulting",
                        "Bienestar y Equilibrio": "Organizational Well-Being Strategies"
                    }
                    ad_text.append(
                        f"Las áreas clave para mejorar incluyen {', '.join(low_display_categories)}. LEAN 2.0 Institute se especializa en: {', '.join([services[cat] for cat in low_categories])}." if st.session_state.language == "Español" else
                        f"Key areas for improvement include {', '.join(low_display_categories)}. LEAN 2.0 Institute specializes in: {', '.join([services[cat] for cat in low_categories])}."
                    )
            else:
                ad_text.append(
                    "¡Felicidades por un lugar de trabajo sobresaliente! Colabora con LEAN 2.0 Institute para mantener estas fortalezas y liderar con innovación." if st.session_state.language == "Español" else
                    "Congratulations on an outstanding workplace! Partner with LEAN 2.0 Institute to sustain these strengths and lead with innovation."
                )
            ad_text.append(
                "Contáctanos en https://lean2institute.mystrikingly.com/ o envíanos un correo a contacto@lean2institute.com para una consulta estratégica." if st.session_state.language == "Español" else
                "Contact us at https://lean2institute.mystrikingly.com/ or email us at contacto@lean2institute.com for a strategic consultation."
            )
            st.markdown("<div class='insights'>" + "<br>".join(ad_text) + "</div>", unsafe_allow_html=True)

            # Download button (Excel only)
            st.markdown(
                '<div class="subheader">Descarga tu Informe</div>' if st.session_state.language == "Español" else
                '<div class="subheader">Download Your Report</div>',
                unsafe_allow_html=True
            )
            with st.spinner("Generando Excel..." if st.session_state.language == "Español" else "Generating Excel..."):
                try:
                    excel_output = io.BytesIO()
                    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                        workbook = writer.book
                        bold = workbook.add_format({'bold': True})
                        percent_format = workbook.add_format({'num_format': '0.0%'})
                        red_format = workbook.add_format({'bg_color': '#D32F2F', 'font_color': '#FFFFFF'})
                        yellow_format = workbook.add_format({'bg_color': '#FFD54F', 'font_color': '#212121'})
                        green_format = workbook.add_format({'bg_color': '#43A047', 'font_color': '#FFFFFF'})
                        border_format = workbook.add_format({'border': 1})

                        # Summary Sheet
                        summary_df = pd.DataFrame({
                            "Puntuación General" if st.session_state.language == "Español" else "Overall Score": [f"{overall_score:.1f}%"],
                            "Calificación" if st.session_state.language == "Español" else "Grade": [grade],
                            "Resumen de Hallazgos" if st.session_state.language == "Español" else "Findings Summary": [
                                f"{len(df[df['Porcentaje' if st.session_state.language == 'Español' else 'Percent'] < 50])} categorías requieren acción urgente (<50%), {len(df[(df['Porcentaje' if st.session_state.language == 'Español' else 'Percent'] >= 50) & (df['Porcentaje' if st.session_state.language == 'Español' else 'Percent'] < 70)])} necesitan mejoras específicas (50-69%). La puntuación general es {overall_score:.1f}% ({grade})." if st.session_state.language == "Español" else
                                f"{len(df[df['Percent'] < 50])} categories require urgent action (<50%), {len(df[(df['Percent'] >= 50) & (df['Percent'] < 70)])} need specific improvements (50-69%). Overall score is {overall_score:.1f}% ({grade})."
                            ]
                        })
                        summary_df.to_excel(writer, sheet_name='Resumen' if st.session_state.language == "Español" else 'Summary', index=False)
                        worksheet_summary = writer.sheets['Resumen' if st.session_state.language == "Español" else 'Summary']
                        worksheet_summary.set_column('A:A', 20)
                        worksheet_summary.set_column('B:B', 15)
                        worksheet_summary.set_column('C:C', 60)
                        for col_num, value in enumerate(summary_df.columns.values):
                            worksheet_summary.write(0, col_num, value, bold)
                        # Add contact details and invitation
                        row = len(summary_df) + 2
                        worksheet_summary.write(row, 0, "Colabora con LEAN 2.0 Institute" if st.session_state.language == "Español" else "Partner with LEAN 2.0 Institute", bold)
                        row += 1
                        invitation = (
                            "¡Colabora con LEAN 2.0 Institute para implementar mejoras en tu lugar de trabajo! Contáctanos para una consulta estratégica." if st.session_state.language == "Español" else
                            "Partner with LEAN 2.0 Institute to implement workplace improvements! Contact us for a strategic consultation."
                        )
                        worksheet_summary.write(row, 0, invitation, wrap_format)
                        row += 1
                        worksheet_summary.write(row, 0, "✉️ Email:", bold)
                        worksheet_summary.write(row, 1, "contacto@lean2institute.org")
                        row += 1
                        worksheet_summary.write(row, 0, "🌐 Website:", bold)
                        worksheet_summary.write_url(row, 1, "https://lean2institute.mystrikingly.com/", hyperlink_format, string="https://lean2institute.mystrikingly.com/")

                        # Results Sheet
                        df_display.to_excel(writer, sheet_name='Resultados' if st.session_state.language == "Español" else 'Results', float_format="%.1f")
                        worksheet_results = writer.sheets['Resultados' if st.session_state.language == "Español" else 'Results']
                        worksheet_results.set_column('A:A', 30)
                        worksheet_results.set_column('B:C', 15)
                        worksheet_results.set_column('D:D', 20)
                        for col_num, value in enumerate(df.columns.values):
                            worksheet_results.write(0, col_num + 1, value, bold)
                        worksheet_results.write(0, 0, 'Categoría' if st.session_state.language == "Español" else 'Category', bold)
                        for row_num, value in enumerate(df['Porcentaje' if st.session_state.language == "Español" else 'Percent']):
                            cell_format = red_format if value < 50 else yellow_format if value < 70 else green_format
                            worksheet_results.write(row_num + 1, 2, value / 100, percent_format)
                            worksheet_results.write(row_num + 1, 2, value / 100, cell_format)

                        # Findings and Suggestions Sheet
                        findings_data = []
                        for cat in questions.keys():
                            display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                            if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 70:
                                findings_data.append([
                                    display_cat,
                                    f"{df.loc[cat, 'Porcentaje' if st.session_state.language == 'Español' else 'Percent']:.1f}%",
                                    "Alta" if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 50 else "Media" if st.session_state.language == "Español" else
                                    "High" if df.loc[cat, "Percent"] < 50 else "Medium",
                                    "Acción urgente requerida." if df.loc[cat, "Porcentaje" if st.session_state.language == "Español" else "Percent"] < 50 else "Se necesitan mejoras específicas." if st.session_state.language == "Español" else
                                    "Urgent action required." if df.loc[cat, "Percent"] < 50 else "Specific improvements needed."
                                ])
                                for idx, score in enumerate(st.session_state.responses[cat]):
                                    if score < 70:
                                        question = questions[cat][st.session_state.language][idx][0]
                                        rec = recommendations[cat][idx]
                                        findings_data.append([
                                            "", "", "", f"Pregunta: {question[:50]}... - Sugerencia: {rec}"
                                        ])
                        findings_df = pd.DataFrame(
                            findings_data,
                            columns=[
                                "Categoría" if st.session_state.language == "Español" else "Category",
                                "Puntuación" if st.session_state.language == "Español" else "Score",
                                "Prioridad" if st.session_state.language == "Español" else "Priority",
                                "Hallazgos y Sugerencias" if st.session_state.language == "Español" else "Findings and Suggestions"
                            ]
                        )
                        findings_df.to_excel(writer, sheet_name='Hallazgos' if st.session_state.language == "Español" else 'Findings', index=False)
                        worksheet_findings = writer.sheets['Hallazgos' if st.session_state.language == "Español" else 'Findings']
                        worksheet_findings.set_column('A:A', 30)
                        worksheet_findings.set_column('B:B', 15)
                        worksheet_findings.set_column('C:C', 15)
                        worksheet_findings.set_column('D:D', 60)
                        for col_num, value in enumerate(findings_df.columns.values):
                            worksheet_findings.write(0, col_num, value, bold)
                        for row_num in range(len(findings_df)):
                            for col_num in range(len(findings_df.columns)):
                                worksheet_findings.write(row_num + 1, col_num, findings_df.iloc[row_num, col_num], border_format)

                        # Bar Chart
                        chart_data_df = df_display[["Porcentaje" if st.session_state.language == "Español" else "Percent"]].reset_index()
                        chart_data_df.to_excel(writer, sheet_name='Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', index=False)
                        worksheet_chart = writer.sheets['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data']
                        chart = workbook.add_chart({'type': 'bar'})
                        chart.add_series({
                            'name': 'Puntuación (%)' if st.session_state.language == "Español" else 'Score (%)',
                            'categories': ['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', 1, 0, len(chart_data_df), 0],
                            'values': ['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', 1, 1, len(chart_data_df), 1],
                            'fill': {'color': '#1E88E5'},
                        })
                        chart.set_title({'name': 'Puntuaciones por Categoría' if st.session_state.language == "Español" else 'Category Scores'})
                        chart.set_x_axis({'name': 'Puntuación (%)' if st.session_state.language == "Español" else 'Score (%)'})
                        chart.set_y_axis({'name': 'Categoría' if st.session_state.language == "Español" else 'Category'})
                        worksheet_results.insert_chart('F2', chart)

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

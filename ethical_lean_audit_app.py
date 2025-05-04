import streamlit as st
import pandas as pd
import plotly.express as px
import base64
import io
import xlsxwriter
import os

# Constants
SCORE_THRESHOLDS = {
    "CRITICAL": 50,
    "NEEDS_IMPROVEMENT": 70,
    "GOOD": 85,
}
TOTAL_QUESTIONS = 20
PROGRESS_DISPLAY_THRESHOLD = 20  # Percentage to show unanswered questions
CHART_COLORS = ["#D32F2F", "#FFD54F", "#43A047"]  # Red, Yellow, Green
CHART_HEIGHT = 400
QUESTION_TRUNCATE_LENGTH = 100

# Translation dictionary
TRANSLATIONS = {
    "Español": {
        "title": "Auditoría Ética de Lugar de Trabajo Lean",
        "header": "¡Evalúa y Mejora tu Lugar de Trabajo!",
        "score": "Puntuación",
        "percent": "Porcentaje",
        "priority": "Prioridad",
        "category": "Categoría",
        "question": "Pregunta",
        "high_priority": "Alta",
        "medium_priority": "Media",
        "low_priority": "Baja",
        "report_title": "Tu Informe de Impacto en el Lugar de Trabajo",
        "download_report": "Descargar Informe Excel",
        "report_filename": "resultados_auditoria_lugar_trabajo_etico.xlsx",
        "unanswered_error": "Preguntas sin responder ({}). Por favor, completa todas las preguntas antes de enviar la auditoría.",
        "missing_questions": "Preguntas faltantes:",
        "category_completed": "¡Categoría '{}' completada! {}/{} preguntas respondidas.",
        "all_answered": "¡Todas las preguntas han sido respondidas! Puedes proceder a ver los resultados.",
        "response_guide": "Selecciona la descripción que mejor represente la situación para cada pregunta. Las opciones describen el grado, frecuencia o cantidad aplicable.",
        "language_change_warning": "Cambiar el idioma reiniciará tus respuestas. ¿Deseas continuar?",
    },
    "English": {
        "title": "Ethical Lean Workplace Audit",
        "header": "Assess and Enhance Your Workplace!",
        "score": "Score",
        "percent": "Percent",
        "priority": "Priority",
        "category": "Category",
        "question": "Question",
        "high_priority": "High",
        "medium_priority": "Medium",
        "low_priority": "Low",
        "report_title": "Your Workplace Impact Report",
        "download_report": "Download Excel Report",
        "report_filename": "ethical_workplace_audit_results.xlsx",
        "unanswered_error": "Unanswered questions ({}). Please complete all questions before submitting the audit.",
        "missing_questions": "Missing Questions:",
        "category_completed": "Category '{}' completed! {}/{} questions answered.",
        "all_answered": "All questions have been answered! You can proceed to view the results.",
        "response_guide": "Select the description that best represents the situation for each question. The options describe the degree, frequency, or quantity applicable.",
        "language_change_warning": "Changing the language will reset your responses. Do you wish to continue?",
    }
}

# Cache static data for performance
@st.cache_data
def load_static_data():
    """Load and cache static data like questions and response options."""
    questions = {
        "Empoderamiento de Empleados": {
            "Español": [
                ("1. ¿Qué porcentaje de sugerencias de empleados presentadas en los últimos 12 meses fueron implementadas con resultados documentados?", "percentage", "Establece un sistema formal para rastrear e implementar sugerencias de empleados con métricas claras."),
                ("2. ¿Cuántos empleados recibieron capacitación en habilidades profesionales en el último año?", "count", "Aumenta las oportunidades de capacitación profesional para todos los empleados."),
                ("3. En los últimos 12 meses, ¿cuántos empleados lideraron proyectos o iniciativas con presupuesto asignado?", "count", "Asigna presupuestos a más iniciativas lideradas por empleados para fomentar la innovación."),
                ("4. ¿Con qué frecuencia se realizan foros o reuniones formales para que los empleados compartan retroalimentación con la gerencia?", "frequency", "Programa foros mensuales para retroalimentación directa entre empleados y gerencia.")
            ],
            "English": [
                ("1. What percentage of employee suggestions submitted in the past 12 months were implemented with documented outcomes?", "percentage", "Establish a formal system to track and implement employee suggestions with clear metrics."),
                ("2. How many employees received professional skills training in the past year?", "count", "Increase professional training opportunities for all employees."),
                ("3. In the past 12 months, how many employees led projects or initiatives with allocated budgets?", "count", "Allocate budgets to more employee-led initiatives to foster innovation."),
                ("4. How frequently are formal forums or meetings held for employees to share feedback with management?", "frequency", "Schedule monthly forums for direct employee-management feedback.")
            ]
        },
        "Liderazgo Ético": {
            "Español": [
                ("5. ¿Con qué frecuencia los líderes compartieron actualizaciones escritas sobre decisiones que afectan a los empleados en los últimos 12 meses?", "frequency", "Implementa boletines mensuales para comunicar decisiones de liderazgo de manera transparente."),
                ("6. ¿Qué porcentaje de políticas laborales nuevas o revisadas en el último año incluyó consulta formal con empleados?", "percentage", "Incluye a representantes de empleados en la revisión de todas las políticas laborales nuevas."),
                ("7. ¿Cuántos casos de comportamiento ético destacado fueron reconocidos formalmente en los últimos 12 meses?", "count", "Crea un programa formal de reconocimiento para comportamientos éticos, con incentivos claros.")
            ],
            "English": [
                ("5. How frequently did leaders share written updates on decisions affecting employees in the past 12 months?", "frequency", "Implement monthly newsletters to transparently communicate leadership decisions."),
                ("6. What percentage of new or revised workplace policies in the past year included formal employee consultation?", "percentage", "Include employee representatives in reviewing all new workplace policies."),
                ("7. How many instances of exemplary ethical behavior were formally recognized in the past 12 months?", "count", "Create a formal recognition program for ethical behavior with clear incentives.")
            ]
        },
        "Operaciones Centradas en las Personas": {
            "Español": [
                ("8. ¿Qué porcentaje de procesos lean revisados en los últimos 12 meses incorporó retroalimentación de empleados para reducir tareas redundantes?", "percentage", "Integra retroalimentación de empleados en cada revisión de procesos lean para eliminar redundancias."),
                ("9. ¿Con qué frecuencia se auditan las prácticas operativas para evaluar su impacto en el bienestar de los empleados?", "frequency", "Realiza auditorías trimestrales de prácticas operativas con enfoque en el bienestar."),
                ("10. ¿Cuántos empleados recibieron capacitación en herramientas lean con énfasis en colaboración en el último año?", "count", "Capacita a todos los empleados en herramientas lean, priorizando la colaboración.")
            ],
            "English": [
                ("8. What percentage of lean processes revised in the past 12 months incorporated employee feedback to reduce redundant tasks?", "percentage", "Integrate employee feedback into every lean process review to eliminate redundancies."),
                ("9. How frequently are operational practices audited to assess their impact on employee well-being?", "frequency", "Conduct quarterly audits of operational practices focusing on well-being."),
                ("10. How many employees received training on lean tools emphasizing collaboration in the past year?", "count", "Train all employees on lean tools, prioritizing collaboration.")
            ]
        },
        "Prácticas Sostenibles y Éticas": {
            "Español": [
                ("11. ¿Qué porcentaje de iniciativas lean implementadas en los últimos 12 meses redujo el consumo de recursos?", "percentage", "Lanza iniciativas lean específicas para reducir el consumo de recursos, con metas medibles."),
                ("12. ¿Qué porcentaje de proveedores principales fueron auditados en el último año para verificar estándares laborales y ambientales?", "percentage", "Audita anualmente a todos los proveedores principales para garantizar estándares éticos."),
                ("13. ¿Cuántos empleados participaron en proyectos de sostenibilidad con impacto comunitario o laboral en los últimos 12 meses?", "count", "Involucra a más empleados en proyectos de sostenibilidad con impacto comunitario.")
            ],
            "English": [
                ("11. What percentage of lean initiatives implemented in the past 12 months reduced resource consumption?", "percentage", "Launch specific lean initiatives to reduce resource consumption with measurable goals."),
                ("12. What percentage of primary suppliers were audited in the past year to verify labor and environmental standards?", "percentage", "Audit all primary suppliers annually to ensure ethical standards."),
                ("13. How many employees participated in sustainability projects with community or workplace impact in the past 12 months?", "count", "Engage more employees in sustainability projects with community impact.")
            ]
        },
        "Bienestar y Equilibrio": {
            "Español": [
                ("14. ¿Qué porcentaje de empleados accedió a recursos de bienestar en los últimos 12 meses?", "percentage", "Amplía el acceso a recursos de bienestar, como asesoramiento y horarios flexibles."),
                ("15. ¿Con qué frecuencia se realizan encuestas o revisiones para evaluar el agotamiento o la fatiga de los empleados?", "frequency", "Implementa encuestas mensuales para monitorear el agotamiento y actuar rápidamente."),
                ("16. ¿Cuántos casos de desafíos personales o profesionales reportados por empleados fueron abordados con planes de acción documentados en el último año?", "count", "Establece procesos formales para abordar desafíos reportados con planes de acción.")
            ],
            "English": [
                ("14. What percentage of employees accessed well-being resources in the past 12 months?", "percentage", "Expand access to well-being resources, such as counseling and flexible schedules."),
                ("15. How frequently are surveys or check-ins conducted to assess employee burnout or fatigue?", "frequency", "Implement monthly surveys to monitor burnout and act swiftly."),
                ("16. How many reported employee personal or professional challenges were addressed with documented action plans in the past year?", "count", "Establish formal processes to address reported challenges with action plans.")
            ]
        },
        "Iniciativas Organizacionales Centradas en las Personas": {
            "Español": [
                ("17. En nuestra organización se han implementado o se están explorando tecnologías como Industria 4.0, Inteligencia Artificial, robótica o automatización digital con el propósito de mejorar tanto la eficiencia operativa como las condiciones laborales del personal.", "frequency", "Desarrolla un plan estratégico para integrar tecnologías como IA y robótica, priorizando el impacto positivo en las condiciones laborales."),
                ("18. Contamos con metodologías de excelencia operacional (Lean, Six Sigma, TPM, etc.) que no solo buscan eficiencia y calidad, sino que también integran activamente el bienestar del personal en su diseño e implementación.", "frequency", "Rediseña las metodologías de excelencia operacional para incluir métricas de bienestar del personal en cada fase."),
                ("19. Antes de implementar nuevas tecnologías o iniciativas (sociales, ambientales u operativas), se consulta al personal para asegurar que los cambios beneficien su experiencia y condiciones laborales.", "frequency", "Establece un proceso formal de consulta con los empleados antes de implementar cualquier nueva tecnología o iniciativa."),
                ("20. Las iniciativas actuales (tecnológicas, sociales y operativas) han contribuido de forma tangible a un ambiente laboral más saludable, inclusivo y respetuoso para todos los colaboradores.", "frequency", "Evalúa regularmente el impacto de las iniciativas en el ambiente laboral y ajusta según retroalimentación de los empleados.")
            ],
            "English": [
                ("17. Our organization has implemented or is exploring technologies such as Industry 4.0, AI, robotics, or digital automation to enhance both operational efficiency and employee working conditions.", "frequency", "Develop a strategic plan to integrate technologies like AI and robotics, prioritizing positive impacts on working conditions."),
                ("18. We have operational excellence methodologies (Lean, Six Sigma, TPM, etc.) that not only pursue efficiency and quality but also actively integrate employee well-being into their design and implementation.", "frequency", "Redesign operational excellence methodologies to include employee well-being metrics in every phase."),
                ("19. Before implementing new technologies or initiatives (social, environmental, or operational), employees are consulted to ensure changes benefit their experience and working conditions.", "frequency", "Establish a formal employee consultation process before implementing any new technology or initiative."),
                ("20. Current initiatives (technological, social, and operational) have tangibly contributed to a healthier, more inclusive, and respectful workplace for all employees.", "frequency", "Regularly evaluate the impact of initiatives on the workplace environment and adjust based on employee feedback.")
            ]
        }
    }

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
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Selecciona la descripción que mejor refleje la proporción de casos aplicados."
            },
            "English": {
                "descriptions": [
                    "No suggestions/processes were implemented.",
                    "About one-quarter were implemented.",
                    "Half were implemented.",
                    "Three-quarters were implemented.",
                    "All suggestions/processes were implemented."
                ],
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Select the description that best reflects the proportion of cases applied."
            }
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
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Selecciona la descripción que mejor refleje la frecuencia de la práctica."
            },
            "English": {
                "descriptions": [
                    "This never occurs.",
                    "Occurs very few times a year.",
                    "Occurs several times a year.",
                    "Occurs regularly, almost always.",
                    "Occurs every time."
                ],
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Select the description that best reflects the frequency of the practice."
            }
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
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Selecciona la descripción que mejor refleje la cantidad de empleados o casos afectados."
            },
            "English": {
                "descriptions": [
                    "No employees or cases (0%).",
                    "Less than a quarter of employees (1-25%).",
                    "Between a quarter and half (25-50%).",
                    "More than half but not most (50-75%).",
                    "Over 75% of employees or cases."
                ],
                "scores": [0, 25, 50, 75, 100],
                "tooltip": "Select the description that best reflects the number of employees or cases affected."
            }
        }
    }

    # Validate question types
    valid_q_types = set(response_options.keys())
    for cat in questions:
        for lang in questions[cat]:
            for _, q_type, _ in questions[cat][lang]:
                if q_type not in valid_q_types:
                    raise ValueError(f"Tipo de pregunta inválido '{q_type}' en categoría {cat}, idioma {lang}")

    return questions, response_options

# Load static data
questions, response_options = load_static_data()

# Set page configuration
st.set_page_config(page_title=TRANSLATIONS["Español"]["title"], layout="wide", initial_sidebar_state="expanded")

# Load external CSS
try:
    with open("styles.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
except FileNotFoundError:
    st.warning("Archivo styles.css no encontrado. Usando estilos predeterminados.")

# Configuration from environment variables
CONFIG = {
    "contact": {
        "email": os.getenv("CONTACT_EMAIL", "contacto@lean2institute.org"),
        "website": os.getenv("CONTACT_WEBSITE", "https://lean2institute.mystrikingly.com/")
    }
}

# Category mapping for bilingual support
category_mapping = {
    "Español": {
        "Empoderamiento de Empleados": "Empoderamiento de Empleados",
        "Liderazgo Ético": "Liderazgo Ético",
        "Operaciones Centradas en las Personas": "Operaciones Centradas en las Personas",
        "Prácticas Sostenibles y Éticas": "Prácticas Sostenibles y Éticas",
        "Bienestar y Equilibrio": "Bienestar y Equilibrio",
        "Iniciativas Organizacionales Centradas en las Personas": "Iniciativas Organizacionales Centradas en las Personas"
    },
    "English": {
        "Empowering Employees": "Empoderamiento de Empleados",
        "Ethical Leadership": "Liderazgo Ético",
        "Human-Centered Operations": "Operaciones Centradas en las Personas",
        "Sustainable and Ethical Practices": "Prácticas Sostenibles y Éticas",
        "Well-Being and Balance": "Bienestar y Equilibrio",
        "Human-Centered Organizational Initiatives": "Iniciativas Organizacionales Centradas en las Personas"
    }
}

# Initialize session state with validation
def initialize_session_state():
    """Initialize session state with default values and repair responses if needed."""
    defaults = {
        "language": "Español",
        "responses": {},
        "current_category": 0,
        "prev_language": "Español",
        "show_intro": True,
        "language_changed": False,
        "show_results": False,
        "reset_log": []  # Debug log for tracking adjustments
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    
    # Validate language
    if st.session_state.language not in ["Español", "English"]:
        st.session_state.language = "Español"
        st.session_state.prev_language = "Español"
        st.session_state.reset_log.append("Language reset to Español due to invalid value")
    
    # Initialize or repair responses
    expected_counts = {cat: len(questions[cat]["Español"]) for cat in questions}
    if not st.session_state.responses:
        # Initialize responses if empty
        st.session_state.responses = {
            cat: [None] * expected_counts[cat] for cat in questions
        }
        st.session_state.reset_log.append("Initialized empty responses")
    else:
        # Repair responses for each category
        for cat in questions:
            if cat not in st.session_state.responses:
                # Add missing category
                st.session_state.responses[cat] = [None] * expected_counts[cat]
                st.session_state.reset_log.append(f"Added missing category: {cat}")
            elif len(st.session_state.responses[cat]) != expected_counts[cat]:
                # Adjust response list length
                current_responses = st.session_state.responses[cat]
                st.session_state.responses[cat] = (
                    current_responses[:expected_counts[cat]] +
                    [None] * (expected_counts[cat] - len(current_responses))
                ) if len(current_responses) < expected_counts[cat] else current_responses[:expected_counts[cat]]
                st.session_state.reset_log.append(f"Adjusted responses for {cat}: Expected {expected_counts[cat]}, Found {len(current_responses)}")
    
    # Display debug log only for unexpected adjustments
    if st.session_state.reset_log and any("Added missing category" in log or "Adjusted responses" in log for log in st.session_state.reset_log[-1:]):
        st.warning(f"Debug: Response state adjusted. Log: {st.session_state.reset_log[-1]}")

initialize_session_state()

# Sidebar navigation
with st.sidebar:
    st.markdown('<section class="sidebar-container" role="navigation" aria-label="Navegación de la auditoría">', unsafe_allow_html=True)
    
    # Language selection
    def update_language():
        """Handle language change with confirmation."""
        if st.session_state.language_select != st.session_state.language:
            if any(any(score is not None for score in scores) for scores in st.session_state.responses.values()):
                if not st.session_state.get("language_change_confirmed", False):
                    st.session_state.language_changed = True
                    return
            st.session_state.language = st.session_state.language_select
            st.session_state.current_category = 0
            st.session_state.responses = {
                cat: [None] * len(questions[cat][st.session_state.language]) for cat in questions
            }
            st.session_state.prev_language = st.session_state.language
            st.session_state.show_intro = True
            st.session_state.show_results = False
            st.session_state.language_changed = False
            st.session_state.language_change_confirmed = False
            st.session_state.reset_log.append("Responses reset due to language change")
    
    st.selectbox(
        "Idioma / Language",
        ["Español", "English"],
        key="language_select",
        help="Selecciona tu idioma preferido / Select your preferred language",
        on_change=update_language
    )
    
    # Language change confirmation
    if st.session_state.get("language_changed", False):
        st.warning(TRANSLATIONS[st.session_state.language]["language_change_warning"])
        if st.button("Confirmar / Confirm", key="confirm_language_change", type="primary"):
            st.session_state.language_change_confirmed = True
            update_language()
        if st.button("Cancelar / Cancel", key="cancel_language_change"):
            st.session_state.language_select = st.session_state.language
            st.session_state.language_changed = False
    
    st.markdown('<h2 class="sidebar-title" role="heading" aria-level="2">Progreso</h2>', unsafe_allow_html=True)
    display_categories = list(category_mapping[st.session_state.language].keys())
    for i, display_cat in enumerate(display_categories):
        status = 'active' if i == st.session_state.current_category else 'completed' if i < st.session_state.current_category else ''
        if st.button(
            f"{display_cat}",
            key=f"nav_{i}",
            help=f"Ir a la categoría {display_cat}",
            disabled=False,
            use_container_width=True,
            type="primary" if status == 'active' else "secondary"
        ):
            st.session_state.current_category = max(0, min(i, len(display_categories) - 1))
            st.session_state.show_intro = False
            st.session_state.show_results = False
    st.markdown('</section>', unsafe_allow_html=True)

# Introductory modal
if st.session_state.show_intro:
    with st.container():
        st.markdown('<section class="container intro-section" role="main">', unsafe_allow_html=True)
        st.markdown(
            f'<h1 class="main-title" role="heading" aria-level="1">🤝 LEFingerprint 2.0 Institute</h1>',
            unsafe_allow_html=True
        )
        with st.expander("", expanded=True):
            st.markdown(
                f"""
                <div class="intro-content">
                    Esta evaluación está diseñada para ser completada por la gerencia en conjunto con Recursos Humanos, proporcionando una evaluación objetiva de tu entorno laboral. Responde {TOTAL_QUESTIONS} preguntas en {len(questions)} categorías (5–10 minutos) con datos específicos y ejemplos verificables. Tus respuestas son confidenciales y generarán un informe detallado con recomendaciones accionables que podemos ayudarte a implementar. Al completar la evaluación, contáctanos para consultas personalizadas: ✉️ Email: <a href="mailto:{CONFIG['contact']['email']}">{CONFIG['contact']['email']}</a> 🌐 Website: <a href="{CONFIG['contact']['website']}">{CONFIG['contact']['website']}</a>
                    
                    <h3 class="subsection-title">Pasos:</h3>
                    <ol class="steps-list" role="list" aria-label="Pasos para completar la auditoría">
                        <li>Responde las preguntas de cada categoría.</li>
                        <li>Revisa y descarga tu informe.</li>
                    </ol>
                    
                    ¡Empecemos!
                </div>
                """
                if st.session_state.language == "Español" else
                f"""
                <div class="intro-content">
                    This assessment is designed for management and HR to provide an objective evaluation of your workplace. Answer {TOTAL_QUESTIONS} questions across {len(questions)} categories (5–10 minutes) with specific data and verifiable examples. Your responses are confidential and will generate a detailed report with actionable recommendations we can help implement. Upon completion, contact us for personalized consultations: ✉️ Email: <a href="mailto:{CONFIG['contact']['email']}">{CONFIG['contact']['email']}</a> 🌐 Website: <a href="{CONFIG['contact']['website']}">{CONFIG['contact']['website']}</a>
                    
                    <h3 class="subsection-title">Steps:</h3>
                    <ol class="steps-list" role="list" aria-label="Steps to complete the audit">
                        <li>Answer questions for each category.</li>
                        <li>Review and download your report.</li>
                    </ol>
                    
                    Let’s get started!
                </div>
                """
            )
            if st.button(
                "Iniciar Auditoría / Start Audit",
                use_container_width=True,
                key="start_audit",
                help="Comenzar la evaluación / Begin the assessment",
                type="primary"
            ):
                st.session_state.show_intro = False
        st.markdown('</section>', unsafe_allow_html=True)

# Main content
if not st.session_state.show_intro:
    with st.container():
        st.markdown('<section class="container main-section" role="main">', unsafe_allow_html=True)
        st.markdown(
            f'<h1 class="main-title" role="heading" aria-level="1">{TRANSLATIONS[st.session_state.language]["header"]}</h1>',
            unsafe_allow_html=True
        )

        # Progress indicator
        completed_questions = sum(len([s for s in scores if s is not None]) for scores in st.session_state.responses.values())
        completion_percentage = (completed_questions / TOTAL_QUESTIONS) * 100 if TOTAL_QUESTIONS > 0 else 0
        st.markdown(
            f"""
            <div class="progress-circle" role="progressbar" aria-label="Progreso: {completion_percentage:.1f}% completado">
                <svg class="progress-ring" width="120" height="120">
                    <circle class="progress-ring__background" cx="60" cy="60" r="54" />
                    <circle class="progress-ring__circle" cx="60" cy="60" r="54" stroke-dasharray="339.292" stroke-dashoffset="{339.292 * (1 - completion_percentage / 100)}" />
                    <text x="50%" y="50%" text-anchor="middle" dy=".3em">{completed_questions}/{TOTAL_QUESTIONS}</text>
                </svg>
                <div class="progress-label">{completion_percentage:.1f}% Completado</div>
            </div>
            <span class="sr-only">{completed_questions} de {TOTAL_QUESTIONS} preguntas completadas ({completion_percentage:.1f}%)</span>
            """,
            unsafe_allow_html=True
        )

        # Check if audit is complete
        audit_complete = all(all(score is not None for score in scores) for scores in st.session_state.responses.values())

        # Display unanswered questions after threshold
        if completion_percentage >= PROGRESS_DISPLAY_THRESHOLD:
            unanswered_questions = []
            question_counter = 1
            for cat in questions.keys():
                for i, (q, _, _) in enumerate(questions[cat][st.session_state.language]):
                    if st.session_state.responses[cat][i] is None:
                        display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                        truncated_q = q[:QUESTION_TRUNCATE_LENGTH] + ("..." if len(q) > QUESTION_TRUNCATE_LENGTH else "")
                        unanswered_questions.append(f"{display_cat}: Pregunta {question_counter} - {truncated_q}")
                    question_counter += 1
            if unanswered_questions:
                st.error(
                    TRANSLATIONS[st.session_state.language]["unanswered_error"].format(len(unanswered_questions)),
                    icon="⚠️"
                )
                st.markdown(
                    f"""
                    <div class="alert alert-warning" role="alert">
                        <strong>{TRANSLATIONS[st.session_state.language]["missing_questions"]}</strong>
                        <ul>
                            {"".join([f"<li>{q}</li>" for q in unanswered_questions])}
                        </ul>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            else:
                st.success(
                    TRANSLATIONS[st.session_state.language]["all_answered"],
                    icon="✅"
                )

        # Category questions
        if not st.session_state.show_results:
            category_index = max(0, min(st.session_state.current_category, len(display_categories) - 1))
            display_category = display_categories[category_index]
            category = category_mapping[st.session_state.language][display_category]
            
            with st.container():
                st.markdown(f'<div class="card-modern" role="region" aria-label="Preguntas de la categoría {display_category}">', unsafe_allow_html=True)
                st.markdown(f'<h2 class="section-title" role="heading" aria-level="2">{display_category}</h2>', unsafe_allow_html=True)
                
                # Response guide
                with st.expander(TRANSLATIONS[st.session_state.language]["response_guide"], expanded=True):
                    st.markdown(f'<div class="response-guide">{TRANSLATIONS[st.session_state.language]["response_guide"]}</div>', unsafe_allow_html=True)

                for idx, (q, q_type, _) in enumerate(questions[category][st.session_state.language]):
                    with st.container():
                        is_unanswered = st.session_state.responses[category][idx] is None
                        st.markdown(
                            f"""
                            <div class="question-container">
                                <label class="question-text" for="{category}_{idx}">
                                    {q} {'<span class="required" aria-label="Requerido">*</span>' if is_unanswered else ''}
                                </label>
                                <div class="tooltip">
                                    <span class="tooltip-icon">?</span>
                                    <span class="tooltip-text">{response_options[q_type][st.session_state.language]['tooltip']}</span>
                                </div>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                        descriptions = response_options[q_type][st.session_state.language]["descriptions"]
                        scores = response_options[q_type][st.session_state.language]["scores"]
                        selected_description = st.radio(
                            "",
                            descriptions,
                            format_func=lambda x: x,
                            key=f"{category}_{idx}",
                            horizontal=False,
                            help=response_options[q_type][st.session_state.language]['tooltip'],
                            label_visibility="hidden",
                            args=({"aria-label": f"Respuesta para la pregunta: {q}"},)
                        )
                        score_idx = descriptions.index(selected_description)
                        st.session_state.responses[category][idx] = scores[score_idx]
                st.markdown('</div>', unsafe_allow_html=True)

                # Progress checkpoint after category completion
                if all(score is not None for score in st.session_state.responses[category]):
                    st.success(
                        TRANSLATIONS[st.session_state.language]["category_completed"].format(display_category, completed_questions, TOTAL_QUESTIONS),
                        icon="🎉"
                    )

            # Sticky navigation
            with st.container():
                st.markdown('<nav class="sticky-nav" role="navigation" aria-label="Navegación entre categorías">', unsafe_allow_html=True)
                col1, col2 = st.columns([1, 1], gap="small")
                with col1:
                    if st.button(
                        "⬅ Anterior" if st.session_state.language == "Español" else "⬅ Previous",
                        disabled=category_index == 0,
                        use_container_width=True,
                        key="prev_category",
                        help="Volver a la categoría anterior" if st.session_state.language == "Español" else "Go to previous category",
                        type="secondary"
                    ):
                        st.session_state.current_category = max(category_index - 1, 0)
                        st.session_state.show_results = False
                with col2:
                    if category_index < len(display_categories) - 1:
                        if st.button(
                            "Siguiente ➡" if st.session_state.language == "Español" else "Next ➡",
                            disabled=category_index == len(display_categories) - 1,
                            use_container_width=True,
                            key="next_category",
                            help="Avanzar a la siguiente categoría" if st.session_state.language == "Español" else "Go to next category",
                            type="primary"
                        ):
                            if all(score is not None for score in st.session_state.responses[category]):
                                st.session_state.current_category = min(category_index + 1, len(display_categories) - 1)
                                st.session_state.show_results = False
                            else:
                                unanswered = [q for i, (q, _, _) in enumerate(questions[category][st.session_state.language]) if st.session_state.responses[category][i] is None]
                                st.error(
                                    f"Por favor, responde todas las preguntas en esta categoría antes de continuar." if st.session_state.language == "Español" else
                                    f"Please answer all questions in this category before proceeding.",
                                    icon="⚠️"
                                )
                                first_unanswered_idx = next((i for i, score in enumerate(st.session_state.responses[category]) if score is None), None)
                                if first_unanswered_idx is not None:
                                    st.markdown(
                                        f"""
                                        <script>
                                            document.getElementById('{category}_{idx}').scrollIntoView({{behavior: 'smooth', block: 'center'}});
                                            document.getElementById('{category}_{idx}').focus();
                                        </script>
                                        """,
                                        unsafe_allow_html=True
                                    )
                    else:
                        if st.button(
                            "Ver Resultados" if st.session_state.language == "Español" else "View Results",
                            use_container_width=True,
                            key="view_results",
                            help="Ver el informe de la auditoría" if st.session_state.language == "Español" else "View audit report",
                            disabled=not audit_complete,
                            type="primary"
                        ):
                            if audit_complete:
                                st.session_state.show_results = True
                            else:
                                st.error(
                                    f"Por favor, responde todas las preguntas en todas las categorías. Revisa las preguntas faltantes arriba." if st.session_state.language == "Español" else
                                    f"Please answer all questions in all categories. Review the missing questions above.",
                                    icon="⚠️"
                                )
                st.markdown('</nav>', unsafe_allow_html=True)

        # Grading matrix
        def get_grade(score):
            """
            Determine the grade and description based on the score.
            
            Args:
                score (float): The average score for a category or overall.
            
            Returns:
                tuple: (grade, description, CSS class)
            """
            if score >= SCORE_THRESHOLDS["GOOD"]:
                return (
                    "Excelente" if st.session_state.language == "Español" else "Excellent",
                    "Tu lugar de trabajo demuestra prácticas sobresalientes. ¡Continúa fortaleciendo estas áreas!" if st.session_state.language == "Español" else
                    "Your workplace demonstrates outstanding practices. Continue strengthening these areas!",
                    "grade-excellent"
                )
            elif score >= SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                return (
                    "Bueno" if st.session_state.language == "Español" else "Good",
                    "Tu lugar de trabajo tiene fortalezas, pero requiere mejoras específicas para alcanzar la excelencia." if st.session_state.language == "Español" else
                    "Your workplace has strengths but requires specific improvements to achieve excellence.",
                    "grade-good"
                )
            elif score >= SCORE_THRESHOLDS["CRITICAL"]:
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
            st.markdown(f'<div class="card-modern report-section" role="region" aria-label="{TRANSLATIONS[st.session_state.language]["report_title"]}">', unsafe_allow_html=True)
            st.markdown(
                f'<h2 class="section-title" role="heading" aria-level="2">{TRANSLATIONS[st.session_state.language]["report_title"]}</h2>',
                unsafe_allow_html=True
            )
            st.markdown(
                '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por tu compromiso con la construcción de un entorno laboral saludable, seguro y respetuoso para todas las personas!</div>' if st.session_state.language == "Español" else
                '<div class="badge">🏆 Audit Completed! Thank you for your commitment to fostering a healthy, safe, and respectful work environment for everyone!</div>', 
                unsafe_allow_html=True
            )

            # Calculate scores
            results = {cat: sum(scores) / len(scores) for cat, scores in st.session_state.responses.items()}
            df = pd.DataFrame.from_dict(results, orient="index", columns=[TRANSLATIONS[st.session_state.language]["score"]])
            df[TRANSLATIONS[st.session_state.language]["percent"]] = df[TRANSLATIONS[st.session_state.language]["score"]]
            df[TRANSLATIONS[st.session_state.language]["priority"]] = df[TRANSLATIONS[st.session_state.language]["percent"]].apply(
                lambda x: TRANSLATIONS[st.session_state.language]["high_priority"] if x < SCORE_THRESHOLDS["CRITICAL"] else 
                          TRANSLATIONS[st.session_state.language]["medium_priority"] if x < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else 
                          TRANSLATIONS[st.session_state.language]["low_priority"]
            )

            # Summary dashboard
            st.markdown('<h3 class="subsection-title" role="heading" aria-level="3">Resumen Ejecutivo</h3>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1:
                overall_score = df[TRANSLATIONS[st.session_state.language]["percent"]].mean()
                grade, grade_description, grade_class = get_grade(overall_score)
                st.markdown(
                    f'<div class="grade {grade_class}">Calificación General: {grade} ({overall_score:.1f}%)</div>' if st.session_state.language == "Español" else
                    f'<div class="grade {grade_class}">Overall Grade: {grade} ({overall_score:.1f}%)</div>',
                    unsafe_allow_html=True
                )
                st.markdown(f'<p class="grade-description">{grade_description}</p>', unsafe_allow_html=True)
            with col2:
                st.metric(
                    "Categorías con Alta Prioridad" if st.session_state.language == "Español" else "High Priority Categories",
                    len(df[df[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
                )
            with col3:
                st.metric(
                    "Puntuación Promedio" if st.session_state.language == "Español" else "Average Score",
                    f"{overall_score:.1f}%"
                )

            # Color-coded dataframe
            def color_percent(val):
                color = CHART_COLORS[0] if val < SCORE_THRESHOLDS["CRITICAL"] else CHART_COLORS[1] if val < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else CHART_COLORS[2]
                return f'background-color: {color}; color: white;'
            
            st.markdown(
                '<p class="chart-legend">Puntuaciones por debajo del 50% (rojo) requieren acción urgente, 50–69% (amarillo) sugieren mejoras, y por encima del 70% (verde) indican fortalezas.</p>' if st.session_state.language == "Español" else
                '<p class="chart-legend">Scores below 50% (red) need urgent action, 50–69% (yellow) suggest improvement, and above 70% (green) indicate strengths.</p>',
                unsafe_allow_html=True
            )
            styled_df = df.style.applymap(color_percent, subset=[TRANSLATIONS[st.session_state.language]["percent"]]).format({TRANSLATIONS[st.session_state.language]["percent"]: "{:.1f}%"})
            st.dataframe(styled_df, use_container_width=True)

            # Interactive bar chart
            df_display = df.copy()
            df_display.index = [next(k for k, v in category_mapping[st.session_state.language].items() if v == idx) for idx in df.index]
            fig = px.bar(
                df_display.reset_index(),
                y="index",
                x=TRANSLATIONS[st.session_state.language]["percent"],
                orientation='h',
                title="Fortalezas y Oportunidades del Lugar de Trabajo" if st.session_state.language == "Español" else "Workplace Strengths and Opportunities",
                labels={
                    "index": TRANSLATIONS[st.session_state.language]["category"],
                    TRANSLATIONS[st.session_state.language]["percent"]: "Puntuación (%)" if st.session_state.language == "Español" else "Score (%)"
                },
                color=TRANSLATIONS[st.session_state.language]["percent"],
                color_continuous_scale=CHART_COLORS,
                range_x=[0, 100],
                height=CHART_HEIGHT
            )
            fig.add_vline(x=SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"], line_dash="dash", line_color="blue", annotation_text="Objetivo (70%)" if st.session_state.language == "Español" else "Target (70%)", annotation_position="top")
            for i, row in df_display.iterrows():
                if row[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                    fig.add_annotation(
                        x=row[TRANSLATIONS[st.session_state.language]["percent"]], y=i,
                        text=TRANSLATIONS[st.session_state.language]["priority"], showarrow=True, arrowhead=2, ax=20, ay=-30,
                        font=dict(color="red", size=12)
                    )
            fig.update_layout(
                showlegend=False,
                title_x=0.5,
                xaxis_title="Puntuación (%)" if st.session_state.language == "Español" else "Score (%)",
                yaxis_title=TRANSLATIONS[st.session_state.language]["category"],
                coloraxis_showscale=False,
                margin=dict(l=150, r=50, t=100, b=50)
            )
            st.plotly_chart(fig, use_container_width=True)

            # Question-level breakdown
            with st.expander("Análisis Detallado: Perspectivas a Nivel de Pregunta" if st.session_state.language == "Español" else "Drill Down: Question-Level Insights"):
                selected_display_category = st.selectbox(
                    "Seleccionar Categoría para Explorar" if st.session_state.language == "Español" else "Select Category to Explore",
                    display_categories,
                    key="category_explore",
                    help="Elige una categoría para ver las puntuaciones de sus preguntas" if st.session_state.language == "Español" else "Choose a category to view its question scores"
                )
                selected_category = category_mapping[st.session_state.language][selected_display_category]
                question_scores = pd.DataFrame({
                    TRANSLATIONS[st.session_state.language]["question"]: [q for q, _, _ in questions[selected_category][st.session_state.language]],
                    TRANSLATIONS[st.session_state.language]["score"]: st.session_state.responses[selected_category]
                })
                fig_questions = px.bar(
                    question_scores,
                    x=TRANSLATIONS[st.session_state.language]["score"],
                    y=TRANSLATIONS[st.session_state.language]["question"],
                    orientation='h',
                    title=f"Puntuaciones de Preguntas para {selected_display_category}" if st.session_state.language == "Español" else f"Question Scores for {selected_display_category}",
                    labels={
                        TRANSLATIONS[st.session_state.language]["score"]: "Puntuación (%)" if st.session_state.language == "Español" else "Score (%)",
                        TRANSLATIONS[st.session_state.language]["question"]: TRANSLATIONS[st.session_state.language]["question"]
                    },
                    color=TRANSLATIONS[st.session_state.language]["score"],
                    color_continuous_scale=CHART_COLORS,
                    range_x=[0, 100],
                    height=300 + len(question_scores) * 50
                )
                fig_questions.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    xaxis_title="Puntuación (%)" if st.session_state.language == "Español" else "Score (%)",
                    yaxis_title=TRANSLATIONS[st.session_state.language]["question"],
                    coloraxis_showscale=False,
                    margin=dict(l=150, r=50, t=100, b=50)
                )
                st.plotly_chart(fig_questions, use_container_width=True)

            # Actionable insights
            with st.expander("Perspectivas Accionables" if st.session_state.language == "Español" else "Actionable Insights"):
                insights = []
                for cat in questions.keys():
                    display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                    score = df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]]
                    if score < SCORE_THRESHOLDS["CRITICAL"]:
                        insights.append(
                            f"**{display_cat}** obtuvo {score:.1f}% (Alta Prioridad). Enfócate en mejoras inmediatas." if st.session_state.language == "Español" else
                            f"**{display_cat}** scored {score:.1f}% (High Priority). Focus on immediate improvements."
                        )
                    elif score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        insights.append(
                            f"**{display_cat}** obtuvo {score:.1f}% (Prioridad Media). Considera acciones específicas." if st.session_state.language == "Español" else
                            f"**{display_cat}** scored {score:.1f}% (Medium Priority). Consider targeted actions."
                        )
                if insights:
                    st.markdown(
                        "<div class='alert alert-info'>" + "<br>".join(insights) + "</div>",
                        unsafe_allow_html=True
                    )
                else:
                    st.markdown(
                        "<div class='alert alert-success'>¡Todas las categorías obtuvieron más del 70%! Continúa manteniendo estas fortalezas.</div>" if st.session_state.language == "Español" else 
                        "<div class='alert alert-success'>All categories scored above 70%! Continue maintaining these strengths.</div>",
                        unsafe_allow_html=True
                    )

            # LEAN 2.0 Institute Advertisement
            st.markdown(
                "<h3 class='subsection-title'>Optimiza tu Lugar de Trabajo con LEAN 2.0 Institute</h3>" if st.session_state.language == "Español" else 
                "<h3 class='subsection-title'>Optimize Your Workplace with LEAN 2.0 Institute</h3>",
                unsafe_allow_html=True
            )
            ad_text = []
            if overall_score < SCORE_THRESHOLDS["GOOD"]:
                ad_text.append(
                    "Los resultados de tu auditoría indican oportunidades para optimizar el lugar de trabajo. LEAN 2Ԙ Institute ofrece consultoría especializada para directivos, gerentes y Recursos Humanos, transformando tu entorno laboral en uno ético y eficiente." if st.session_state.language == "Español" else
                    "Your audit results indicate opportunities to optimize the workplace. LEAN 2.0 Institute offers specialized consulting for directors, managers and HR, transforming your workplace into an ethical and efficient environment."
                )
                if df[TRANSLATIONS[st.session_state.language]["percent"]].min() < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                    low_categories = df[df[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]].index.tolist()
                    low_display_categories = [next(k for k, v in category_mapping[st.session_state.language].items() if v == cat) for cat in low_categories]
                    services = {
                        "Empoderamiento de Empleados": "Programas de Compromiso y Liderazgo de Empleados",
                        "Liderazgo Ético": "Capacitación en Liderazgo Ético y Gobernanza",
                        "Operaciones Centradas en las Personas": "Optimización de Procesos con Enfoque Humano",
                        "Prácticas Sostenibles y Éticas": "Consultoría en Sostenibilidad y Ética Empresarial",
                        "Bienestar y Equilibrio": "Estrategias de Bienestar Organizacional",
                        "Iniciativas Organizacionales Centradas en las Personas": "Consultoría en Iniciativas Tecnológicas y Sociales Centradas en las Personas"
                    } if st.session_state.language == "Español" else {
                        "Empoderamiento de Empleados": "Employee Engagement and Leadership Programs",
                        "Liderazgo Ético": "Ethical Leadership and Governance Training",
                        "Operaciones Centradas en las Personas": "Process Optimization with Human Focus",
                        "Prácticas Sostenibles y Éticas": "Sustainability and Business Ethics Consulting",
                        "Bienestar y Equilibrio": "Organizational Well-Being Strategies",
                        "Iniciativas Organizacionales Centradas en las Personas": "Consulting on Human-Centered Technological and Social Initiatives"
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
                f"Contáctanos en <a href='{CONFIG['contact']['website']}'>{CONFIG['contact']['website']}</a> o envíanos un correo a <a href='mailto:{CONFIG['contact']['email']}'>{CONFIG['contact']['email']}</a> para una consulta estratégica." if st.session_state.language == "Español" else
                f"Contact us at <a href='{CONFIG['contact']['website']}'>{CONFIG['contact']['website']}</a> or email us at <a href='mailto:{CONFIG['contact']['email']}'>{CONFIG['contact']['email']}</a> for a strategic consultation."
            )
            st.markdown("<div class='alert alert-info'>" + "<br>".join(ad_text) + "</div>", unsafe_allow_html=True)

            # Download button (Excel only)
            def generate_excel_report():
                """
                Generate an Excel report with summary, results, findings, and chart data.
                
                Returns:
                    BytesIO: Excel file buffer
                """
                excel_output = io.BytesIO()
                with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    bold = workbook.add_format({'bold': True})
                    percent_format = workbook.add_format({'num_format': '0.0%'})
                    red_format = workbook.add_format({'bg_color': CHART_COLORS[0], 'font_color': '#FFFFFF'})
                    yellow_format = workbook.add_format({'bg_color': CHART_COLORS[1], 'font_color': '#212121'})
                    green_format = workbook.add_format({'bg_color': CHART_COLORS[2], 'font_color': '#FFFFFF'})
                    border_format = workbook.add_format({'border': 1})
                    wrap_format = workbook.add_format({'text_wrap': True})
                    hyperlink_format = workbook.add_format({'font_color': 'blue', 'underline': 1})

                    # Summary Sheet
                    summary_df = pd.DataFrame({
                        "Puntuación General" if st.session_state.language == "Español" else "Overall Score": [f"{overall_score:.1f}%"],
                        "Calificación" if st.session_state.language == "Español" else "Grade": [grade],
                        "Resumen de Hallazgos" if st.session_state.language == "Español" else "Findings Summary": [
                            f"{len(df[df[TRANSLATIONS[st.session_state.language]['percent']] < SCORE_THRESHOLDS['CRITICAL']])} categorías requieren acción urgente (<{SCORE_THRESHOLDS['CRITICAL']}%), {len(df[(df[TRANSLATIONS[st.session_state.language]['percent']] >= SCORE_THRESHOLDS['CRITICAL']) & (df[TRANSLATIONS[st.session_state.language]['percent']] < SCORE_THRESHOLDS['NEEDS_IMPROVEMENT'])])} necesitan mejoras específicas ({SCORE_THRESHOLDS['CRITICAL']}-{SCORE_THRESHOLDS['NEEDS_IMPROVEMENT']-1}%). La puntuación general es {overall_score:.1f}% ({grade})." if st.session_state.language == "Español" else
                            f"{len(df[df[TRANSLATIONS[st.session_state.language]['percent']] < SCORE_THRESHOLDS['CRITICAL']])} categories require urgent action (<{SCORE_THRESHOLDS['CRITICAL']}%), {len(df[(df[TRANSLATIONS[st.session_state.language]['percent']] >= SCORE_THRESHOLDS['CRITICAL']) & (df[TRANSLATIONS[st.session_state.language]['percent']] < SCORE_THRESHOLDS['NEEDS_IMPROVEMENT'])])} need specific improvements ({SCORE_THRESHOLDS['CRITICAL']}-{SCORE_THRESHOLDS['NEEDS_IMPROVEMENT']-1}%). Overall score is {overall_score:.1f}% ({grade})."
                        ]
                    })
                    summary_df.to_excel(writer, sheet_name='Resumen' if st.session_state.language == "Español" else 'Summary', index=False)
                    worksheet_summary = writer.sheets['Resumen' if st.session_state.language == "Español" else 'Summary']
                    worksheet_summary.set_column('A:A', 20)
                    worksheet_summary.set_column('B:B', 15)
                    worksheet_summary.set_column('C:C', 80)
                    for col_num, value in enumerate(summary_df.columns.values):
                        worksheet_summary.write(0, col_num, value, bold)
                    # Add contact details
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
                    worksheet_summary.write_url(row, 1, f"mailto:{CONFIG['contact']['email']}", hyperlink_format, string=CONFIG['contact']['email'])
                    row += 1
                    worksheet_summary.write(row, 0, "🌐 Website:", bold)
                    worksheet_summary.write_url(row, 1, CONFIG['contact']['website'], hyperlink_format, string=CONFIG['contact']['website'])

                    # Results Sheet
                    df_display.to_excel(writer, sheet_name='Resultados' if st.session_state.language == "Español" else 'Results', float_format="%.1f")
                    worksheet_results = writer.sheets['Resultados' if st.session_state.language == "Español" else 'Results']
                    worksheet_results.set_column('A:A', 30)
                    worksheet_results.set_column('B:C', 15)
                    worksheet_results.set_column('D:D', 20)
                    for col_num, value in enumerate(df.columns.values):
                        worksheet_results.write(0, col_num + 1, value, bold)
                    worksheet_results.write(0, 0, TRANSLATIONS[st.session_state.language]["category"], bold)
                    for row_num, value in enumerate(df[TRANSLATIONS[st.session_state.language]["percent"]]):
                        cell_format = red_format if value < SCORE_THRESHOLDS["CRITICAL"] else yellow_format if value < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else green_format
                        worksheet_results.write(row_num + 1, 2, value / 100, percent_format)
                        worksheet_results.write(row_num + 1, 2, value / 100, cell_format)

                    # Findings and Suggestions Sheet
                    findings_data = []
                    for cat in questions.keys():
                        display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                        if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                            findings_data.append([
                                display_cat,
                                f"{df.loc[cat, TRANSLATIONS[st.session_state.language]['percent']]:.1f}%",
                                TRANSLATIONS[st.session_state.language]["high_priority"] if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else TRANSLATIONS[st.session_state.language]["medium_priority"],
                                "Acción urgente requerida." if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else "Se necesitan mejoras específicas." if st.session_state.language == "Español" else
                                "Urgent action required." if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else "Specific improvements needed."
                            ])
                            for idx, score in enumerate(st.session_state.responses[cat]):
                                if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                                    question, _, rec = questions[cat][st.session_state.language][idx]
                                    findings_data.append([
                                        "", "", "", f"Pregunta: {question[:50]}... - Sugerencia: {rec}"
                                    ])
                    findings_df = pd.DataFrame(
                        findings_data,
                        columns=[
                            TRANSLATIONS[st.session_state.language]["category"],
                            TRANSLATIONS[st.session_state.language]["score"],
                            TRANSLATIONS[st.session_state.language]["priority"],
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
                    chart_data_df = df_display[[TRANSLATIONS[st.session_state.language]["percent"]]].reset_index()
                    chart_data_df.to_excel(writer, sheet_name='Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', index=False)
                    worksheet_chart = writer.sheets['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data']
                    chart = workbook.add_chart({'type': 'bar'})
                    chart.add_series({
                        'name': 'Puntuación (%)' if st.session_state.language == "Español" else 'Score (%)',
                        'categories': ['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', 1, 0, len(chart_data_df), 0],
                        'values': ['Datos_Gráfico' if st.session_state.language == "Español" else 'Chart_Data', 1, 1, len(chart_data_df), 1],
                        'fill': {'color': CHART_COLORS[2]},
                    })
                    chart.set_title({'name': 'Puntuaciones por Categoría' if st.session_state.language == "Español" else 'Category Scores'})
                    chart.set_x_axis({'name': 'Puntuación (%)' if st.session_state.language == "Español" else 'Score (%)'})
                    chart.set_y_axis({'name': TRANSLATIONS[st.session_state.language]["category"]})
                    chart.set_size({'width': 600, 'height': CHART_HEIGHT})
                    worksheet_results.insert_chart('F2', chart)

                excel_output.seek(0)
                return excel_output

            st.markdown(
                f'<h3 class="subsection-title">Descarga tu Informe</h3>',
                unsafe_allow_html=True
            )
            with st.spinner("Generando Excel..." if st.session_state.language == "Español" else "Generating Excel..."):
                try:
                    excel_output = generate_excel_report()
                    b64_excel = base64.b64encode(excel_output.getvalue()).decode()
                    href_excel = (
                        f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="{TRANSLATIONS[st.session_state.language]["report_filename"]}" class="btn btn-primary" role="button" aria-label="{TRANSLATIONS[st.session_state.language]["download_report"]}">{TRANSLATIONS[st.session_state.language]["download_report"]}</a>'
                    )
                    st.markdown(href_excel, unsafe_allow_html=True)
                    excel_output.close()
                except ImportError:
                    st.error(
                        "La exportación a Excel requiere 'xlsxwriter'. Por favor, instálalo usando `pip install xlsxwriter`. Si estás en Streamlit Cloud, agrega 'xlsxwriter' a tu archivo requirements.txt." if st.session_state.language == "Español" else
                        "Excel export requires 'xlsxwriter'. Please install it using `pip install xlsxwriter`. If on Streamlit Cloud, add 'xlsxwriter' to your requirements.txt file.",
                        icon="❌"
                    )
                except Exception as e:
                    st.error(
                        f"No se pudo generar el archivo Excel: {str(e)}" if st.session_state.language == "Español" else
                        f"Failed to generate Excel file: {str(e)}",
                        icon="❌"
                    )

            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</section>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</section>', unsafe_allow_html=True)

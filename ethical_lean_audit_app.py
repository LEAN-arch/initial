import streamlit as st
import pandas as pd
import plotly.express as px
import io
import xlsxwriter
import os
import uuid
import re
from datetime import datetime
from typing import Dict, List, Tuple

# Constants
SCORE_THRESHOLDS = {
    "CRITICAL": 50,
    "NEEDS_IMPROVEMENT": 70,
    "GOOD": 85,
}
TOTAL_QUESTIONS = 25
CHART_COLORS = ["#D32F2F", "#FFD54F", "#43A047"]
CHART_HEIGHT = 400
QUESTION_TRUNCATE_LENGTH = 100
REPORT_DATE = datetime.now().strftime("%Y-%m-%d")

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
        "report_title": "Tu Informe de Bienestar Laboral",
        "download_excel": "Descargar Informe Excel",
        "report_filename_excel": "resultados_auditoria_lugar_trabajo_etico.xlsx",
        "unanswered_error": "No se pueden mostrar los resultados. Hay {} preguntas sin responder. Por favor, completa todas las preguntas.",
        "missing_questions": "Preguntas faltantes:",
        "all_answered": "¡Todas las preguntas han sido respondidas! Revisa los resultados abajo.",
        "response_guide": "Selecciona la descripción que mejor represente la situación para cada pregunta. Las opciones describen el grado, frecuencia o cantidad aplicable.",
        "language_change_warning": "Cambiar el idioma reiniciará tus respuestas. ¿Deseas continuar?",
        "reset_audit": "Reiniciar Auditoría",
        "reset_warning": "Reiniciar la auditoría eliminará todas las respuestas. ¿Deseas continuar?",
        "contact_info": "Contáctanos en {} o {} para soporte adicional.",
        "high_priority_categories": "Categorías con Alta Prioridad",
        "average_score": "Puntuación Promedio",
        "chart_title": "Fortalezas y Oportunidades del Lugar de Trabajo",
        "score_percent": "Puntuación (%)",
        "question_breakdown": "Análisis Detallado: Perspectivas a Nivel de Pregunta",
        "select_category": "Seleccionar Categoría para Explorar",
        "question_scores_for": "Puntuaciones de Preguntas para",
        "actionable_insights": "Perspectivas Accionables",
        "all_categories_above_70": "¡Todas las categorías obtuvieron más del 70%! Continúa manteniendo estas fortalezas.",
        "summary": "Resumen",
        "results": "Resultados",
        "findings": "Hallazgos",
        "overall_score": "Puntuación General",
        "grade": "Calificación",
        "findings_summary": "Resumen de Hallazgos",
        "findings_summary_text": "{} categorías requieren acción urgente (<{}%), {} necesitan mejoras específicas ({}-{}%). La puntuación general es {}%.",
        "action_required": "Acción {} requerida.",
        "findings_and_suggestions": "Hallazgos y Sugerencias",
        "contact": "Contacto",
        "generating_excel": "Generando Excel...",
        "excel_error": "No se pudo generar el archivo Excel: {}",
        "grade_excellent_desc": "Tu lugar de trabajo demuestra prácticas sobresalientes. ¡Continúa fortaleciendo estas áreas!",
        "grade_good_desc": "Tu lugar de trabajo tiene fortalezas, pero requiere mejoras específicas para alcanzar la excelencia.",
        "grade_needs_improvement_desc": "Se identificaron debilidades moderadas. Prioriza acciones correctivas en áreas críticas.",
        "grade_critical_desc": "Existen problemas significativos que requieren intervención urgente. Considera apoyo externo.",
        "suggestion": "Sugerencia",
        "actionable_charts": "Gráficos Accionables",
        "marketing_message": "¡Transforme su lugar de trabajo con LEAN 2.0 Institute! Colaboramos con usted para implementar soluciones sostenibles que aborden los hallazgos de esta auditoría, promoviendo un entorno laboral ético, inclusivo y productivo. Contáctenos para comenzar hoy mismo.",
        "submit_answers": "Enviar Respuestas",
        "reference_lines": "**Líneas de Referencia:** Discontinua = 50%, Punteada = 70%, Discontinua-Punteada = 85%",
        "show_low_scores": "Mostrar solo preguntas que necesitan mejora (<70%)",
        "actionable": "Accionable"
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
        "report_title": "Your Workplace Wellness Report",
        "download_excel": "Download Excel Report",
        "report_filename_excel": "ethical_workplace_audit_results.xlsx",
        "unanswered_error": "Cannot display results. There are {} unanswered questions. Please complete all questions.",
        "missing_questions": "Missing Questions:",
        "all_answered": "All questions have been answered! Review results below.",
        "response_guide": "Select the description that best represents the situation for each question. The options describe the degree, frequency, or quantity applicable.",
        "language_change_warning": "Changing the language will reset your responses. Do you wish to continue?",
        "reset_audit": "Reset Audit",
        "reset_warning": "Resetting the audit will clear all responses. Do you wish to continue?",
        "contact_info": "Contact us at {} or {} for additional support.",
        "high_priority_categories": "High Priority Categories",
        "average_score": "Average Score",
        "chart_title": "Workplace Strengths and Opportunities",
        "score_percent": "Score (%)",
        "question_breakdown": "Drill Down: Question-Level Insights",
        "select_category": "Select Category to Explore",
        "question_scores_for": "Question Scores for",
        "actionable_insights": "Actionable Insights",
        "all_categories_above_70": "All categories scored above 70%! Continue maintaining these strengths.",
        "summary": "Summary",
        "results": "Results",
        "findings": "Findings",
        "overall_score": "Overall Score",
        "grade": "Grade",
        "findings_summary": "Findings Summary",
        "findings_summary_text": "{} categories require urgent action (<{}%), {} need specific improvements ({}-{}%). Overall score is {}%.",
        "action_required": "{} action required.",
        "findings_and_suggestions": "Findings and Suggestions",
        "contact": "Contact",
        "generating_excel": "Generating Excel...",
        "excel_error": "Failed to generate Excel file: {}",
        "grade_excellent_desc": "Your workplace demonstrates outstanding practices. Continue strengthening these areas!",
        "grade_good_desc": "Your workplace has strengths but requires specific improvements to achieve excellence.",
        "grade_needs_improvement_desc": "Moderate weaknesses identified. Prioritize corrective actions in critical areas.",
        "grade_critical_desc": "Significant issues exist requiring urgent intervention. Consider external support.",
        "suggestion": "Suggestion",
        "actionable_charts": "Actionable Charts",
        "marketing_message": "Transform your workplace with LEAN 2.0 Institute! We partner with you to implement sustainable solutions that address the findings of this audit, fostering an ethical, inclusive, and productive work environment. Contact us to start today.",
        "submit_answers": "Submit Answers",
        "reference_lines": "**Reference Lines:** Dashed = 50%, Dotted = 70%, Dash-Dot = 85%",
        "show_low_scores": "Show only questions needing improvement (<70%)",
        "actionable": "Actionable"
    }
}

# Set page configuration at the top
st.set_page_config(
    page_title=TRANSLATIONS["Español"]["title"],
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cache static data
@st.cache_data
def load_static_data() -> Tuple[Dict, Dict]:
    questions = {
        "Empoderamiento de Empleados": {
            "Español": [
                ("1. ¿Qué porcentaje de sugerencias de empleados presentadas en los últimos 12 meses fueron implementadas con resultados documentados?", "percentage", "Establece un sistema formal para rastrear e implementar sugerencias de empleados con métricas claras."),
                ("2. ¿Cuántos empleados recibieron capacitación en habilidades profesionales en el último año?", "count", "Aumenta las oportunidades de capacitación profesional para todos los empleados."),
                ("3. En los últimos 12 meses, ¿cuántos empleados lideraron proyectos o iniciativas con presupuesto asignado?", "count", "Asigna presupuestos a más iniciativas lideradas por empleados para fomentar la innovación."),
                ("4. ¿Con qué frecuencia se realizan foros formales para que los empleados compartan retroalimentación con la gerencia?", "frequency", "Programa foros mensuales para retroalimentación directa entre empleados y gerencia.")
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
                ("16. ¿Cuántos casos de desafíos personales o profesionales reportados por empleados fueron abordados con planes de acción documentados en el último año?", "count", "Establece procesos formales para abordar desafíos reportados por empleados con planes de acción documentados.")
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
        },
        "Impacto Humano de Procesos Lean": {
            "Español": [
                ("21. ¿Qué porcentaje de sugerencias de mejora de empleados fue implementado con impacto positivo en la carga mental o emocional del trabajo?", "percentage", "Implementa un sistema para priorizar y rastrear sugerencias que reduzcan la carga mental o emocional."),
                ("22. ¿Con qué frecuencia la alta dirección comunica cómo las decisiones lean impactan en el bienestar, seguridad y desarrollo del personal?", "frequency", "Establece comunicaciones regulares de la alta dirección sobre el impacto de decisiones lean en el personal."),
                ("23. ¿Con qué frecuencia se evalúan los efectos de los cambios lean sobre la fatiga, carga cognitiva o sentido de propósito de los empleados?", "frequency", "Realiza evaluaciones trimestrales del impacto de cambios lean en fatiga, carga cognitiva y propósito."),
                ("24. ¿Qué porcentaje de procesos rediseñados eliminó tareas percibidas como sin sentido, humillantes o redundantes por los trabajadores?", "percentage", "Incluye retroalimentación de empleados en el rediseño de procesos para eliminar tareas sin valor."),
                ("25. ¿Qué porcentaje de proyectos lean en los últimos 12 meses incluyó objetivos explícitos de equidad, inclusión o sostenibilidad humana?", "percentage", "Define objetivos de equidad e inclusión en todos los proyectos lean con métricas claras.")
            ],
            "English": [
                ("21. What percentage of employee improvement suggestions were implemented with a positive impact on mental or emotional workload?", "percentage", "Implement a system to prioritize and track suggestions that reduce mental or emotional workload."),
                ("22. How frequently does senior management communicate how lean decisions impact employee well-being, safety, and development?", "frequency", "Establish regular communications from senior management on the impact of lean decisions on employees."),
                ("23. How frequently are the effects of lean changes evaluated on employee fatigue, cognitive load, or sense of purpose?", "frequency", "Conduct quarterly evaluations of lean changes’ impact on fatigue, cognitive load, and purpose."),
                ("24. What percentage of redesigned processes eliminated tasks perceived as meaningless, humiliating, or redundant by workers?", "percentage", "Include employee feedback in process redesigns to eliminate valueless tasks."),
                ("25. What percentage of lean projects in the past 12 months included explicit goals for equity, inclusion, or human sustainability?", "percentage", "Define equity and inclusion goals in all lean projects with clear metrics.")
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

    valid_q_types = set(response_options.keys())
    for cat in questions:
        for lang in questions[cat]:
            for _, q_type, _ in questions[cat][lang]:
                if q_type not in valid_q_types:
                    raise ValueError(f"Invalid question type '{q_type}' in category {cat}, language {lang}")

    return questions, response_options

# Configuration
CONFIG = {
    "contact": {
        "email": os.getenv("CONTACT_EMAIL", "contacto@lean2institute.org"),
        "website": os.getenv("CONTACT_WEBSITE", "https://lean2institute.mystrikingly.com/")
    }
}

# Category mapping
category_mapping = {
    "Español": {
        "Empoderamiento de Empleados": "Empoderamiento de Empleados",
        "Liderazgo Ético": "Liderazgo Ético",
        "Operaciones Centradas en las Personas": "Operaciones Centradas en las Personas",
        "Prácticas Sostenibles y Éticas": "Prácticas Sostenibles y Éticas",
        "Bienestar y Equilibrio": "Bienestar y Equilibrio",
        "Iniciativas Organizacionales Centradas en las Personas": "Iniciativas Organizacionales Centradas en las Personas",
        "Impacto Humano de Procesos Lean": "Impacto Humano de Procesos Lean"
    },
    "English": {
        "Employee Empowerment": "Empoderamiento de Empleados",
        "Ethical Leadership": "Liderazgo Ético",
        "Human-Centered Operations": "Operaciones Centradas en las Personas",
        "Sustainable and Ethical Practices": "Prácticas Sostenibles y Éticas",
        "Well-Being and Balance": "Bienestar y Equilibrio",
        "Human-Centered Organizational Initiatives": "Iniciativas Organizacionales Centradas en las Personas",
        "Human Impact of Lean Processes": "Impacto Humano de Procesos Lean"
    }
}

# Load static data before initializing session state
questions, response_options = load_static_data()

# Initialize session state
def initialize_session_state():
    defaults = {
        "language": "Español",
        "responses": {},
        "language_changed": False,
        "reset_confirmed": False,
        "report_id": str(uuid.uuid4()),
        "last_scroll_position": 0,
        "submit_clicked": False
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    if st.session_state.language not in ["Español", "English"]:
        st.session_state.language = "Español"

    # Ensure responses is initialized as a dictionary if empty or invalid
    if not isinstance(st.session_state.responses, dict):
        st.session_state.responses = {
            cat: [None] * len(questions[cat]["Español"]) for cat in questions
        }
    else:
        for cat in questions:
            expected_len = len(questions[cat]["Español"])
            # If category is missing or not a list, initialize it
            if cat not in st.session_state.responses or not isinstance(st.session_state.responses[cat], list):
                st.session_state.responses[cat] = [None] * expected_len
            # Adjust length if necessary
            elif len(st.session_state.responses[cat]) != expected_len:
                current = st.session_state.responses[cat]
                st.session_state.responses[cat] = (
                    current[:expected_len] + [None] * (expected_len - len(current))
                ) if len(current) < expected_len else current[:expected_len]

# Call initialize_session_state
initialize_session_state()

# Load CSS
def load_css():
    default_css = """
        .main-container { overflow-anchor: none; min-height: 100vh; }
        .main-title { font-size: 2.5rem; color: #1E88E5; text-align: center; }
        .section-title { font-size: 1.8rem; color: #424242; margin-bottom: 1rem; }
        .card-modern { background: #FFFFFF; border-radius: 8px; padding: 1.5rem; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 1rem; min-height: 200px; }
        .question-container { display: flex; align-items: center; margin-bottom: 0.5rem; }
        .question-text { font-weight: 500; color: #212121; margin-right: 0.5rem; }
        .required { color: #D32F2F; font-weight: bold; }
        .tooltip { position: relative; display: inline-block; }
        .tooltip-icon { background: #1E88E5; color: white; border-radius: 50%; width: 20px; height: 20px; text-align: center; line-height: 20px; cursor: help; }
        .tooltip-text { visibility: hidden; width: 200px; background: #424242; color: #FFFFFF; text-align: center; border-radius: 6px; padding: 5px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -100px; opacity: 0; transition: opacity 0.3s; }
        .tooltip:hover .tooltip-text { visibility: visible; opacity: 1; }
        .sticky-nav { position: sticky; bottom: 0; background: #F5F5F5; padding: 1rem; border-radius: 8px; box-shadow: 0 -2px 4px rgba(0,0,0,0.1); z-index: 10; }
        .grade-excellent { background: #43A047; color: #FFFFFF; padding: 0.5rem; border-radius: 4px; }
        .grade-good { background: #FFD54F; color: #212121; padding: 0.5rem; border-radius: 4px; }
        .grade-needs-improvement { background: #FF9800; color: #FFFFFF; padding: 0.5rem; border-radius: 4px; }
        .grade-critical { background: #D32F2F; color: #FFFFFF; padding: 0.5rem; border-radius: 4px; }
        .alert { padding: 1rem; border-radius: 4px; margin-bottom: 1rem; min-height: 60px; }
        .alert-info { background: #E3F2FD; color: #1E88E5; }
        .alert-success { background: #E8F5E9; color: #43A047; }
        .alert-warning { background: #FFF3E0; color: #FF9800; }
        .badge { background: #1E88E5; color: #FFFFFF; padding: 0.5rem 1rem; border-radius: 16px; display: inline-block; margin: 1rem 0; }
        .contact-info { margin-top: 1rem; font-size: 0.9rem; color: #424242; }
        @media (max-width: 768px) { 
            .main-title { font-size: 2rem; } 
            .section-title { font-size: 1.5rem; } 
            .card-modern { padding: 1rem; } 
        }
    """
    st.markdown(f"<style>{default_css}</style>", unsafe_allow_html=True)

# JavaScript for scroll preservation
st.markdown("""
    <script>
        document.addEventListener("DOMContentLoaded", function() {
            const scrollPos = sessionStorage.getItem("scrollPosition");
            if (scrollPos) {
                window.scrollTo(0, parseInt(scrollPos));
            }
        });
        document.addEventListener("change", function() {
            sessionStorage.setItem("scrollPosition", window.scrollY);
        });
        document.addEventListener("click", function() {
            sessionStorage.setItem("scrollPosition", window.scrollY);
        });
        function scrollToCategory(categoryId) {
            const element = document.getElementById(categoryId);
            if (element) {
                element.scrollIntoView({ behavior: 'smooth', block: 'start' });
                sessionStorage.setItem("scrollPosition", element.offsetTop);
            }
        }
    </script>
""", unsafe_allow_html=True)

# Helper functions
def sanitize_input(text: str) -> str:
    return re.sub(r'[<>]', '', text)

def get_grade(score: float) -> Tuple[str, str, str]:
    lang = st.session_state.language
    if score >= SCORE_THRESHOLDS["GOOD"]:
        return (
            "Excelente" if lang == "Español" else "Excellent",
            TRANSLATIONS[lang]["grade_excellent_desc"],
            "grade-excellent"
        )
    elif score >= SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
        return (
            "Bueno" if lang == "Español" else "Good",
            TRANSLATIONS[lang]["grade_good_desc"],
            "grade-good"
        )
    elif score >= SCORE_THRESHOLDS["CRITICAL"]:
        return (
            "Necesita Mejora" if lang == "Español" else "Needs Improvement",
            TRANSLATIONS[lang]["grade_needs_improvement_desc"],
            "grade-needs-improvement"
        )
    else:
        return (
            "Crítico" if lang == "Español" else "Critical",
            TRANSLATIONS[lang]["grade_critical_desc"],
            "grade-critical"
        )

def update_language():
    if st.session_state.language_select != st.session_state.language:
        if any(any(score is not None for score in scores) for scores in st.session_state.responses.values()):
            if not st.session_state.get("language_change_confirmed", False):
                st.session_state.language_changed = True
                return
        st.session_state.language = st.session_state.language_select
        st.session_state.responses = {
            cat: [None] * len(questions[cat][st.session_state.language]) for cat in questions
        }
        st.session_state.language_changed = False
        st.session_state.language_change_confirmed = False
        st.session_state.report_id = str(uuid.uuid4())
        st.session_state.submit_clicked = False

def reset_audit():
    if st.session_state.get("reset_confirmed", False):
        st.session_state.responses = {
            cat: [None] * len(questions[cat]["Español"]) for cat in questions
        }
        st.session_state.reset_confirmed = False
        st.session_state.report_id = str(uuid.uuid4())
        st.session_state.submit_clicked = False

# Sidebar
with st.sidebar:
    st.markdown('<section class="sidebar-container" role="navigation" aria-label="Audit Navigation">', unsafe_allow_html=True)
    st.image("assets/FOBO2.png", width=250, caption="LEAN 2.0 Institute")
    # Contact Information
    email_link = f'<a href="mailto:{CONFIG["contact"]["email"]}">{CONFIG["contact"]["email"]}</a>'
    website_link = f'<a href="{CONFIG["contact"]["website"]}">{CONFIG["contact"]["website"]}</a>'
    contact_text = TRANSLATIONS[st.session_state.language]["contact_info"].format(email_link, website_link)
    st.markdown(f'<div class="contact-info">{contact_text}</div>', unsafe_allow_html=True)
    st.selectbox(
        "Idioma / Language",
        ["Español", "English"],
        key="language_select",
        on_change=update_language,
        help="Selecciona tu idioma preferido / Select your preferred language"
    )
    if st.session_state.get("language_changed", False):
        st.warning(TRANSLATIONS[st.session_state.language]["language_change_warning"])
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Confirmar / Confirm", key="confirm_language_change", type="primary"):
                st.session_state.language_change_confirmed = True
                update_language()
        with col2:
            if st.button("Cancelar / Cancel", key="cancel_language_change"):
                st.session_state.language_select = st.session_state.language
                st.session_state.language_changed = False
    st.markdown('<h2 class="sidebar-title">Navegación / Navigation</h2>', unsafe_allow_html=True)
    display_categories = list(category_mapping[st.session_state.language].keys())
    for i, display_cat in enumerate(display_categories):
        category_id = f"category_{i}"
        if st.button(
            display_cat,
            key=f"nav_{i}",
            use_container_width=True,
            type="secondary"
        ):
            st.markdown(f'<script>scrollToCategory("{category_id}")</script>', unsafe_allow_html=True)
    if st.button(TRANSLATIONS[st.session_state.language]["reset_audit"], key="reset_audit_button", type="secondary"):
        st.session_state.reset_confirmed = True
        st.warning(TRANSLATIONS[st.session_state.language]["reset_warning"])
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Confirmar / Confirm", key="confirm_reset"):
                reset_audit()
        with col2:
            if st.button("Cancelar / Cancel", key="cancel_reset"):
                st.session_state.reset_confirmed = False
    st.markdown('</section>', unsafe_allow_html=True)

# Main content
with st.container():
    st.markdown('<section class="main-container" role="main">', unsafe_allow_html=True)
    st.markdown(f'<h1 class="main-title">{TRANSLATIONS[st.session_state.language]["header"]}</h1>', unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="intro-content">
            {TRANSLATIONS[st.session_state.language]["response_guide"]}
        </div>
        """,
        unsafe_allow_html=True
    )

    # Display all categories and questions
    for idx, display_category in enumerate(display_categories):
        category = category_mapping[st.session_state.language][display_category]
        category_id = f"category_{idx}"
        with st.container():
            st.markdown(f'<div id="{category_id}" class="card-modern" role="region" aria-label="Category {display_category} Questions">', unsafe_allow_html=True)
            st.markdown(f'<h2 class="section-title">{display_category}</h2>', unsafe_allow_html=True)
            for q_idx, (q, q_type, _) in enumerate(questions[category][st.session_state.language]):
                with st.container():
                    is_unanswered = st.session_state.responses[category][q_idx] is None
                    st.markdown(
                        f"""
                        <div class="question-container">
                            <label class="question-text" for="{category}_{q_idx}">
                                {sanitize_input(q)} {'<span class="required" aria-label="Required">*</span>' if is_unanswered else ''}
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
                    radio_key = f"{category}_{q_idx}_{st.session_state.report_id}"
                    selected_description = st.radio(
                        "",
                        descriptions,
                        key=radio_key,
                        horizontal=False,
                        help=response_options[q_type][st.session_state.language]['tooltip'],
                        label_visibility="hidden"
                    )
                    score_idx = descriptions.index(selected_description)
                    st.session_state.responses[category][q_idx] = scores[score_idx]
            st.markdown('</div>', unsafe_allow_html=True)

    # Submit Answers button
    if st.button(
        TRANSLATIONS[st.session_state.language]["submit_answers"],
        key="submit_answers",
        type="primary",
        use_container_width=True
    ):
        st.session_state.submit_clicked = True

    # Check audit completion
    audit_complete = all(
        all(score is not None for score in scores)
        for scores in st.session_state.responses.values()
    )

    # Handle submit logic
    if st.session_state.submit_clicked:
        if not audit_complete:
            unanswered_questions = []
            question_counter = 1
            for cat in questions.keys():
                for i, (q, _, _) in enumerate(questions[cat][st.session_state.language]):
                    if st.session_state.responses[cat][i] is None:
                        display_cat = next(
                            k for k, v in category_mapping[st.session_state.language].items() if v == cat
                        )
                        truncated_q = (
                            q[:QUESTION_TRUNCATE_LENGTH] + ("..." if len(q) > QUESTION_TRUNCATE_LENGTH else "")
                        )
                        unanswered_questions.append(
                            f"{display_cat}: Pregunta {question_counter} - {truncated_q}"
                        )
                    question_counter += 1
            st.error(
                TRANSLATIONS[st.session_state.language]["unanswered_error"].format(
                    len(unanswered_questions)
                ),
                icon="⚠️"
            )
            st.markdown(
                f"""
                <div class="alert alert-warning" role="alert">
                    <strong>{TRANSLATIONS[st.session_state.language]["missing_questions"]}</strong>
                    <ul>
                        {"".join([f"<li>{sanitize_input(q)}</li>" for q in unanswered_questions])}
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
            # Results section
            st.markdown(f'<div class="card-modern report-section" role="region" aria-label="{TRANSLATIONS[st.session_state.language]["report_title"]}">', unsafe_allow_html=True)
            st.markdown(f'<h2 class="section-title">{TRANSLATIONS[st.session_state.language]["report_title"]}</h2>', unsafe_allow_html=True)
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
            st.markdown('<h3 class="subsection-title">Resumen Ejecutivo</h3>', unsafe_allow_html=True)
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
                    TRANSLATIONS[st.session_state.language]["high_priority_categories"],
                    len(df[df[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
                )
            with col3:
                st.metric(
                    TRANSLATIONS[st.session_state.language]["average_score"],
                    f"{overall_score:.1f}%"
                )

            # Color-coded dataframe
            def color_percent(val):
                color = CHART_COLORS[0] if val < SCORE_THRESHOLDS["CRITICAL"] else CHART_COLORS[1] if val < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else CHART_COLORS[2]
                return f'background-color: {color}; color: white;'

            st.dataframe(
                df.style.applymap(color_percent, subset=[TRANSLATIONS[st.session_state.language]["percent"]]).format({TRANSLATIONS[st.session_state.language]["percent"]: "{:.1f}%"}),
                use_container_width=True
            )

            # Bar chart with improvements
            df_display = df.copy()
            df_display.index = [next(k for k, v in category_mapping[st.session_state.language].items() if v == idx) for idx in df.index]
            df_display = df_display.sort_values(by=TRANSLATIONS[st.session_state.language]["percent"], ascending=True)  # Sort by score ascending
            fig = px.bar(
                df_display.reset_index(),
                y="index",
                x=TRANSLATIONS[st.session_state.language]["percent"],
                orientation='h',
                title=TRANSLATIONS[st.session_state.language]["chart_title"],
                labels={
                    "index": TRANSLATIONS[st.session_state.language]["category"],
                    TRANSLATIONS[st.session_state.language]["percent"]: TRANSLATIONS[st.session_state.language]["score_percent"]
                },
                color=TRANSLATIONS[st.session_state.language]["percent"],
                color_continuous_scale=CHART_COLORS,
                range_x=[0, 100],
                height=CHART_HEIGHT
            )
            fig.update_layout(
                showlegend=False,
                title_x=0.5,
                xaxis_title=TRANSLATIONS[st.session_state.language]["score_percent"],
                yaxis_title=TRANSLATIONS[st.session_state.language]["category"],
                coloraxis_showscale=False
            )
            # Add reference lines
            fig.add_vline(x=50, line_dash="dash", line_color="black")
            fig.add_vline(x=70, line_dash="dot", line_color="black")
            fig.add_vline(x=85, line_dash="dashdot", line_color="black")
            st.plotly_chart(fig, use_container_width=True)
            st.markdown(TRANSLATIONS[st.session_state.language]["reference_lines"], unsafe_allow_html=True)

            # Question-level breakdown with improvements
            with st.expander(TRANSLATIONS[st.session_state.language]["question_breakdown"]):
                selected_display_category = st.selectbox(
                    TRANSLATIONS[st.session_state.language]["select_category"],
                    display_categories,
                    key="category_explore"
                )
                selected_category = category_mapping[st.session_state.language][selected_display_category]
                question_scores = pd.DataFrame({
                    TRANSLATIONS[st.session_state.language]["question"]: [q for q, _, _ in questions[selected_category][st.session_state.language]],
                    TRANSLATIONS[st.session_state.language]["score"]: st.session_state.responses[selected_category]
                })
                show_low_scores = st.checkbox(TRANSLATIONS[st.session_state.language]["show_low_scores"], key="show_low_scores")
                if show_low_scores:
                    filtered_scores = question_scores[question_scores[TRANSLATIONS[st.session_state.language]["score"]] < 70]
                    title_suffix = " (Below 70%)"
                else:
                    filtered_scores = question_scores
                    title_suffix = ""
                fig_questions = px.bar(
                    filtered_scores,
                    x=TRANSLATIONS[st.session_state.language]["score"],
                    y=TRANSLATIONS[st.session_state.language]["question"],
                    orientation='h',
                    title=f"{TRANSLATIONS[st.session_state.language]['question_scores_for']} {selected_display_category}{title_suffix}",
                    color=TRANSLATIONS[st.session_state.language]["score"],
                    color_continuous_scale=CHART_COLORS,
                    range_x=[0, 100],
                    height=300 + len(filtered_scores) * 50
                )
                fig_questions.update_layout(
                    showlegend=False,
                    title_x=0.5,
                    xaxis_title=TRANSLATIONS[st.session_state.language]["score_percent"],
                    yaxis_title=TRANSLATIONS[st.session_state.language]["question"],
                    coloraxis_showscale=False
                )
                # Add reference lines
                fig_questions.add_vline(x=50, line_dash="dash", line_color="black")
                fig_questions.add_vline(x=70, line_dash="dot", line_color="black")
                fig_questions.add_vline(x=85, line_dash="dashdot", line_color="black")
                st.plotly_chart(fig_questions, use_container_width=True)

            # Actionable insights
            with st.expander(TRANSLATIONS[st.session_state.language]["actionable_insights"]):
                insights = []
                for cat in questions.keys():
                    display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                    score = df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]]
                    if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        insights.append(
                            f"**{display_cat}** scored {score:.1f}% ({TRANSLATIONS[st.session_state.language]['high_priority'] if score < SCORE_THRESHOLDS['CRITICAL'] else TRANSLATIONS[st.session_state.language]['medium_priority']}). Focus on immediate improvements."
                        )
                if insights:
                    st.markdown("<div class='alert alert-info'>" + "<br>".join(insights) + "</div>", unsafe_allow_html=True)
                else:
                    st.markdown(
                        f"<div class='alert alert-success'>{TRANSLATIONS[st.session_state.language]['all_categories_above_70']}</div>",
                        unsafe_allow_html=True
                    )

            # Download Excel report
            def generate_excel_report() -> io.BytesIO:
                excel_output = io.BytesIO()
                with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    bold = workbook.add_format({'bold': True})
                    percent_format = workbook.add_format({'num_format': '0.0%'})
                    wrap_format = workbook.add_format({'text_wrap': True})
                    border_format = workbook.add_format({'border': 1})
                    header_format = workbook.add_format({'bold': True, 'bg_color': '#1E88E5', 'color': 'white', 'border': 1})

                    # Single Actionable Worksheet
                    worksheet = workbook.add_worksheet(TRANSLATIONS[st.session_state.language]["actionable"])
                    row = 0

                    # Report Title and Date
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["report_title"], bold)
                    row += 1
                    worksheet.write(row, 0, f"Date: {REPORT_DATE}", bold)
                    row += 2

                    # Summary Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["summary"], bold)
                    row += 1
                    critical_count = len(df[df[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
                    improvement_count = len(df[(df[TRANSLATIONS[st.session_state.language]["percent"]] >= SCORE_THRESHOLDS["CRITICAL"]) & (df[TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"])])
                    summary_df = pd.DataFrame({
                        TRANSLATIONS[st.session_state.language]["overall_score"]: [f"{overall_score:.1f}%"],
                        TRANSLATIONS[st.session_state.language]["grade"]: [grade],
                        TRANSLATIONS[st.session_state.language]["findings_summary"]: [
                            TRANSLATIONS[st.session_state.language]["findings_summary_text"].format(
                                critical_count,
                                SCORE_THRESHOLDS["CRITICAL"],
                                improvement_count,
                                SCORE_THRESHOLDS["CRITICAL"],
                                SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1,
                                overall_score
                            )
                        ]
                    })
                    summary_df.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], index=False, startrow=row)
                    for col_num, value in enumerate(summary_df.columns.values):
                        worksheet.write(row, col_num, value, header_format)
                    row += len(summary_df) + 2

                    # Contact Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["contact"], bold)
                    row += 1
                    contact_df = pd.DataFrame({
                        "Contact Method": ["Email", "Website"],
                        "Details": [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
                    })
                    contact_df.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], index=False, startrow=row)
                    for col_num, value in enumerate(contact_df.columns.values):
                        worksheet.write(row, col_num, value, header_format)
                    row += len(contact_df) + 1
                    worksheet.write(row, 0, "¡Trabajemos juntos!|Let's work together!", bold)
                    row += 1
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["marketing_message"], wrap_format)
                    row += 2

                    # Results Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["results"], bold)
                    row += 1
                    df_display.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], float_format="%.1f", startrow=row+1)
                    for col_num, value in enumerate(df_display.columns.values):
                        worksheet.write(row+1, col_num + 1, value, header_format)
                    worksheet.write(row+1, 0, TRANSLATIONS[st.session_state.language]["category"], header_format)
                    row += len(df_display) + 3

                    # Findings Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["findings"], bold)
                    row += 1
                    findings_data = []
                    for cat in questions.keys():
                        display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                        if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                            findings_data.append([
                                display_cat,
                                f"{df.loc[cat, TRANSLATIONS[st.session_state.language]['percent']]:.1f}%",
                                TRANSLATIONS[st.session_state.language]["high_priority"] if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else TRANSLATIONS[st.session_state.language]["medium_priority"],
                                TRANSLATIONS[st.session_state.language]["action_required"].format(
                                    "Urgent" if df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else "Specific"
                                )
                            ])
                            for idx, score in enumerate(st.session_state.responses[cat]):
                                if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                                    question, _, rec = questions[cat][st.session_state.language][idx]
                                    findings_data.append([
                                        "", "", "", f"{TRANSLATIONS[st.session_state.language]['question']}: {question[:50]}... - {TRANSLATIONS[st.session_state.language]['suggestion']}: {rec}"
                                    ])
                    findings_df = pd.DataFrame(
                        findings_data,
                        columns=[
                            TRANSLATIONS[st.session_state.language]["category"],
                            TRANSLATIONS[st.session_state.language]["score"],
                            TRANSLATIONS[st.session_state.language]["priority"],
                            TRANSLATIONS[st.session_state.language]["findings_and_suggestions"]
                        ]
                    )
                    findings_df.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], index=False, startrow=row)
                    for col_num, value in enumerate(findings_df.columns.values):
                        worksheet.write(row, col_num, value, header_format)
                    row += len(findings_df) + 2

                    # Actionable Insights Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["actionable_insights"], bold)
                    row += 1
                    insights_data = []
                    for cat in questions.keys():
                        display_cat = next(k for k, v in category_mapping[st.session_state.language].items() if v == cat)
                        score = df.loc[cat, TRANSLATIONS[st.session_state.language]["percent"]]
                        if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                            insights_data.append([display_cat, f"{score:.1f}%", "Focus on immediate improvements."])
                    insights_df = pd.DataFrame(
                        insights_data,
                        columns=[
                            TRANSLATIONS[st.session_state.language]["category"],
                            TRANSLATIONS[st.session_state.language]["score"],
                            TRANSLATIONS[st.session_state.language]["actionable_insights"]
                        ]
                    )
                    insights_df.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], index=False, startrow=row)
                    for col_num, value in enumerate(insights_df.columns.values):
                        worksheet.write(row, col_num, value, header_format)
                    row += len(insights_df) + 2

                    # Actionable Charts Section
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["actionable_charts"], bold)
                    row += 1
                    chart_data = df_display[[TRANSLATIONS[st.session_state.language]["percent"]]].reset_index()
                    chart_data.to_excel(writer, sheet_name=TRANSLATIONS[st.session_state.language]["actionable"], startrow=row, index=False)
                    worksheet.write(row, 0, TRANSLATIONS[st.session_state.language]["category"], header_format)
                    worksheet.write(row, 1, TRANSLATIONS[st.session_state.language]["score_percent"], header_format)
                    bar_chart = workbook.add_chart({'type': 'bar'})
                    bar_chart.add_series({
                        'name': TRANSLATIONS[st.session_state.language]["score_percent"],
                        'categories': f"='{TRANSLATIONS[st.session_state.language]['actionable']}'!$A${row+1}:$A${row+len(chart_data)}",
                        'values': f"='{TRANSLATIONS[st.session_state.language]['actionable']}'!$B${row+1}:$B${row+len(chart_data)}",
                        'fill': {'color': '#1E88E5'}
                    })
                    bar_chart.set_title({'name': TRANSLATIONS[st.session_state.language]["chart_title"]})
                    bar_chart.set_x_axis({'name': TRANSLATIONS[st.session_state.language]["score_percent"], 'min': 0, 'max': 100})
                    bar_chart.set_y_axis({'name': TRANSLATIONS[st.session_state.language]["category"]})
                    worksheet.insert_chart(f'D{row+1}', bar_chart)
                    row += len(chart_data) + 2

                    # Set column widths
                    worksheet.set_column('A:A', 30)
                    worksheet.set_column('B:B', 15)
                    worksheet.set_column('C:C', 20)
                    worksheet.set_column('D:D', 80, wrap_format)

                excel_output.seek(0)
                return excel_output

            with st.spinner(TRANSLATIONS[st.session_state.language]["generating_excel"]):
                try:
                    excel_file = generate_excel_report()
                    st.download_button(
                        label=TRANSLATIONS[st.session_state.language]["download_excel"],
                        data=excel_file,
                        file_name=TRANSLATIONS[st.session_state.language]["report_filename_excel"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel",
                        use_container_width=True,
                        type="primary"
                    )
                except Exception as e:
                    st.error(TRANSLATIONS[st.session_state.language]["excel_error"].format(str(e)), icon="❌")

            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</section>', unsafe_allow_html=True)

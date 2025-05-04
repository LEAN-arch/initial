import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import base64
import io
import numpy as np

# Set page configuration as the first Streamlit command
st.set_page_config(page_title="Ethical Lean Workplace Audit", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for a vibrant, engaging interface
st.markdown("""
    <style>
        body {
            font-family: 'Arial', sans-serif;
            background-color: #f0f4f8;
        }
        .main {
            background-color: #ffffff;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 6px 12px rgba(0, 0, 0, 0.1);
            max-width: 1200px;
            margin: auto;
        }
        .stButton>button {
            background-color: #007bff;
            color: white;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: bold;
            transition: background-color 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #0056b3;
        }
        .stRadio>label {
            background-color: #e9f7ff;
            padding: 10px;
            border-radius: 8px;
            margin: 8px 0;
            font-size: 1.1em;
        }
        .header {
            color: #007bff;
            font-size: 2.8em;
            text-align: center;
            margin-bottom: 20px;
            font-weight: bold;
        }
        .subheader {
            color: #333;
            font-size: 1.8em;
            margin-top: 25px;
            text-align: center;
        }
        .sidebar .sidebar-content {
            background-color: #ffffff;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        .download-link {
            color: #007bff;
            font-weight: bold;
            text-decoration: none;
            font-size: 1.1em;
        }
        .download-link:hover {
            color: #0056b3;
            text-decoration: underline;
        }
        .motivation {
            color: #28a745;
            font-size: 1.2em;
            text-align: center;
            margin: 20px 0;
            font-style: italic;
        }
        .badge {
            background-color: #28a745;
            color: white;
            padding: 10px 20px;
            border-radius: 20px;
            font-size: 1.2em;
            text-align: center;
            margin: 20px auto;
            display: block;
            width: fit-content;
        }
        .insights {
            background-color: #e9f7ff;
            padding: 15px;
            border-radius: 10px;
            margin: 20px 0;
        }
        .grade {
            font-size: 1.5em;
            font-weight: bold;
            text-align: center;
            margin: 20px 0;
            padding: 10px;
            border-radius: 8px;
        }
        .grade-excellent { background-color: #28a745; color: white; }
        .grade-good { background-color: #ffd700; color: black; }
        .grade-needs-improvement { background-color: #ff4d4d; color: white; }
        .grade-critical { background-color: #b30000; color: white; }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'language' not in st.session_state:
    st.session_state.language = "English"
if 'responses' not in st.session_state:
    st.session_state.responses = {}
if 'current_category' not in st.session_state:
    st.session_state.current_category = 0
if 'prev_language' not in st.session_state:
    st.session_state.prev_language = st.session_state.language

# Bilingual support
LANG = st.sidebar.selectbox(
    "Language / Idioma", 
    ["English", "Español"], 
    help="Select your preferred language / Selecciona tu idioma preferido"
)

# Reset session state if language changes
if LANG != st.session_state.prev_language:
    st.session_state.current_category = 0
    st.session_state.responses = {}
    st.session_state.prev_language = LANG

st.session_state.language = LANG

# Likert scale labels
labels = {
    "English": ["Not at All", "Rarely", "Sometimes", "Often", "Always"],
    "Español": ["Nunca", "Raramente", "A veces", "A menudo", "Siempre"]
}

# Audit categories and questions
questions = {
    "Empowering Employees": {
        "English": [
            "How often are employee suggestions implemented to improve workplace processes or culture?",
            "Does the workplace provide regular workshops or training to develop employee skills and confidence?",
            "How frequently do employees have opportunities to lead projects or initiatives that impact their team?",
            "How effectively does the workplace encourage open dialogue between employees and management?"
        ],
        "Español": [
            "¿Con qué frecuencia se implementan las sugerencias de los empleados para mejorar los procesos o la cultura laboral?",
            "¿Proporciona el lugar de trabajo talleres o capacitaciones regulares para desarrollar las habilidades y la confianza de los empleados?",
            "¿Con qué frecuencia tienen los empleados oportunidades de liderar proyectos o iniciativas que impacten a su equipo?",
            "¿Qué tan efectivamente fomenta el lugar de trabajo el diálogo abierto entre empleados y la gerencia?"
        ]
    },
    "Ethical Leadership": {
        "English": [
            "How consistently do leaders share clear updates on decisions affecting employees?",
            "Does leadership actively involve employees in shaping workplace policies or ethical standards?",
            "How often do leaders recognize and reward ethical behavior or contributions to team well-being?"
        ],
        "Español": [
            "¿Con qué consistencia comparten los líderes actualizaciones claras sobre decisiones que afectan a los empleados?",
            "¿Involucra activamente el liderazgo a los empleados en la formación de políticas laborales o estándares éticos?",
            "¿Con qué frecuencia reconocen y recompensan los líderes el comportamiento ético o las contribuciones al bienestar del equipo?"
        ]
    },
    "Human-Centered Operations": {
        "English": [
            "How effectively do lean processes incorporate employee feedback to reduce unnecessary workload?",
            "Does the workplace regularly review operational practices to ensure they support employee well-being?",
            "How often are employees trained to use lean tools in ways that enhance collaboration and respect?"
        ],
        "Español": [
            "¿Qué tan efectivamente incorporan los procesos lean la retroalimentación de los empleados para reducir la carga de trabajo innecesaria?",
            "¿Revisa regularmente el lugar de trabajo las prácticas operativas para asegurar que apoyen el bienestar de los empleados?",
            "¿Con qué frecuencia se capacita a los empleados para usar herramientas lean de manera que fomenten la colaboración y el respeto?"
        ]
    },
    "Sustainable and Ethical Practices": {
        "English": [
            "Does the workplace actively reduce waste (e.g., energy, materials) through lean initiatives?",
            "How consistently does the workplace partner with suppliers who prioritize fair labor and environmental standards?",
            "How often are employees involved in sustainability projects that benefit the workplace or community?"
        ],
        "Español": [
            "¿Reduce activamente el lugar de trabajo el desperdicio (por ejemplo, energía, materiales) a través de iniciativas lean?",
            "¿Con qué consistencia se asocia el lugar de trabajo con proveedores que priorizan estándares laborales justos y ambientales?",
            "¿Con qué frecuencia participan los empleados en proyectos de sostenibilidad que benefician al lugar de trabajo o la comunidad?"
        ]
    },
    "Well-Being and Balance": {
        "English": [
            "How consistently does the workplace offer resources (e.g., counseling, flexible schedules) to manage stress and workload?",
            "Does the workplace conduct regular check-ins to assess and address employee burnout or fatigue?",
            "How effectively does the workplace promote a culture where employees feel safe to express personal or professional challenges?"
        ],
        "Español": [
            "¿Con qué consistencia ofrece el lugar de trabajo recursos (por ejemplo, asesoramiento, horarios flexibles) para gestionar el estrés y la carga de trabajo?",
            "¿Realiza el lugar de trabajo revisiones regulares para evaluar y abordar el agotamiento o la fatiga de los empleados?",
            "¿Qué tan efectivamente promueve el lugar de trabajo una cultura donde los empleados se sientan seguros para expresar desafíos personales o profesionales?"
        ]
    }
}

# Validate question counts
for cat in questions:
    if len(questions[cat]["English"]) != len(questions[cat]["Español"]):
        st.error(f"Question count mismatch in category {cat} between English and Spanish.")
        st.stop()

# Initialize responses for the current language if not already done
if not st.session_state.responses or len(st.session_state.responses) != len(questions):
    st.session_state.responses = {cat: [0] * len(questions[cat][LANG]) for cat in questions}

# Welcome message
st.markdown(
    '<div class="header">Shape an Ethical & Human-Centered Workplace!</div>' if LANG == "English" else 
    '<div class="header">¡Construye un Lugar de Trabajo Ético y Centrado en las Personas!</div>', 
    unsafe_allow_html=True
)
st.markdown(
    "Your input is vital to creating a workplace that’s ethical, lean, and truly human-centered. Let’s make a difference together!" if LANG == "English" else
    "Tu aporte es clave para crear un lugar de trabajo ético, lean y verdaderamente centrado en las personas. ¡Hagamos la diferencia juntos!", 
    unsafe_allow_html=True
)

# Progress bar with milestones
categories = list(questions.keys())
progress = min((st.session_state.current_category + 1) / len(categories), 1.0)
st.progress(progress)
if progress >= 0.5 and st.session_state.current_category < len(categories) - 1:
    st.markdown(
        '<div class="motivation">Halfway there! Your insights are shaping a better workplace!</div>' if LANG == "English" else
        '<div class="motivation">¡A mitad de camino! ¡Tus ideas están moldeando un mejor lugar de trabajo!</div>', 
        unsafe_allow_html=True
    )
elif progress == 1.0:
    st.markdown(
        '<div class="motivation">Almost done! Just one step left to complete your impact!</div>' if LANG == "English" else
        '<div class="motivation">¡Casi listo! ¡Solo un paso más para completar tu impacto!</div>', 
        unsafe_allow_html=True
    )

# Category navigation
st.sidebar.subheader("Your Progress" if LANG == "English" else "Tu Progreso")
category_index = st.sidebar.slider(
    "Select Category / Seleccionar Categoría",
    0, len(categories) - 1, st.session_state.current_category,
    disabled=False,
    help="Choose a category to provide your insights / Elige una categoría para compartir tus ideas"
)
st.session_state.current_category = category_index
category_index = min(category_index, len(categories) - 1)
category = categories[category_index]

# Collect responses for the current category
st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
for idx, q in enumerate(questions[category][LANG]):
    score = st.radio(
        f"{q}",
        list(range(1, 6)),
        format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
        key=f"{category}_{idx}",
        horizontal=True,
        help="Your response helps build a more ethical and human-centered workplace!" if LANG == "English" else
             "¡Tu respuesta ayuda a construir un lugar de trabajo más ético y centrado en las personas!"
    )
    st.session_state.responses[category][idx] = score

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Previous Category" if LANG == "English" else "Categoría Anterior", disabled=category_index == 0):
        st.session_state.current_category = max(category_index - 1, 0)
with col2:
    if st.button("Next Category" if LANG == "English" else "Siguiente Categoría", disabled=category_index == len(categories) - 1):
        st.session_state.current_category = min(category_index + 1, len(categories) - 1)

# Grading matrix function
def get_grade(score):
    if score >= 85:
        return "Excellent", "Your workplace excels in ethical, lean, and human-centered practices. Continue maintaining these strengths!", "grade-excellent"
    elif score >= 70:
        return "Good", "Your workplace is strong but has room to refine specific areas for optimal performance. Consider targeted improvements.", "grade-good"
    elif score >= 50:
        return "Needs Improvement", "Your workplace requires targeted interventions to address moderate weaknesses. Prioritize action in low-scoring areas.", "grade-needs-improvement"
    else:
        return "Critical", "Your workplace has significant issues requiring urgent, comprehensive action. Engage experts for a transformation plan.", "grade-critical"

# Generate report
if st.button("Generate Report" if LANG == "English" else "Generar Informe", key="generate_report"):
    st.markdown(
        f'<div class="subheader">{"Your Workplace Impact Report" if LANG == "English" else "Tu Informe de Impacto en el Lugar de Trabajo"}</div>',
        unsafe_allow_html=True
    )
    st.markdown(
        '<div class="badge">🏆 Audit Completed! Thank you for shaping an ethical workplace!</div>' if LANG == "English" else
        '<div class="badge">🏆 ¡Auditoría Completada! ¡Gracias por construir un lugar de trabajo ético!</div>', 
        unsafe_allow_html=True
    )
    
    # Calculate scores
    results = {cat: sum(scores) for cat, scores in st.session_state.responses.items()}
    df = pd.DataFrame.from_dict(results, orient="index", columns=["Score"])
    df["Percent"] = [((score / (len(questions[cat][LANG]) * 5)) * 100) for cat, score in results.items()]
    df["Priority"] = df["Percent"].apply(lambda x: "High" if x < 50 else "Medium" if x < 70 else "Low")
    
    # Overall score and grade
    overall_score = df["Percent"].mean()
    grade, grade_description, grade_class = get_grade(overall_score)
    st.markdown(
        f'<div class="grade {grade_class}">Overall Workplace Grade: {grade} ({overall_score:.1f}%)</div>',
        unsafe_allow_html=True
    )
    st.markdown(grade_description, unsafe_allow_html=True)

    # Color-coded dataframe
    def color_percent(val):
        color = '#ff4d4d' if val < 50 else '#ffd700' if val < 70 else '#28a745'
        return f'background-color: {color}; color: white;'
    
    st.markdown(
        "Scores below 50% (red) need urgent action, 50–69% (yellow) suggest improvement, and above 70% (green) indicate strengths." 
        if LANG == "English" else
        "Puntuaciones por debajo del 50% (rojo) requieren acción urgente, 50–69% (amarillo) sugieren mejoras, y por encima del 70% (verde) indican fortalezas.",
        unsafe_allow_html=True
    )
    styled_df = df.style.applymap(color_percent, subset=["Percent"]).format({"Percent": "{:.1f}%"})
    st.dataframe(styled_df)

    # Interactive bar chart
    fig = px.bar(
        df.reset_index(),
        y="index",
        x="Percent",
        orientation='h',
        title="Workplace Strengths and Opportunities" if LANG == "English" else "Fortalezas y Oportunidades del Lugar de Trabajo",
        labels={"index": "Category", "Percent": "Score (%)"},
        color="Percent",
        color_continuous_scale=["#ff4d4d", "#ffd700", "#28a745"],
        range_x=[0, 100]
    )
    fig.add_vline(x=70, line_dash="dash", line_color="blue", annotation_text="Target (70%)", annotation_position="top")
    for i, row in df.iterrows():
        if row["Percent"] < 70:
            fig.add_annotation(
                x=row["Percent"], y=i,
                text="Priority", showarrow=True, arrowhead=2, ax=20, ay=-30,
                font=dict(color="red", size=12)
            )
    fig.update_layout(
        height=400,
        showlegend=False,
        title_x=0.5,
        xaxis_title="Score (%)",
        yaxis_title="Category",
        coloraxis_showscale=False
    )
    st.plotly_chart(fig, use_container_width=True)

    # Question-level breakdown
    st.markdown(
        "<div class='subheader'>Drill Down: Question-Level Insights</div>" 
        if LANG == "English" else 
        "<div class='subheader'>Análisis Detallado: Perspectivas a Nivel de Pregunta</div>",
        unsafe_allow_html=True
    )
    selected_category = st.selectbox(
        "Select Category to Explore" if LANG == "English" else "Seleccionar Categoría para Explorar",
        categories
    )
    question_scores = pd.DataFrame({
        "Question": questions[selected_category][LANG],
        "Score": [score / 5 * 100 for score in st.session_state.responses[selected_category]]
    })
    fig_questions = px.bar(
        question_scores,
        x="Score",
        y="Question",
        orientation='h',
        title=f"Question Scores for {selected_category}" if LANG == "English" else f"Puntuaciones de Preguntas para {selected_category}",
        labels={"Score": "Score (%)", "Question": "Question"},
        color="Score",
        color_continuous_scale=["#ff4d4d", "#ffd700", "#28a745"],
        range_x=[0, 100]
    )
    fig_questions.update_layout(
        height=300 + len(question_scores) * 50,
        showlegend=False,
        title_x=0.5,
        xaxis_title="Score (%)",
        yaxis_title="Question",
        coloraxis_showscale=False
    )
    st.plotly_chart(fig_questions, use_container_width=True)

    # Actionable insights
    st.markdown(
        "<div class='subheader'>Actionable Insights</div>" if LANG == "English" else "<div class='subheader'>Perspectivas Accionables</div>",
        unsafe_allow_html=True
    )
    insights = []
    recommendations = {
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
        if df.loc[cat, "Percent"] < 50:
            insights.append(
                f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (High Priority). Focus on immediate improvements."
                if LANG == "English" else
                f"**{cat}** obtuvo {df.loc[cat, 'Percent']:.1f}% (Alta Prioridad). Enfócate en mejoras inmediatas."
            )
        elif df.loc[cat, "Percent"] < 70:
            insights.append(
                f"**{cat}** scored {df.loc[cat, 'Percent']:.1f}% (Medium Priority). Consider targeted actions."
                if LANG == "English" else
                f"**{cat}** obtuvo {df.loc[cat, 'Percent']:.1f}% (Prioridad Media). Considera acciones específicas."
            )
    if insights:
        st.markdown(
            "<div class='insights'>" + "<br>".join(insights) + "</div>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            "<div class='insights'>All categories scored above 70%! Continue maintaining these strengths.</div>" 
            if LANG == "English" else 
            "<div class='insights'>¡Todas las categorías obtuvieron más del 70%! Continúa manteniendo estas fortalezas.</div>",
            unsafe_allow_html=True
        )

    # LEAN 2.0 Institute Advertisement
    st.markdown(
        "<div class='subheader'>Optimize Your Workplace with LEAN 2.0 Institute</div>" 
        if LANG == "English" else 
        "<div class='subheader'>Optimiza tu Lugar de Trabajo con LEAN 2.0 Institute</div>",
        unsafe_allow_html=True
    )
    ad_text = []
    if overall_score < 85:
        ad_text.append(
            "Your audit results indicate opportunities to create an optimal workplace. LEAN 2.0 Institute offers tailored consulting services to transform your workplace into an ethical, lean, and human-centered environment."
            if LANG == "English" else
            "Los resultados de tu auditoría indican oportunidades para crear un lugar de trabajo óptimo. LEAN 2.0 Institute ofrece servicios de consultoría personalizados para transformar tu lugar de trabajo en un entorno ético, lean y centrado en las personas."
        )
        if df["Percent"].min() < 70:
            low_categories = df[df["Percent"] < 70].index.tolist()
            services = {
                "Empowering Employees": "Employee Engagement and Empowerment Programs",
                "Ethical Leadership": "Leadership Coaching and Ethical Decision-Making Workshops",
                "Human-Centered Operations": "Lean Process Optimization with Human-Centered Design",
                "Sustainable and Ethical Practices": "Sustainability Strategy and Ethical Supply Chain Consulting",
                "Well-Being and Balance": "Employee Well-Being and Resilience Programs"
            }
            ad_text.append(
                f"Key areas for improvement include {', '.join(low_categories)}. LEAN 2.0 Institute specializes in: {', '.join([services[cat] for cat in low_categories])}."
                if LANG == "English" else
                f"Las áreas clave para mejorar incluyen {', '.join(low_categories)}. LEAN 2.0 Institute se especializa en: {', '.join([services[cat] for cat in low_categories])}."
            )
    else:
        ad_text.append(
            "Congratulations on your excellent workplace! Partner with LEAN 2.0 Institute to sustain and enhance your strengths through advanced lean strategies and leadership development."
            if LANG == "English" else
            "¡Felicidades por tu excelente lugar de trabajo! Asóciate con LEAN 2.0 Institute para mantener y mejorar tus fortalezas mediante estrategias lean avanzadas y desarrollo de liderazgo."
        )
    ad_text.append(
        "Contact us at www.lean2institute.com or info@lean2institute.com for a consultation to elevate your workplace to the next level!"
        if LANG == "English" else
        "¡Contáctanos en www.lean2institute.com o info@lean2institute.com para una consulta y lleva tu lugar de trabajo al siguiente nivel!"
    )
    st.markdown("<div class='insights'>" + "<br>".join(ad_text) + "</div>", unsafe_allow_html=True)

    # PDF Report with Visualizations and Action Plan
    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(0, 123, 255)
            title = "Ethical Workplace Audit Report" if st.session_state.language == "English" else "Informe de Auditoría del Lugar de Trabajo Ético"
            self.cell(0, 10, title, 0, 1, "C")
            self.ln(5)
        
        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(100)
            self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

    try:
        pdf = PDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(51)
        pdf.cell(0, 10, f"Overall Workplace Grade: {grade} ({overall_score:.1f}%)", ln=True)
        pdf.set_font("Helvetica", size=12)
        pdf.multi_cell(0, 10, grade_description)
        pdf.ln(5)
        
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Audit Results" if st.session_state.language == "English" else "Resultados de la Auditoría", ln=True)
        pdf.set_font("Helvetica", size=12)
        pdf.ln(5)
        for cat, row in df.iterrows():
            pdf.cell(0, 10, f"{cat}: {row['Percent']:.1f}% (Priority: {row['Priority']})", ln=True)
        
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Action Plan" if st.session_state.language == "English" else "Plan de Acción", ln=True)
        pdf.set_font("Helvetica", size=12)
        pdf.ln(5)
        for cat in categories:
            if df.loc[cat, "Percent"] < 70:
                pdf.set_font("Helvetica", "B", 12)
                pdf.cell(0, 10, cat, ln=True)
                pdf.set_font("Helvetica", size=12)
                for idx, score in enumerate(st.session_state.responses[cat]):
                    if score / 5 * 100 < 70:
                        question = questions[cat][st.session_state.language][idx]
                        rec = recommendations[cat][idx]
                        pdf.multi_cell(0, 10, f"- {question}: {rec}")
                pdf.ln(5)
        
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Partner with LEAN 2.0 Institute" if st.session_state.language == "English" else "Asóciate con LEAN 2.0 Institute", ln=True)
        pdf.set_font("Helvetica", size=12)
        pdf.ln(5)
        for text in ad_text:
            pdf.multi_cell(0, 10, text)
        
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(40, 167, 69)
        pdf.cell(0, 10, "Certificate of Completion" if st.session_state.language == "English" else "Certificado de Finalización", ln=True, align="C")
        pdf.ln(10)
        pdf.set_font("Helvetica", size=12)
        pdf.set_text_color(51)
        pdf.multi_cell(
            0, 10, 
            "Congratulations on completing the Ethical Workplace Audit! Your insights are helping to create a more ethical, human-centered, and sustainable workplace." 
            if st.session_state.language == "English" else 
            "¡Felicidades por completar la Auditoría del Lugar de Trabajo Ético! Tus aportes están ayudando a crear un lugar de trabajo más ético, centrado en las personas y sostenible."
        )
        
        pdf_output = io.BytesIO()
        pdf_content = pdf.output(dest='S')  # Get PDF as bytearray
        pdf_output.write(pdf_content)  # Write bytearray directly to BytesIO
        pdf_output.seek(0)
        b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
        href_pdf = (
            f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_workplace_audit_report.pdf" class="download-link">Download PDF Report & Action Plan</a>' 
            if st.session_state.language == "English" else 
            f'<a href="data:application/pdf;base64,{b64_pdf}" download="informe_auditoria_lugar_trabajo_etico.pdf" class="download-link">Descargar Informe PDF y Plan de Acción</a>'
        )
        st.markdown(href_pdf, unsafe_allow_html=True)
        pdf_output.close()
    except Exception as e:
        st.error(f"Failed to generate PDF: {str(e)}")

    # Excel export
    try:
        excel_output = io.BytesIO()
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            df.to_excel(
                writer, 
                sheet_name='Audit Results' if st.session_state.language == "English" else 'Resultados de la Auditoría', 
                float_format="%.1f"
            )
            pd.DataFrame({"Overall Score": [overall_score], "Grade": [grade]}).to_excel(
                writer, 
                sheet_name='Summary' if st.session_state.language == "English" else 'Resumen', 
                index=False
            )
        excel_output.seek(0)
        b64_excel = base64.b64encode(excel_output.getvalue()).decode()
        href_excel = (
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_workplace_audit_results.xlsx" class="download-link">Download Excel Report</a>' 
            if LANG == "English" else 
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="resultados_auditoria_lugar_trabajo_etico.xlsx" class="download-link">Descargar Informe Excel</a>'
        )
        st.markdown(href_excel, unsafe_allow_html=True)
        excel_output.close()
    except ImportError:
        st.error("Excel export requires 'xlsxwriter'. Please install it using `pip install xlsxwriter`.")
    except Exception as e:
        st.error(f"Failed to generate Excel file: {str(e)}")

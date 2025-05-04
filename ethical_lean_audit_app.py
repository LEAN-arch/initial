import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
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
    "English": ["Not at All", "Rarely", "Somewhat", "Mostly", "Fully"],
    "Español": ["Nada", "Raramente", "Parcialmente", "Mayormente", "Completamente"]
}

# Curated audit categories and questions
questions = {
    "Empowering Employees": {
        "English": [
            "How effectively are employees empowered to contribute to ethical workplace practices?",
            "To what extent do employees feel their ideas for improvement are valued and acted upon?",
            "How well does the workplace foster a culture of trust and open communication?",
            "Are employees provided with meaningful opportunities for growth and development?"
        ],
        "Español": [
            "¿Con qué eficacia se empodera a los empleados para contribuir a prácticas laborales éticas?",
            "¿Hasta qué punto sienten los empleados que sus ideas de mejora son valoradas y puestas en práctica?",
            "¿Qué tan bien fomenta el lugar de trabajo una cultura de confianza y comunicación abierta?",
            "¿Se les proporciona a los empleados oportunidades significativas para el crecimiento y desarrollo?"
        ]
    },
    "Ethical Leadership": {
        "English": [
            "How consistently do leaders demonstrate ethical and transparent decision-making?",
            "To what degree are ethical values integrated into workplace policies and practices?",
            "How effectively do leaders encourage accountability for ethical behavior?"
        ],
        "Español": [
            "¿Con qué consistencia demuestran los líderes una toma de decisiones ética y transparente?",
            "¿En qué medida se integran los valores éticos en las políticas y prácticas del lugar de trabajo?",
            "¿Qué tan efectivamente fomentan los líderes la responsabilidad por el comportamiento ético?"
        ]
    },
    "Human-Centered Operations": {
        "English": [
            "How well do workplace processes prioritize employee well-being alongside efficiency?",
            "To what extent are lean practices designed to enhance human dignity and respect?",
            "How effectively does the workplace adapt to feedback to improve human-centered operations?"
        ],
        "Español": [
            "¿Qué tan bien priorizan los procesos del lugar de trabajo el bienestar de los empleados junto con la eficiencia?",
            "¿En qué medida están diseñadas las prácticas lean para mejorar la dignidad y el respeto humano?",
            "¿Qué tan efectivamente se adapta el lugar de trabajo a la retroalimentación para mejorar las operaciones centradas en las personas?"
        ]
    },
    "Sustainable and Ethical Practices": {
        "English": [
            "How strongly does the workplace integrate sustainability into its lean strategies?",
            "To what extent are suppliers and partners chosen based on ethical and sustainable practices?",
            "How effectively does the workplace reduce its environmental impact through lean practices?"
        ],
        "Español": [
            "¿Con qué fuerza integra el lugar de trabajo la sostenibilidad en sus estrategias lean?",
            "¿En qué medida se eligen proveedores y socios basados en prácticas éticas y sostenibles?",
            "¿Qué tan efectivamente reduce el lugar de trabajo su impacto ambiental a través de prácticas lean?"
        ]
    },
    "Well-Being and Balance": {
        "English": [
            "How well does the workplace support employee well-being and work-life balance?",
            "To what extent are stress and burnout proactively monitored and addressed?",
            "How effectively does the workplace foster a culture of empathy and support?"
        ],
        "Español": [
            "¿Qué tan bien apoya el lugar de trabajo el bienestar de los empleados y el equilibrio entre trabajo y vida?",
            "¿En qué medida se monitorean y abordan proactivamente el estrés y el agotamiento?",
            "¿Qué tan efectivamente fomenta el lugar de trabajo una cultura de empatía y apoyo?"
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
# Motivational messages at milestones
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
    st.dataframe(df.style.format({"Percent": "{:.1f}%"}))

    # Radar chart
    @st.cache_data
    def generate_radar_chart(categories, values):
        fig, ax = plt.subplots(figsize=(10, 7), subplot_kw=dict(polar=True))
        angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
        values += values[:1]
        angles += angles[:1]
        ax.plot(angles, values, linewidth=2, linestyle='solid', label='Score', color='#007bff')
        ax.fill(angles, values, '#e9f7ff', alpha=0.5)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(categories, fontsize=10, wrap=True, color='#333', ha='center')
        for label, angle in zip(ax.get_xticklabels(), angles[:-1]):
            label.set_rotation(angle * 180 / np.pi - 90)
        ax.set_yticklabels([])
        ax.set_title(
            "Ethical Workplace Radar" if st.session_state.language == "English" else 
            "Radar del Lugar de Trabajo Ético", 
            size=16, pad=20, color='#007bff'
        )
        ax.grid(True, color='#ccc', linestyle='--')
        plt.tight_layout()
        return fig

    categories = list(df.index)
    values = df["Percent"].values.tolist()
    st.pyplot(generate_radar_chart(categories, values))

    # PDF Report with Completion Certificate
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
        pdf.set_font("Helvetica", size=12)
        pdf.set_text_color(51)
        pdf.cell(
            0, 10, 
            "Audit Results" if st.session_state.language == "English" else "Resultados de la Auditoría", 
            ln=True
        )
        pdf.ln(5)
        
        for cat, row in df.iterrows():
            pdf.cell(0, 10, f"{cat}: {row['Percent']:.1f}%", ln=True)
        
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
            f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_workplace_audit_report.pdf" class="download-link">Download PDF Report & Certificate</a>' 
            if st.session_state.language == "English" else 
            f'<a href="data:application/pdf;base64,{b64_pdf}" download="informe_auditoria_lugar_trabajo_etico.pdf" class="download-link">Descargar Informe PDF y Certificado</a>'
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
        excel_output.seek(0)
        b64_excel = base64.b64encode(excel_output.getvalue()).decode()
        href_excel = (
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_workplace_audit_results.xlsx" class="download-link">Download Excel Report</a>' 
            if st.session_state.language == "English" else 
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="resultados_auditoria_lugar_trabajo_etico.xlsx" class="download-link">Descargar Informe Excel</a>'
        )
        st.markdown(href_excel, unsafe_allow_html=True)
        excel_output.close()
    except ImportError:
        st.error("Excel export requires 'xlsxwriter'. Please install it using `pip install xlsxwriter`.")
    except Exception as e:
        st.error(f"Failed to generate Excel file: {str(e)}")

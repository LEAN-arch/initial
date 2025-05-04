import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import base64
import io
import numpy as np

# Custom CSS for enhanced visual impact
st.markdown("""
    <style>
        body {
            font-family: 'Helvetica Neue', Arial, sans-serif;
            background-color: #f9f9f9;
        }
        .main {
            background-color: #ffffff;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
        }
        .stButton>button {
            background-color: #005b96;
            color: white;
            border-radius: 5px;
            padding: 10px 20px;
            font-weight: bold;
            transition: background-color 0.3s;
        }
        .stButton>button:hover {
            background-color: #003366;
        }
        .stRadio>label {
            background-color: #e6f0fa;
            padding: 8px;
            border-radius: 5px;
            margin: 5px 0;
        }
        .header {
            color: #005b96;
            font-size: 2.5em;
            text-align: center;
            margin-bottom: 20px;
        }
        .subheader {
            color: #333;
            font-size: 1.5em;
            margin-top: 20px;
        }
        .sidebar .sidebar-content {
            background-color: #f0f4f8;
            border-radius: 10px;
            padding: 20px;
        }
        .download-link {
            color: #005b96;
            font-weight: bold;
            text-decoration: none;
        }
        .download-link:hover {
            color: #003366;
            text-decoration: underline;
        }
    </style>
""", unsafe_allow_html=True)

# Set page configuration
st.set_page_config(page_title="Ethical Lean Audit", layout="wide", initial_sidebar_state="expanded")

# Bilingual support
LANG = st.sidebar.selectbox("Language / Idioma", ["English", "Español"], help="Select your preferred language / Selecciona tu idioma preferido")

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# Audit categories and questions in both languages
questions = {
    "Employee Empowerment and Engagement": {
        "English": [
            "Are employees actively involved in decision-making processes that directly impact their daily work and job satisfaction?",
            "Do employees feel their contributions to continuous improvement are valued and recognized by leadership?",
            "Is there a structured feedback system where employees can openly express concerns and suggest improvements?",
            "Are all employees given equal access to training and career development opportunities?",
            "Is there a culture of trust and respect where employees are encouraged to suggest innovations?"
        ],
        "Español": [
            "¿Están los empleados activamente involucrados en los procesos de toma de decisiones que impactan directamente en su trabajo diario y satisfacción laboral?",
            "¿Sienten los empleados que sus contribuciones a la mejora continua son valoradas y reconocidas por el liderazgo?",
            "¿Existe un sistema estructurado de retroalimentación donde los empleados puedan expresar abiertamente preocupaciones y sugerir mejoras?",
            "¿Tienen todos los empleados igual acceso a oportunidades de capacitación y desarrollo profesional?",
            "¿Hay una cultura de confianza y respeto donde se fomenta que los empleados sugieran innovaciones?"
        ]
    },
    "Operational Integrity and Continuous Improvement": {
        "English": [
            "Are Lean processes regularly assessed for compliance with ethical standards?",
            "Is there a link between Lean waste reduction and environmental/social responsibility?",
            "Are bottlenecks proactively identified and addressed?",
            "Does the Lean system adapt based on internal/external feedback?",
            "Are digital tools integrated to ensure transparency and traceability?"
        ],
        "Español": [
            "¿Se evalúan regularmente los procesos Lean para garantizar el cumplimiento de los estándares éticos?",
            "¿Existe una conexión entre la reducción de desperdicios Lean y la responsabilidad ambiental/social?",
            "¿Se identifican y abordan proactivamente los cuellos de botella?",
            "¿Se adapta el sistema Lean en función de retroalimentación interna/externa?",
            "¿Se integran herramientas digitales para garantizar transparencia y trazabilidad?"
        ]
    },
    "Customer Satisfaction and Value": {
        "English": [
            "Is customer satisfaction a key metric for Lean success?",
            "Is feedback integrated into continuous improvement?",
            "Do Lean processes prioritize value over efficiency?",
            "Is customer satisfaction regularly measured?",
            "Are product/service improvements validated by customers?"
        ],
        "Español": [
            "¿Es la satisfacción del cliente una métrica clave para el éxito de Lean?",
            "¿Se integra la retroalimentación en la mejora continua?",
            "¿Priorizan los procesos Lean el valor sobre la eficiencia?",
            "¿Se mide regularmente la satisfacción del cliente?",
            "¿Se validan las mejoras de productos/servicios con los clientes?"
        ]
    },
    "Ethical Leadership and Decision-Making": {
        "English": [
            "Does leadership model ethical behavior and transparency?",
            "Are ethical considerations part of Lean tool implementation?",
            "Is the workforce impact of Lean transparent to stakeholders?",
            "Are social and ethical Lean impacts tracked and addressed?",
            "Are leaders accountable for ethical Lean implementation?"
        ],
        "Español": [
            "¿Modela el liderazgo un comportamiento ético y transparencia?",
            "¿Forman parte las consideraciones éticas de la implementación de herramientas Lean?",
            "¿Es transparente el impacto de Lean en la fuerza laboral para las partes interesadas?",
            "¿Se rastrean y abordan los impactos sociales y éticos de Lean?",
            "¿Son los líderes responsables de la implementación ética de Lean?"
        ]
    },
    "Sustainability and Long-Term Impact": {
        "English": [
            "Does Lean focus on long-term sustainable practices?",
            "Are Lean principles part of CSR strategies?",
            "Is carbon footprint reduction part of Lean?",
            "Are sustainability metrics part of Lean performance?",
            "Is there a scaling plan that maintains ethics and sustainability?"
        ],
        "Español": [
            "¿Se enfoca Lean en prácticas sostenibles a largo plazo?",
            "¿Forman parte los principios Lean de las estrategias de responsabilidad social corporativa?",
            "¿Es la reducción de la huella de carbono parte de Lean?",
            "¿Forman parte las métricas de sostenibilidad del rendimiento Lean?",
            "¿Existe un plan de escalado que mantenga la ética y la sostenibilidad?"
        ]
    },
    "Supplier and Partner Relationships": {
        "English": [
            "Are Lean principles applied fairly to suppliers?",
            "Does Lean avoid unfair labor or environmental harm in supply chains?",
            "Do suppliers collaborate in Lean process improvements?",
            "Are ethical suppliers prioritized?",
            "Are supplier ethics regularly assessed?"
        ],
        "Español": [
            "¿Se aplican los principios Lean de manera justa a los proveedores?",
            "¿Evita Lean el daño laboral o ambiental injusto en las cadenas de suministro?",
            "¿Colaboran los proveedores en las mejoras de procesos Lean?",
            "¿Se priorizan los proveedores éticos?",
            "¿Se evalúan regularmente la ética de los proveedores?"
        ]
    },
    "Lean System Transparency and Adaptability": {
        "English": [
            "Are Lean processes transparent to stakeholders?",
            "Is the Lean system flexible to change?",
            "Are Lean strategies regularly reviewed?",
            "Can employees propose Lean process changes?",
            "Is continuous learning fostered in Lean culture?"
        ],
        "Español": [
            "¿Son los procesos Lean transparentes para las partes interesadas?",
            "¿Es el sistema Lean flexible al cambio?",
            "¿Se revisan regularmente las estrategias Lean?",
            "¿Pueden los empleados proponer cambios en los procesos Lean?",
            "¿Se fomenta el aprendizaje continuo en la cultura Lean?"
        ]
    },
    "Cultural Alignment and Organizational Health": {
        "English": [
            "Is Lean aligned with organizational mission and culture?",
            "Are Lean tools adapted to company culture?",
            "Is cross-departmental collaboration promoted?",
            "Are voices from all levels heard in Lean?",
            "Is psychological safety encouraged?"
        ],
        "Español": [
            "¿Está Lean alineado con la misión y la cultura organizacional?",
            "¿Se adaptan las herramientas Lean a la cultura de la empresa?",
            "¿Se promueve la colaboración interdepartamental?",
            "¿Se escuchan las voces de todos los niveles en Lean?",
            "¿Se fomenta la seguridad psicológica?"
        ]
    },
    "Employee Well-Being and Work-Life Balance": {
        "English": [
            "Is employee well-being prioritized in Lean?",
            "Are stress and burnout monitored?",
            "Are workloads regularly assessed and adjusted?",
            "Is there a culture of empathy and support?",
            "Is Lean performance measured holistically (efficiency + satisfaction)?"
        ],
        "Español": [
            "¿Se prioriza el bienestar de los empleados en Lean?",
            "¿Se monitorean el estrés y el agotamiento?",
            "¿Se evalúan y ajustan regularmente las cargas de trabajo?",
            "¿Hay una cultura de empatía y apoyo?",
            "¿Se mide el rendimiento Lean de manera holística (eficiencia + satisfacción)?"
        ]
    },
    "Ethical Decision-Making and Transparency in Metrics": {
        "English": [
            "Are ethical frameworks used in Lean performance evaluation?",
            "Are KPIs aligned with ethical values?",
            "Are Lean outcomes celebrated fairly?",
            "Are ethical implications reviewed before Lean decisions?",
            "Are Lean metrics shared externally for accountability?"
        ],
        "Español": [
            "¿Se utilizan marcos éticos en la evaluación del rendimiento Lean?",
            "¿Están los KPI alineados con valores éticos?",
            "¿Se celebran los resultados Lean de manera justa?",
            "¿Se revisan las implicaciones éticas antes de las decisiones Lean?",
            "¿Se comparten las métricas Lean externamente para rendir cuentas?"
        ]
    }
}

# Initialize session state for responses and progress
if 'responses' not in st.session_state:
    st.session_state.responses = {cat: [0] * len(questions[cat][LANG]) for cat in questions}
if 'current_category' not in st.session_state:
    st.session_state.current_category = 0

# Main title with logo placeholder
st.markdown('<div class="header">Ethical Lean Audit</div>' if LANG == "English" else '<div class="header">Auditoría Lean Ética</div>', unsafe_allow_html=True)
st.markdown("![Logo](https://via.placeholder.com/100x50?text=Your+Logo)" if LANG == "English" else "![Logo](https://via.placeholder.com/100x50?text=Tu+Logo)", unsafe_allow_html=True)

# Progress bar
categories = list(questions.keys())
progress = (st.session_state.current_category / len(categories)) * 100
st.progress(progress)

# Category navigation
st.sidebar.subheader("Progress" if LANG == "English" else "Progreso")
category_index = st.sidebar.slider(
    "Select Category / Seleccionar Categoría",
    0, len(categories) - 1, st.session_state.current_category,
    disabled=False
)
st.session_state.current_category = category_index
category = categories[category_index]

# Collect responses for the current category
st.markdown(f'<div class="subheader">{category}</div>', unsafe_allow_html=True)
for idx, q in enumerate(questions[category][LANG]):
    score = st.radio(
        f"{q}",
        list(range(1, 6)),
        format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
        key=f"{category}_{idx}",
        horizontal=True
    )
    st.session_state.responses[category][idx] = score

# Navigation buttons
col1, col2 = st.columns(2)
with col1:
    if st.button("Previous Category" if LANG == "English" else "Categoría Anterior", disabled=category_index == 0):
        st.session_state.current_category -= 1
        st.rerun()
with col2:
    if st.button("Next Category" if LANG == "English" else "Siguiente Categoría", disabled=category_index == len(categories) - 1):
        st.session_state.current_category += 1
        st.rerun()

# Generate report
if st.button("Generate Report" if LANG == "English" else "Generar Informe", key="generate_report"):
    st.markdown(f'<div class="subheader">{"Audit Results" if LANG == "English" else "Resultados de la Auditoría"}</div>', unsafe_allow_html=True)
    
    # Calculate scores
    results = {cat: sum(scores) for cat, scores in st.session_state.responses.items()}
    df = pd.DataFrame.from_dict(results, orient="index", columns=["Score"])
    df["Percent"] = (df["Score"] / (len(questions[category][LANG]) * 5)) * 100
    st.dataframe(df.style.format({"Percent": "{:.1f}%"}))

    # Radar chart
    categories = list(df.index)
    values = df["Percent"].values.tolist()
    
    fig, ax = plt.subplots(figsize=(12, 8), subplot_kw=dict(polar=True))
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
    values += values[:1]
    angles += angles[:1]
    
    ax.plot(angles, values, linewidth=2, linestyle='solid', label='Score', color='#005b96')
    ax.fill(angles, values, '#e6f0fa', alpha=0.5)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=10, wrap=True, color='#333')
    ax.set_yticklabels([])
    ax.set_title("Ethical Lean Audit Radar" if LANG == "English" else "Radar de Auditoría Lean Ética", size=18, pad=20, color='#005b96')
    ax.grid(True, color='#ccc', linestyle='--')
    plt.tight_layout()
    st.pyplot(fig)

    # PDF Report
    class PDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 14)
            self.set_text_color(0, 91, 150)
            self.cell(0, 10, "Ethical Lean Audit Report" if LANG == "English" else "Informe de Auditoría Lean Ética", 0, 1, "C")
            self.ln(5)
        
        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(100)
            self.cell(0, 10, f"Page {self.page_no()}", 0, 0, "C")

    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=12)
    pdf.set_text_color(51)
    pdf.cell(0, 10, "Audit Results" if LANG == "English" else "Resultados de la Auditoría", ln=True)
    pdf.ln(5)
    
    for cat, row in df.iterrows():
        pdf.cell(0, 10, f"{cat}: {row['Percent']:.1f}%", ln=True)
    
    pdf_output = io.BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)
    b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
    href_pdf = f'<a href="data:application/pdf;base64,{b64_pdf}" download="ethical_lean_audit_report.pdf" class="download-link">Download PDF Report</a>' if LANG == "English" else f'<a href="data:application/pdf;base64,{b64_pdf}" download="informe_auditoria_lean_etica.pdf" class="download-link">Descargar Informe PDF</a>'
    st.markdown(href_pdf, unsafe_allow_html=True)

    # Excel export
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Audit Results' if LANG == "English" else 'Resultados de la Auditoría', float_format="%.1f")
    excel_output.seek(0)
    b64_excel = base64.b64encode(excel_output.getvalue()).decode()
    href_excel = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_lean_audit_results.xlsx" class="download-link">Download Excel Report</a>' if LANG == "English" else f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="resultados_auditoria_lean_etica.xlsx" class="download-link">Descargar Informe Excel</a>'
    st.markdown(href_excel, unsafe_allow_html=True)


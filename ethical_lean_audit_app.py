
import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
import base64

# Set up multilingual question sets
questions_en = {
    "Employee Empowerment and Engagement": [
        "Are employees actively involved in decision-making processes that directly impact their daily work and job satisfaction?",
        "Do employees feel their contributions to continuous improvement are valued and recognized by leadership?",
        "Is there a structured feedback system in place where employees can openly express concerns and suggest improvements?",
        "Are all employees given equal access to training and career development opportunities?",
        "Is there a culture of trust and respect where employees are encouraged to challenge the status quo?"
    ],
    "Operational Integrity and Continuous Improvement": [
        "Are Lean processes regularly assessed for compliance with ethical standards and regulations?",
        "Is there a clear link between Lean’s waste reduction goals and social/environmental responsibility?",
        "Are bottlenecks proactively identified and addressed with clear actions?",
        "Is there a mechanism for continuous feedback from customers?",
        "Are digital tools integrated into Lean for transparency and traceability?"
    ],
    "Customer Satisfaction and Value": [
        "Is customer satisfaction prioritized as a Lean success metric?",
        "Is customer feedback integrated into Lean improvements?",
        "Do Lean practices deliver value without compromising quality?",
        "Is customer satisfaction measured and monitored?",
        "Are product/service improvements validated by customers?"
    ]
}

questions_es = {
    "Empoderamiento y Compromiso de los Empleados": [
        "¿Participan activamente los empleados en los procesos de toma de decisiones que afectan su trabajo diario?",
        "¿Sienten los empleados que sus contribuciones a la mejora continua son valoradas por los líderes?",
        "¿Existe un sistema estructurado de retroalimentación donde los empleados puedan expresar inquietudes y sugerencias?",
        "¿Tienen todos los empleados acceso equitativo a la capacitación y al desarrollo profesional?",
        "¿Hay una cultura de confianza y respeto que promueva la innovación?"
    ],
    "Integridad Operacional y Mejora Continua": [
        "¿Se evalúan regularmente los procesos Lean para garantizar el cumplimiento ético y normativo?",
        "¿Existe un vínculo claro entre la reducción de desperdicios y la responsabilidad social y ambiental?",
        "¿Se identifican y abordan proactivamente los cuellos de botella?",
        "¿Existe un mecanismo de retroalimentación continua de clientes?",
        "¿Se integran herramientas digitales para transparencia y trazabilidad?"
    ],
    "Satisfacción del Cliente y Valor": [
        "¿Se prioriza la satisfacción del cliente como métrica clave del éxito Lean?",
        "¿Se incorpora la retroalimentación del cliente en las mejoras Lean?",
        "¿Aseguran las prácticas Lean el valor sin comprometer la calidad?",
        "¿Se mide y supervisa la satisfacción del cliente?",
        "¿Se validan con clientes las mejoras en productos o servicios?"
    ]
}

# Language selection
language = st.radio("Select Language / Selecciona Idioma:", ("English", "Español"))
questions = questions_en if language == "English" else questions_es

# Title
st.title("Ethical Lean Audit")

# Initialize answers list
responses = []

# Display questions and collect responses
st.header("Audit Questions")
for category, qs in questions.items():
    st.subheader(category)
    for q in qs:
        score = st.slider(q, 1, 5, 3)
        responses.append({"Category": category, "Question": q, "Score": score})

# Dataframe creation
df = pd.DataFrame(responses)

# Visualization
st.header("Audit Results Visualization")
summary = df.groupby("Category").mean().reset_index()

fig, ax = plt.subplots()
sns.barplot(data=summary, x="Score", y="Category", palette="coolwarm", ax=ax)
ax.set_title("Average Score by Category")
st.pyplot(fig)

# Downloadable report generator
st.header("Download Your Report")

def create_pdf(data):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Ethical Lean Audit Report", ln=True, align="C")
    for category in data["Category"].unique():
        pdf.ln(10)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(200, 10, txt=f"Category: {category}", ln=True)
        for _, row in data[data["Category"] == category].iterrows():
            pdf.set_font("Arial", "", 11)
            pdf.multi_cell(0, 10, f"{row['Question']} - Score: {row['Score']}")
    # Return PDF as bytes
    pdf_output = BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)
    return pdf_output

pdf_data = create_pdf(df)
b64_pdf = base64.b64encode(pdf_data.read()).decode('utf-8')
st.markdown(f'<a href="data:application/octet-stream;base64,{b64_pdf}" download="ethical_lean_audit_report.pdf">Download PDF Report</a>', unsafe_allow_html=True)

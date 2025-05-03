
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import base64
import io
import numpy as np

st.set_page_config(page_title="Ethical Lean Audit", layout="wide")

# Bilingual support
LANG = st.sidebar.selectbox("Language / Idioma", ["English", "Español"])

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely Practiced", "Partially Implemented", "Mostly Practiced", "Fully Integrated"],
    "Español": ["No practicado", "Raramente practicado", "Parcialmente implementado", "Mayormente practicado", "Totalmente integrado"]
}

# All 10 audit categories with 5 questions each
questions = {
    "Employee Empowerment and Engagement": [
        "Are employees actively involved in decision-making processes that directly impact their daily work and job satisfaction?",
        "Do employees feel their contributions to continuous improvement are valued and recognized by leadership?",
        "Is there a structured feedback system where employees can openly express concerns and suggest improvements?",
        "Are all employees given equal access to training and career development opportunities?",
        "Is there a culture of trust and respect where employees are encouraged to suggest innovations?"
    ],
    "Operational Integrity and Continuous Improvement": [
        "Are Lean processes regularly assessed for compliance with ethical standards?",
        "Is there a link between Lean waste reduction and environmental/social responsibility?",
        "Are bottlenecks proactively identified and addressed?",
        "Does the Lean system adapt based on internal/external feedback?",
        "Are digital tools integrated to ensure transparency and traceability?"
    ],
    "Customer Satisfaction and Value": [
        "Is customer satisfaction a key metric for Lean success?",
        "Is feedback integrated into continuous improvement?",
        "Do Lean processes prioritize value over efficiency?",
        "Is customer satisfaction regularly measured?",
        "Are product/service improvements validated by customers?"
    ],
    "Ethical Leadership and Decision-Making": [
        "Does leadership model ethical behavior and transparency?",
        "Are ethical considerations part of Lean tool implementation?",
        "Is the workforce impact of Lean transparent to stakeholders?",
        "Are social and ethical Lean impacts tracked and addressed?",
        "Are leaders accountable for ethical Lean implementation?"
    ],
    "Sustainability and Long-Term Impact": [
        "Does Lean focus on long-term sustainable practices?",
        "Are Lean principles part of CSR strategies?",
        "Is carbon footprint reduction part of Lean?",
        "Are sustainability metrics part of Lean performance?",
        "Is there a scaling plan that maintains ethics and sustainability?"
    ],
    "Supplier and Partner Relationships": [
        "Are Lean principles applied fairly to suppliers?",
        "Does Lean avoid unfair labor or environmental harm in supply chains?",
        "Do suppliers collaborate in Lean process improvements?",
        "Are ethical suppliers prioritized?",
        "Are supplier ethics regularly assessed?"
    ],
    "Lean System Transparency and Adaptability": [
        "Are Lean processes transparent to stakeholders?",
        "Is the Lean system flexible to change?",
        "Are Lean strategies regularly reviewed?",
        "Can employees propose Lean process changes?",
        "Is continuous learning fostered in Lean culture?"
    ],
    "Cultural Alignment and Organizational Health": [
        "Is Lean aligned with organizational mission and culture?",
        "Are Lean tools adapted to company culture?",
        "Is cross-departmental collaboration promoted?",
        "Are voices from all levels heard in Lean?",
        "Is psychological safety encouraged?"
    ],
    "Employee Well-Being and Work-Life Balance": [
        "Is employee well-being prioritized in Lean?",
        "Are stress and burnout monitored?",
        "Are workloads regularly assessed and adjusted?",
        "Is there a culture of empathy and support?",
        "Is Lean performance measured holistically (efficiency + satisfaction)?"
    ],
    "Ethical Decision-Making and Transparency in Metrics": [
        "Are ethical frameworks used in Lean performance evaluation?",
        "Are KPIs aligned with ethical values?",
        "Are Lean outcomes celebrated fairly?",
        "Are ethical implications reviewed before Lean decisions?",
        "Are Lean metrics shared externally for accountability?"
    ]
}

st.title("Ethical Lean Audit")

responses = {}
for category, qs in questions.items():
    st.header(category)
    responses[category] = []
    for idx, q in enumerate(qs):
        score = st.radio(
            f"{q}",
            list(range(1, 6)),
            format_func=lambda x: f"{x} - {labels[LANG][x-1]}",
            key=f"{category}_{idx}"
        )
        responses[category].append(score)

# Score and visualization
if st.button("Generate Report"):
    st.subheader("Audit Results")
    results = {cat: sum(scores) for cat, scores in responses.items()}
    df = pd.DataFrame.from_dict(results, orient="index", columns=["Score"])
    df["Percent"] = (df["Score"] / (len(list(questions.values())[0]) * 5)) * 100
    st.dataframe(df)

    # Radar chart
    categories = list(df.index)
    values = df["Percent"].values.tolist()
    values += values[:1]
    categories += categories[:1]

    fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(polar=True))
    angles = [n / float(len(categories)) * 2 * np.pi for n in range(len(categories))]
    ax.plot(angles, values, linewidth=2, linestyle='solid')
    ax.fill(angles, values, 'skyblue', alpha=0.4)
    ax.set_yticklabels([])
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=8)
    ax.set_title("Ethical Lean Audit Radar", size=15)
    st.pyplot(fig)

    # PDF Report
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, "Ethical Lean Audit Report", ln=True, align="C")
    pdf.set_font("Arial", size=12)
    for cat, row in df.iterrows():
        pdf.cell(200, 10, f"{cat}: {row['Percent']:.1f}%", ln=True)
    pdf_output = io.BytesIO()
    pdf.output(pdf_output)
    b64_pdf = base64.b64encode(pdf_output.getvalue()).decode()
    href_pdf = f'<a href="data:application/octet-stream;base64,{b64_pdf}" download="ethical_lean_audit_report.pdf">Download PDF Report</a>'
    st.markdown(href_pdf, unsafe_allow_html=True)

    # Excel export
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=True, sheet_name='Audit Results')
    b64_excel = base64.b64encode(excel_output.getvalue()).decode()
    href_excel = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" download="ethical_lean_audit_results.xlsx">Download Excel Report</a>'
    st.markdown(href_excel, unsafe_allow_html=True)


pdf_data = create_pdf(df)
b64_pdf = base64.b64encode(pdf_data.read()).decode('utf-8')
st.markdown(f'<a href="data:application/octet-stream;base64,{b64_pdf}" download="ethical_lean_audit_report.pdf">Download PDF Report</a>', unsafe_allow_html=True)

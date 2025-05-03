import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Ethical Lean Audit", layout="wide")

# CSS for styling the Streamlit app
st.markdown("""
<style>
    .main {background-color: #f9f9f9;}
    .block-container {padding-top: 2rem;}
    h1, h2, h3, h4 {color: #2c3e50;}
    .stRadio > div {flex-direction: row;}
</style>
""", unsafe_allow_html=True)

# Language selection
language = st.selectbox("Select Language", ["English", "Español"])

# Define sections and questions in both languages
audit_sections = {
    "Respect for People": {
        "English": [
            "Are employees engaged in designing solutions that directly affect their work?",
            "Is there a formal program for upskilling and cross-training employees to ensure they are equipped for evolving roles?",
            "Do employees have clear career advancement pathways within the Lean system?",
            "Is there a recognition system for employees who contribute to continuous improvement?"
        ],
        "Español": [
            "¿Están los empleados involucrados en diseñar soluciones que afectan directamente a su trabajo?",
            "¿Existe un programa formal para mejorar y capacitar a los empleados para que estén preparados para roles cambiantes?",
            "¿Los empleados tienen caminos claros de desarrollo profesional dentro del sistema Lean?",
            "¿Existe un sistema de reconocimiento para los empleados que contribuyen a la mejora continua?"
        ]
    },
    "Operational Integrity": {
        "English": [
            "Are Lean processes continuously evaluated for compliance with industry best practices and regulations?",
            "Is waste reduction and process optimization directly linked to customer satisfaction?",
            "Is there a framework for continuously identifying and mitigating bottlenecks in production or service delivery?",
            "Are digital tools and automation incorporated into Lean to enhance data-driven decision-making?"
        ],
        "Español": [
            "¿Los procesos Lean se evalúan continuamente para cumplir con las mejores prácticas y regulaciones de la industria?",
            "¿La reducción de desperdicios y la optimización de procesos están directamente relacionados con la satisfacción del cliente?",
            "¿Existe un marco para identificar y mitigar continuamente los cuellos de botella en la producción o entrega de servicios?",
            "¿Se incorporan herramientas digitales y automatización en Lean para mejorar la toma de decisiones basada en datos?"
        ]
    },
    # Add other sections in similar format
}

results = {}
st.title(f"Ethical Lean Audit Checklist - {language}")

st.write("### Step 1: Answer the Questions")

with st.form("audit_form"):
    for section, questions in audit_sections.items():
        st.subheader(section)
        for i, q in enumerate(questions[language]):
            response = st.radio(q, ["Yes", "No"], key=f"{section}_{i}")
            results[q] = 1 if response == "Yes" else 0
    submitted = st.form_submit_button("Generate Report")

# Once form is submitted, generate results
if submitted:
    st.markdown("---")
    st.header("Step 2: Audit Results & Summary")
    section_scores = {}
    for section, questions in audit_sections.items():
        score = sum([results[q] for q in questions[language]])
        section_scores[section] = score

    df = pd.DataFrame.from_dict(section_scores, orient='index', columns=['Score'])
    df["Max Score"] = 4
    df["Percentage"] = (df["Score"] / df["Max Score"] * 100).round(1)
    st.dataframe(df.style.format({"Percentage": "{:.1f}%"}))

    total_score = df["Score"].sum()
    st.metric("Total Ethical Lean Score", f"{total_score} / 20")

    # Generate analytics and charts
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.barplot(x=df.index, y="Percentage", data=df, palette="viridis", ax=ax)
    ax.set_title("Audit Section Performance")
    ax.set_ylabel("Percentage")
    st.pyplot(fig)

    # Generate a pie chart of the overall performance
    fig2, ax2 = plt.subplots(figsize=(6, 6))
    ax2.pie([total_score, 20 - total_score], labels=["Score", "Remaining"], autopct='%1.1f%%', startangle=90, colors=["#4CAF50", "#FFC107"])
    ax2.axis('equal')
    ax2.set_title("Overall Score Distribution")
    st.pyplot(fig2)

    # Excel export
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            dataframe.to_excel(writer, sheet_name='Audit Summary')
            
            # Adding charts to Excel
            workbook = writer.book
            worksheet = workbook.get_worksheet_by_name('Audit Summary')
            chart = workbook.add_chart({'type': 'column'})
            chart.add_series({'values': f"=Audit Summary!$C$2:$C${len(df) + 1}"})
            worksheet.insert_chart('F2', chart)
            # Pie chart as image
            img_path = "/tmp/overall_score_pie_chart.png"
            fig2.savefig(img_path)
            worksheet.insert_image('G2', img_path)
            
        output.seek(0)
        return output

    excel_file = to_excel(df)
    st.download_button("Download Excel Report", excel_file, "Ethical_Lean_Audit_Report_with_Analytics.xlsx")

    # Actionable recommendations based on scores
    st.markdown("### Actionable Insights")

    if total_score <= 10:
        st.write("**Recommendation:** You are on the right track, but there are areas for improvement. Focus on increasing employee engagement and aligning your Lean processes with customer satisfaction. Consider improving workforce autonomy and adopting more transparent metrics.")
    elif total_score <= 15:
        st.write("**Recommendation:** You're doing well, but there are still opportunities to strengthen ethical Lean practices. Continue to refine your approach to cross-functional collaboration and streamline processes for improved customer satisfaction.")
    else:
        st.write("**Recommendation:** Excellent alignment with Lean 2.0 principles. To scale further, focus on expanding your innovation and feedback loops, and ensure that incentives align with long-term organizational goals.")

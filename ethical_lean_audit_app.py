import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
import uuid
import os

# Set Streamlit page configuration
st.set_page_config(page_title="Ethical Lean Audit", layout="wide")

# CSS for styling the Streamlit app
st.markdown("""
<style>
    .main {background-color: #f9f9f9;}
    .block-container {padding-top: 2rem;}
    h1, h2, h3, h4 {color: #2c3e50;}
    .stRadio > div {flex-direction: row;}
    .stButton > button {background-color: #4CAF50; color: white;}
    .stDownloadButton > button {background-color: #2196F3; color: white;}
</style>
""", unsafe_allow_html=True)

# Language selection with session state to persist choice
if 'language' not in st.session_state:
    st.session_state.language = "English"

language = st.selectbox("Select Language", ["English", "Español"], index=["English", "Español"].index(st.session_state.language))
st.session_state.language = language

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
    }
}

# Initialize results in session state
if 'results' not in st.session_state:
    st.session_state.results = {}

st.title(f"Ethical Lean Audit Checklist - {language}")

st.write("### Step 1: Answer the Questions")

# Form for audit questions
with st.form("audit_form"):
    for section, questions in audit_sections.items():
        st.subheader(section)
        for i, q in enumerate(questions[language]):
            # Use unique key for each radio button
            key = f"{section}_{i}_{str(uuid.uuid4())[:8]}"
            response = st.radio(q, ["Yes", "No"], key=key)
            st.session_state.results[q] = 1 if response == "Yes" else 0
    submitted = st.form_submit_button("Generate Report")

# Process results after form submission
if submitted:
    st.markdown("---")
    st.header("Step 2: Audit Results & Summary")
    
    # Calculate section scores
    section_scores = {}
    for section, questions in audit_sections.items():
        score = sum([st.session_state.results.get(q, 0) for q in questions[language]])
        section_scores[section] = score

    # Create DataFrame for results
    df = pd.DataFrame.from_dict(section_scores, orient='index', columns=['Score'])
    df["Max Score"] = len(audit_sections[list(audit_sections.keys())[0]][language])
    df["Percentage"] = (df["Score"] / df["Max Score"] * 100).round(1)
    
    # Display results table
    st.dataframe(df.style.format({"Percentage": "{:.1f}%"}), use_container_width=True)

    # Calculate and display total score
    total_score = df["Score"].sum()
    max_total_score = df["Max Score"].sum()
    st.metric("Total Ethical Lean Score", f"{total_score} / {max_total_score}")

    # Generate and display bar chart
    plt.style.use('seaborn')
    fig, ax = plt.subplots(figsize=(10, 6))
    sns.barplot(x=df.index, y="Percentage", data=df, palette="viridis")
    ax.set_title("Audit Section Performance", pad=20)
    ax.set_ylabel("Percentage")
    ax.set_xlabel("Sections")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

    # Generate and display pie chart
    fig2, ax2 = plt.subplots(figsize=(6, 6))
    ax2.pie([total_score, max_total_score - total_score], 
            labels=["Score", "Remaining"], 
            autopct='%1.1f%%', 
            startangle=90, 
            colors=["#4CAF50", "#FFC107"])
    ax2.axis('equal')
    ax2.set_title("Overall Score Distribution")
    plt.tight_layout()
    st.pyplot(fig2)
    
    # Excel export function
    def to_excel(dataframe):
        output = BytesIO()
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                dataframe.to_excel(writer, sheet_name='Audit Summary', index=True)
                
                # Adding bar chart to Excel
                workbook = writer.book
                worksheet = writer.sheets['Audit Summary']
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Percentage',
                    'categories': f"=Audit Summary!$A$2:$A${len(dataframe) + 1}",
                    'values': f"=Audit Summary!$C$2:$C${len(dataframe) + 1}"
                })
                chart.set_title({'name': 'Section Performance'})
                chart.set_x_axis({'name': 'Sections'})
                chart.set_y_axis({'name': 'Percentage (%)'})
                worksheet.insert_chart('F2', chart)
                
                # Save pie chart as image
                temp_img_path = f"temp_pie_chart_{str(uuid.uuid4())[:8]}.png"
                fig2.savefig(temp_img_path, bbox_inches='tight')
                worksheet.insert_image('F18', temp_img_path)
                
            output.seek(0)
        finally:
            # Clean up temporary file
            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
        
        return output

    # Generate and provide Excel download
    excel_file = to_excel(df)
    st.download_button(
        label="Download Excel Report",
        data=excel_file,
        file_name="Ethical_Lean_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Actionable recommendations
    st.markdown("### Actionable Insights")
    
    percentage_score = (total_score / max_total_score * 100) if max_total_score > 0 else 0
    if percentage_score <= 50:
        st.write("""
        **Recommendation:** Your organization is in the early stages of implementing ethical Lean practices. 
        Prioritize employee engagement through participatory decision-making and establish formal upskilling programs. 
        Align Lean processes with customer satisfaction metrics and conduct regular compliance audits.
        """)
    elif percentage_score <= 75:
        st.write("""
        **Recommendation:** You're making good progress with ethical Lean practices. 
        Focus on strengthening cross-functional collaboration and integrating digital tools for data-driven decisions. 
        Enhance bottleneck identification processes and ensure consistent employee recognition programs.
        """)
    else:
        st.write("""
        **Recommendation:** Your organization demonstrates strong alignment with ethical Lean principles. 
        To maintain excellence, scale innovation through advanced feedback loops and predictive analytics. 
        Ensure incentives are aligned with long-term goals and expand employee-driven continuous improvement initiatives.
        """)

if __name__ == "__main__":
    st.write("Run this app with: streamlit run ethical_lean_audit_app.py")

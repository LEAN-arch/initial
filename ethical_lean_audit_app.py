import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
import seaborn as sns
import uuid
import os
from datetime import datetime

# Set Streamlit page configuration
st.set_page_config(
    page_title="Ethical Lean Audit", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS for styling the Streamlit app
st.markdown("""
<style>
    .main {background-color: #f9f9f9;}
    .block-container {padding-top: 2rem; padding-bottom: 2rem;}
    h1, h2, h3, h4 {color: #2c3e50;}
    .stRadio > div {flex-direction: row; gap: 1rem;}
    .stButton > button {background-color: #4CAF50; color: white; border-radius: 4px;}
    .stDownloadButton > button {background-color: #2196F3; color: white; border-radius: 4px;}
    .stAlert {border-radius: 4px;}
    .section-box {border: 1px solid #e0e0e0; border-radius: 8px; padding: 1.5rem; margin-bottom: 1.5rem;}
    @media (max-width: 768px) {
        .stRadio > div {flex-direction: column; gap: 0.5rem;}
    }
</style>
""", unsafe_allow_html=True)

# Language selection with session state to persist choice
if 'language' not in st.session_state:
    st.session_state.language = "English"

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
    st.session_state.audit_completed = False

# Sidebar for navigation and settings
with st.sidebar:
    st.title("Settings")
    language = st.selectbox(
        "Select Language", 
        ["English", "Español"], 
        index=["English", "Español"].index(st.session_state.language),
        key="lang_select"
    )
    st.session_state.language = language
    
    if st.session_state.audit_completed:
        if st.button("Start New Audit"):
            st.session_state.results = {}
            st.session_state.audit_completed = False
            st.rerun()

st.title(f"Ethical Lean Audit Checklist - {language}")

if not st.session_state.audit_completed:
    st.write("### Step 1: Answer the Questions")
    
    # Form for audit questions
    with st.form("audit_form"):
        for section, questions in audit_sections.items():
            with st.container():
                st.subheader(section)
                for i, q in enumerate(questions[language]):
                    # Use consistent key based on section and question index
                    key = f"{section}_{i}"
                    response = st.radio(
                        q, 
                        ["Yes", "No"], 
                        key=key,
                        horizontal=True,
                        index=None
                    )
                    st.session_state.results[key] = response
                    
        submitted = st.form_submit_button("Generate Report", type="primary")
        
        if submitted:
            # Validate all questions are answered
            unanswered = [k for k, v in st.session_state.results.items() if v is None]
            if unanswered:
                st.error("Please answer all questions before submitting.")
            else:
                st.session_state.audit_completed = True
                st.rerun()

# Process results after form submission
if st.session_state.audit_completed:
    st.markdown("---")
    st.header("Audit Results & Summary")
    
    # Calculate section scores
    section_scores = {}
    question_details = []
    
    for section, questions in audit_sections.items():
        score = 0
        for i, q in enumerate(questions[language]):
            key = f"{section}_{i}"
            response = st.session_state.results.get(key, "No")
            points = 1 if response == "Yes" else 0
            score += points
            question_details.append({
                "Section": section,
                "Question": q,
                "Response": response,
                "Points": points
            })
        
        section_scores[section] = score

    # Create DataFrames
    df_questions = pd.DataFrame(question_details)
    df_summary = pd.DataFrame.from_dict(section_scores, orient='index', columns=['Score'])
    df_summary["Max Score"] = len(audit_sections[list(audit_sections.keys())[0]][language])
    df_summary["Percentage"] = (df_summary["Score"] / df_summary["Max Score"] * 100).round(1)
    
    # Display detailed results with expanders
    with st.expander("View Detailed Question Responses"):
        st.dataframe(
            df_questions,
            column_config={
                "Points": st.column_config.ProgressColumn(
                    "Score",
                    help="Points earned for this question",
                    format="%d",
                    min_value=0,
                    max_value=1,
                )
            },
            hide_index=True,
            use_container_width=True
        )

    # Display summary table
    st.subheader("Section Performance Summary")
    st.dataframe(
        df_summary.style.format({"Percentage": "{:.1f}%"}).background_gradient(
            subset=["Percentage"], cmap="YlGn"
        ),
        use_container_width=True
    )

    # Calculate and display total score
    total_score = df_summary["Score"].sum()
    max_total_score = df_summary["Max Score"].sum()
    overall_percentage = (total_score / max_total_score * 100).round(1)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Score", f"{total_score} / {max_total_score}")
    with col2:
        st.metric("Overall Percentage", f"{overall_percentage}%")

    # Visualization section
    st.subheader("Performance Visualization")
    
    # Generate and display bar chart
    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(
        x=df_summary.index,
        y="Percentage",
        data=df_summary,
        palette="viridis",
        ax=ax
    )
    ax.set_title("Section Performance", pad=20)
    ax.set_ylabel("Percentage (%)")
    ax.set_xlabel("")
    plt.xticks(rotation=45, ha='right')
    plt.ylim(0, 100)
    
    # Add value labels
    for p in ax.patches:
        ax.annotate(
            f"{p.get_height():.1f}%",
            (p.get_x() + p.get_width() / 2., p.get_height()),
            ha='center',
            va='center',
            xytext=(0, 5),
            textcoords='offset points'
        )
    
    plt.tight_layout()
    st.pyplot(fig)
    plt.close(fig)

    # Generate and display pie chart
    fig2, ax2 = plt.subplots(figsize=(6, 6))
    ax2.pie(
        [total_score, max_total_score - total_score],
        labels=["Achieved", "Remaining"],
        autopct='%1.1f%%',
        startangle=90,
        colors=["#4CAF50", "#FFC107"],
        wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
    )
    ax2.axis('equal')
    ax2.set_title("Overall Score Distribution")
    plt.tight_layout()
    st.pyplot(fig2)
    plt.close(fig2)

    # Excel export function
    def generate_excel_report():
        output = BytesIO()
        temp_img_path = None
        
        try:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write summary sheet
                df_summary.to_excel(writer, sheet_name='Summary', index=True)
                
                # Write detailed responses sheet
                df_questions.to_excel(writer, sheet_name='Question Details', index=False)
                
                # Add charts to Excel
                workbook = writer.book
                worksheet = writer.sheets['Summary']
                
                # Bar chart
                chart1 = workbook.add_chart({'type': 'column'})
                chart1.add_series({
                    'name': 'Percentage',
                    'categories': '=Summary!$A$2:$A$3',
                    'values': '=Summary!$D$2:$D$3',
                    'data_labels': {'value': True, 'percentage': True}
                })
                chart1.set_title({'name': 'Section Performance'})
                chart1.set_x_axis({'name': 'Sections'})
                chart1.set_y_axis({'name': 'Percentage (%)', 'max': 100})
                worksheet.insert_chart('F2', chart1)
                
                # Save pie chart as image and insert
                temp_img_path = f"temp_pie_chart_{str(uuid.uuid4())[:8]}.png"
                fig2.savefig(temp_img_path, bbox_inches='tight', dpi=300)
                worksheet.insert_image('F20', temp_img_path)
                
                # Formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#4472C4',
                    'font_color': 'white',
                    'border': 1
                })
                
                for sheet in writer.sheets:
                    writer.sheets[sheet].set_column('A:Z', 25)
                    writer.sheets[sheet].freeze_panes(1, 0)
                    for col_num, value in enumerate(df_summary.columns.values):
                        writer.sheets[sheet].write(0, col_num, value, header_format)
                
            output.seek(0)
            return output
        finally:
            # Clean up temporary file
            if temp_img_path and os.path.exists(temp_img_path):
                os.remove(temp_img_path)

    # Generate and provide Excel download
    st.subheader("Export Report")
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Download Excel Report",
            data=generate_excel_report(),
            file_name=f"Ethical_Lean_Audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        if st.button("Generate PDF Report (Coming Soon)", use_container_width=True, disabled=True):
            pass

    # Actionable recommendations
    st.subheader("Actionable Insights")
    
    if overall_percentage <= 50:
        rec = """
        **Priority Areas for Improvement:**  
        • Implement structured employee engagement programs for participatory decision-making  
        • Develop formal upskilling and cross-training programs  
        • Establish clear career pathways within Lean systems  
        • Create a recognition system for continuous improvement contributions  
        • Conduct regular compliance audits of Lean processes  
        """
    elif overall_percentage <= 75:
        rec = """
        **Key Enhancement Opportunities:**  
        • Strengthen cross-functional collaboration mechanisms  
        • Integrate digital tools for better data-driven decisions  
        • Enhance bottleneck identification processes  
        • Ensure consistent employee recognition programs  
        • Align waste reduction metrics with customer satisfaction  
        """
    else:
        rec = """
        **Excellence Maintenance Strategies:**  
        • Scale innovation through advanced feedback loops  
        • Implement predictive analytics for proactive improvements  
        • Align incentives with long-term strategic goals  
        • Expand employee-driven continuous improvement initiatives  
        • Benchmark against industry leaders for best practices  
        """
    
    st.markdown(rec)
    
    st.markdown("---")
    st.info("To start a new audit, use the button in the sidebar.")

if __name__ == "__main__":
    st.write("Audit application ready for use.")

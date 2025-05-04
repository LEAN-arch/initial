import streamlit as st
import pandas as pd
from datetime import datetime
from excel_report_generator import generate_excel_report

# Configuration
CONFIG = {
    "contact": {
        "email": "info@lean2institute.org",
        "website": "www.lean2institute.org"
    }
}
SCORE_THRESHOLDS = {
    "CRITICAL": 50,
    "NEEDS_IMPROVEMENT": 70
}
category_mapping = {
    "English": {
        "Workplace Safety": "safety",
        "Employee Wellbeing": "wellbeing",
        "Ethical Practices": "ethics"
    },
    "Español": {
        "Seguridad Laboral": "safety",
        "Bienestar del Empleado": "wellbeing",
        "Prácticas Éticas": "ethics"
    }
}
questions = {
    "safety": {
        "English": [
            ("Are safety protocols followed?", 100, "Ensure regular safety training."),
            ("Is PPE provided?", 100, "Provide PPE to all employees.")
        ],
        "Español": [
            ("¿Se siguen los protocolos de seguridad?", 100, "Asegure entrenamientos de seguridad regulares."),
            ("¿Se proporciona EPP?", 100, "Proporcione EPP a todos los empleados.")
        ]
    },
    "wellbeing": {
        "English": [
            ("Are mental health resources available?", 100, "Provide mental health support programs.")
        ],
        "Español": [
            ("¿Están disponibles recursos de salud mental?", 100, "Proporcione programas de apoyo a la salud mental.")
        ]
    },
    "ethics": {
        "English": [
            ("Is there a code of conduct?", 100, "Establish a clear code of conduct.")
        ],
        "Español": [
            ("¿Existe un código de conducta?", 100, "Establezca un código de conducta claro.")
        ]
    }
}

st.title("LEAN 2.0 Workplace Audit")
language = st.selectbox("Select Language / Seleccione Idioma", ["English", "Español"])

# Collect responses (simplified example)
st.header("Audit Questions" if language == "English" else "Preguntas de Auditoría")
responses = {}
for cat, q_dict in questions.items():
    responses[cat] = []
    for idx, (q, _, _) in enumerate(q_dict[language]):
        score = st.slider(f"{q}", 0, 100, 50, key=f"{cat}_{idx}")
        responses[cat].append(score)

# Calculate scores
df_data = {
    "Category" if language == "English" else "Categoría": [],
    "Score" if language == "English" else "Puntuación": [],
    "Percentage" if language == "English" else "Porcentaje": [],
    "Priority" if language == "English" else "Prioridad": []
}
for cat in responses:
    display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
    avg_score = sum(responses[cat]) / len(responses[cat])
    priority = (
        "High Priority" if avg_score < SCORE_THRESHOLDS["CRITICAL"] else
        "Medium Priority" if avg_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
        "Low Priority" if language == "English" else
        "Alta Prioridad" if avg_score < SCORE_THRESHOLDS["CRITICAL"] else
        "Prioridad Media" if avg_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
        "Baja Prioridad"
    )
    df_data["Category" if language == "English" else "Categoría"].append(display_cat)
    df_data["Score" if language == "English" else "Puntuación"].append(avg_score)
    df_data["Percentage" if language == "English" else "Porcentaje"].append(avg_score)
    df_data["Priority" if language == "English" else "Prioridad"].append(priority)

df = pd.DataFrame(df_data).set_index("Category" if language == "English" else "Categoría")
df_display = df.copy()
overall_score = df["Percentage" if language == "English" else "Porcentaje"].mean()
grade = "A" if overall_score >= 90 else "B" if overall_score >= 80 else "C" if overall_score >= 70 else "D" if overall_score >= 60 else "F"
REPORT_DATE = datetime.now().strftime("%Y-%m-%d")

# Generate Excel report
if st.button("Generate Report" if language == "English" else "Generar Informe"):
    try:
        excel_buffer = generate_excel_report(
            df=df,
            df_display=df_display,
            questions=questions,
            responses=responses,
            language=language,
            category_mapping=category_mapping,
            SCORE_THRESHOLDS=SCORE_THRESHOLDS,
            CONFIG=CONFIG,
            overall_score=overall_score,
            grade=grade,
            REPORT_DATE=REPORT_DATE
        )
        st.download_button(
            label="Download Report" if language == "English" else "Descargar Informe",
            data=excel_buffer,
            file_name=f"LEAN_Audit_Report_{REPORT_DATE}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Report generated successfully!" if language == "English" else "¡Informe generado con éxito!")
    except Exception as e:
        st.error(f"Failed to generate Excel file: {str(e)}")

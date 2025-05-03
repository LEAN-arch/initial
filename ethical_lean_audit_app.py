import streamlit as st
import pandas as pd
import io
import datetime
from typing import Dict, List, Set

# Set page config as the first Streamlit command
st.set_page_config(page_title="Ethical Lean Audit", layout="wide")

# Questions data
questions = {
    "Employee Experience": {
        "English": [
            "Does work contribute to organizational success and personal growth?",
            "Are leaders evaluated on creating positive work environments?"
        ],
        "Español": [
            "¿Contribuye el trabajo al éxito organizacional y crecimiento personal?",
            "¿Se evalúa a los líderes por crear entornos positivos?"
        ]
    },
    "Process Improvement": {
        "English": [
            "Are human impacts considered when eliminating waste?",
            "Are process changes evaluated for negative consequences?"
        ],
        "Español": [
            "¿Se consideran impactos humanos al eliminar desperdicios?",
            "¿Se evalúan los cambios por consecuencias negativas?"
        ]
    }
}

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely", "Partially", "Mostly", "Fully"],
    "Español": ["No practicado", "Raramente", "Parcialmente", "Mayormente", "Totalmente"]
}

def initialize_session_state():
    """Initialize session state variables."""
    if "lang" not in st.session_state:
        st.session_state.lang = "English"
    if "current_category" not in st.session_state:
        st.session_state.current_category = 0
    if "responses" not in st.session_state:
        st.session_state.responses = {
            cat: [0] * len(questions[cat]["English"]) for cat in questions.keys()
        }
    if "completed_categories" not in st.session_state:
        st.session_state.completed_categories = set()
    if "timestamp" not in st.session_state:
        st.session_state.timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def validate_score(score: int) -> bool:
    """Validate score is between 1 and 5."""
    return 1 <= score <= 5

def generate_report(responses: Dict[str, List[int]], categories: List[str], 
                   timestamp: str, lang: str) -> str:
    """Generate and return audit report as CSV string."""
    results = []
    for cat in categories:
        scores = responses.get(cat, [0, 0])
        total = sum(scores)
        max_score = len(scores) * 5
        percent = (total / max_score * 100) if max_score  max_score > 0 else 0
        results.append({
            "Category": cat,
            "Score": total,
            "Max Score": max_score,
            "Percent": round(percent, 1)
        })
    
    df = pd.DataFrame(results)
    csv_buffer = io.StringIO()
    csv_buffer.write(f"Ethical Lean Audit Report\nDate: {timestamp}\nLanguage: {lang}\n\n")
    df.to_csv(csv_buffer, index=False)
    return csv_buffer.getvalue()

def main():
    """Main Streamlit app logic."""
    initialize_session_state()
    categories = list(questions.keys())
    
    # Language selection
    st.sidebar.header("Settings")
    st.session_state.lang = st.sidebar.radio(
        "Select Language / Seleccionar Idioma",
        ["English", "Español"],
        index=0 if st.session_state.lang == "English" else 1
    )
    
    # Main content
    st.title("Ethical Lean Audit")
    
    # Current category
    category = categories[st.session_state.current_category]
    st.header(category)
    
    # Display questions and collect scores
    scores = st.session_state.responses[category]
    for idx, q in enumerate(questions[category][st.session_state.lang]):
        score = st.radio(
            f"{q}",
            options=[0, 1, 2, 3, 4, 5],
            format_func=lambda x: "Skip" if x == 0 else labels[st.session_state.lang][x-1],
            index=scores[idx],
            key=f"{category}_{idx}"
        )
        scores[idx] = score
    
    # Update completion status
    if all(validate_score(score) for score in scores):
        st.session_state.completed_categories.add(category)
    else:
        st.session_state.completed_categories.discard(category)
    
    # Display score
    score_sum = sum(scores)
    max_score = len(questions[category][st.session_state.lang]) * 5
    score_percent = (score_sum / max_score * 100) if max_score > 0 else 0
    st.write(f"Score: {score_sum}/{max_score} ({score_percent:.1f}%)")
    if category in st.session_state.completed_categories:
        st.success(f"{'Category completed!' if st.session_state.lang == 'English' else '¡Categoría completada!'}")
    st.write(f"Progress: {len(st.session_state.completed_categories)}/{len(categories)} completed "
             f"({(len(st.session_state.completed_categories)/len(categories))*100:.1f}%)")
    
    # Navigation
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("Previous", disabled=st.session_state.current_category == 0):
            st.session_state.current_category -= 1
            st.rerun()
    with col2:
        if st.button("Next", disabled=st.session_state.current_category == len(categories) - 1):
            st.session_state.current_category += 1
            st.rerun()
    with col3:
        if st.button("Reset"):
            st.session_state.responses = {
                cat: [0] * len(questions[cat]["English"]) for cat in categories
            }
            st.session_state.completed_categories.clear()
            st.session_state.current_category = 0
            st.rerun()
    with col4:
        if st.button("Generate Report"):
            incomplete = [cat for cat in categories 
                         if not all(validate_score(score) for score in st.session_state.responses.get(cat, [0, 0]))]
            if incomplete:
                st.error(f"{'Complete all questions for: ' if st.session_state.lang == 'English' else 'Complete todas las preguntas para: '}"
                         f"{', '.join(incomplete)}")
            else:
                report_csv = generate_report(
                    st.session_state.responses,
                    categories,
                    st.session_state.timestamp,
                    st.session_state.lang
                )
                st.download_button(
                    label="Download Report",
                    data=report_csv,
                    file_name="audit_results.csv",
                    mime="text/csv"
                )

if __name__ == "__main__":
    main()

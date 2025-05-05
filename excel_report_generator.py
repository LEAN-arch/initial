import pandas as pd
import io
import xlsxwriter
import logging
from typing import Dict
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def generate_excel_report(
    df: pd.DataFrame,
    df_display: pd.DataFrame,
    questions: Dict,
    responses: Dict,
    language: str,
    category_mapping: Dict,
    SCORE_THRESHOLDS: Dict,
    CONFIG: Dict,
    overall_score: float,
    grade: str,
    REPORT_DATE: str
) -> io.BytesIO:
    """
    Generate an Excel report with a single worksheet containing all content.

    Args:
        df: DataFrame with category scores and priorities
        df_display: DataFrame with display-friendly category names
        questions: Dictionary of questions by category and language
        responses: Dictionary of user responses
        language: Selected language ("Español" or "English")
        category_mapping: Mapping of display categories to internal categories
        SCORE_THRESHOLDS: Thresholds for score categories
        CONFIG: Configuration dictionary with contact info
        overall_score: Overall audit score
        grade: Overall grade
        REPORT_DATE: Report generation date

    Returns:
        io.BytesIO: Excel file buffer with a single worksheet
    """
    logger.debug("Starting Excel report generation with language: %s", language)

    # Validate inputs
    if language not in ["English", "Español"]:
        logger.error("Unsupported language: %s", language)
        raise ValueError(f"Unsupported language: {language}")
    
    if not isinstance(df, pd.DataFrame) or df.empty:
        logger.error("df must be a non-empty pandas DataFrame")
        raise ValueError("df must be a non-empty pandas DataFrame")
    
    percent_col = "Percentage" if language == "English" else "Porcentaje"
    required_columns = ["Score", percent_col, "Priority"] if language == "English" else ["Puntuación", percent_col, "Prioridad"]
    if not all(col in df.columns for col in required_columns):
        logger.error("Required columns missing in df: %s", required_columns)
        raise ValueError(f"Required columns missing in df: {required_columns}")

    # Translation dictionary
    translations = {
        "English": {
            "report_title": "LEAN 2.0 Workplace Audit Report",
            "summary": "Executive Summary",
            "results": "Results & Action Plan",
            "chart": "Performance Overview",
            "contact": "Contact Us",
            "category": "Category",
            "score": "Score",
            "percent": "Percentage",
            "priority": "Priority",
            "high_priority": "High",
            "medium_priority": "Medium",
            "low_priority": "Low",
            "overall_score": "Overall Score",
            "grade": "Grade",
            "action_plan": "Action Plan",
            "effort": "Effort",
            "type": "Type",
            "category_type": "Category",
            "question": "Question",
            "recommendation": "Recommendation",
            "marketing_message": "Partner with LEAN 2.0 Institute to transform your workplace! Contact us today.",
            "date_format": "%m/%d/%Y",
            "metric": "Metric",
            "value": "Value",
            "prepared_by": "Prepared by: LEAN 2.0 Institute"
        },
        "Español": {
            "report_title": "Informe de Auditoría LEAN 2.0",
            "summary": "Resumen Ejecutivo",
            "results": "Resultados y Plan de Acción",
            "chart": "Resumen de Desempeño",
            "contact": "Contáctenos",
            "category": "Categoría",
            "score": "Puntuación",
            "percent": "Porcentaje",
            "priority": "Prioridad",
            "high_priority": "Alta",
            "medium_priority": "Media",
            "low_priority": "Baja",
            "overall_score": "Puntuación General",
            "grade": "Calificación",
            "action_plan": "Plan de Acción",
            "effort": "Esfuerzo",
            "type": "Tipo",
            "category_type": "Categoría",
            "question": "Pregunta",
            "recommendation": "Recomendación",
            "marketing_message": "¡Asóciese con el Instituto LEAN 2.0 para transformar su lugar de trabajo! Contáctenos hoy.",
            "date_format": "%d/%m/%Y",
            "metric": "Métrica",
            "value": "Valor",
            "prepared_by": "Preparado por: Instituto LEAN 2.0"
        }
    }

    # Format date
    try:
        report_date_formatted = datetime.strptime(REPORT_DATE, "%Y-%m-%d").strftime(translations[language]["date_format"])
    except ValueError:
        report_date_formatted = datetime.now().strftime(translations[language]["date_format"])
        logger.warning("Invalid REPORT_DATE format. Using current date: %s", report_date_formatted)

    # Initialize Excel output with a single worksheet
    excel_output = io.BytesIO()
    sheet_name = "Audit Report" if language == "English" else "Informe de Auditoría"
    logger.debug("Creating Excel file with single worksheet: %s", sheet_name)

    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        logger.debug("Single worksheet created: %s", sheet_name)

        # Define formats
        title_format = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'bg_color': '#1E88E5', 'font_color': 'white'})
        subtitle_format = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'left'})
        header_format = workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'bg_color': '#E0E0E0'})
        cell_format = workbook.add_format({'font_size': 10, 'border': 1})
        percent_format = workbook.add_format({'num_format': '0.0%', 'font_size': 10, 'border': 1})
        center_format = workbook.add_format({'align': 'center', 'font_size': 10, 'border': 1})
        wrap_format = workbook.add_format({'text_wrap': True, 'font_size': 10, 'border': 1})

        # Write all sections to the single worksheet
        current_row = 0

        # Title Section
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["report_title"], title_format)
        current_row += 1
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["prepared_by"], subtitle_format)
        current_row += 1
        worksheet.merge_range(current_row, 0, current_row, 6, f"Date: {report_date_formatted}", subtitle_format)
        current_row += 2

        # Executive Summary
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["summary"], title_format)
        current_row += 1
        summary_data = {
            translations[language]["metric"]: [translations[language]["overall_score"], translations[language]["grade"]],
            translations[language]["value"]: [f"{overall_score:.1f}%", grade]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        current_row += len(summary_df) + 2

        # Results & Action Plan
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["results"], title_format)
        current_row += 1
        action_plan_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            score = df.loc[cat, translations[language]["percent"]]
            priority = (
                translations[language]["high_priority"] if score < SCORE_THRESHOLDS["CRITICAL"] else
                translations[language]["medium_priority"] if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
                translations[language]["low_priority"]
            )
            effort = (
                "High" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                "Medium" if score < SCORE_THRESHOLDS["CRITICAL"] else
                "Low" if language == "English" else
                "Alto" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                "Medio" if score < SCORE_THRESHOLDS["CRITICAL"] else
                "Bajo"
            )
            action_plan_data.append({
                translations[language]["category"]: display_cat,
                translations[language]["score"]: score / 100,
                translations[language]["priority"]: priority,
                translations[language]["effort"]: effort,
                translations[language]["action_plan"]: f"Improve {display_cat} with immediate actions." if language == "English" else f"Mejorar {display_cat} con acciones inmediatas."
            })
        action_plan_df = pd.DataFrame(action_plan_data)
        action_plan_df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 15, percent_format)
        worksheet.set_column('C:E', 15, center_format)
        current_row += len(action_plan_df) + 2

        # Performance Overview
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["chart"], title_format)
        current_row += 1
        chart_data = df_display[[translations[language]["percent"]]].reset_index()
        chart_data.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        current_row += len(chart_data) + 2

        # Contact Section
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["contact"], title_format)
        current_row += 1
        contact_df = pd.DataFrame({
            translations[language]["metric"]: ["Email" if language == "English" else "Correo", "Website" if language == "English" else "Sitio Web"],
            translations[language]["value"]: [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        current_row += len(contact_df) + 1
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["marketing_message"], wrap_format)

        # Verify single worksheet
        worksheets = workbook.worksheets()
        if len(worksheets) > 1:
            logger.error("Multiple worksheets detected: %s", [ws.get_name() for ws in worksheets])
            raise ValueError(f"Multiple worksheets created: {[ws.get_name() for ws in worksheets]}")
        logger.debug("Verified single worksheet: %s", sheet_name)

    excel_output.seek(0)
    logger.debug("Excel report generated successfully with single worksheet")
    return excel_output

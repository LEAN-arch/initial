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
    Generate an Excel report with all content consolidated into a single worksheet.

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
            "findings": "Findings",
            "actionable_insights": "Actionable Insights",
            "actionable_charts": "Actionable Charts",
            "contact": "Contact",
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
            "findings": "Hallazgos",
            "actionable_insights": "Perspectivas Accionables",
            "actionable_charts": "Gráficos Accionables",
            "contact": "Contacto",
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

        # Initialize row counter
        current_row = 0

        # Helper function to write section header
        def write_section_header(title):
            nonlocal current_row
            worksheet.merge_range(current_row, 0, current_row, 6, title, title_format)
            current_row += 1

        # Helper function to write DataFrame
        def write_dataframe(df, start_row):
            nonlocal current_row
            df.to_excel(writer, sheet_name=sheet_name, startrow=start_row, index=False)
            current_row = start_row + len(df) + 1

        # Title Section
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["report_title"], title_format)
        current_row += 1
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["prepared_by"], subtitle_format)
        current_row += 1
        worksheet.merge_range(current_row, 0, current_row, 6, f"Date: {report_date_formatted}", subtitle_format)
        current_row += 2

        # Resumen (Summary)
        write_section_header(translations[language]["summary"])
        summary_data = {
            translations[language]["metric"]: [translations[language]["overall_score"], translations[language]["grade"]],
            translations[language]["value"]: [f"{overall_score:.1f}%", grade]
        }
        summary_df = pd.DataFrame(summary_data)
        write_dataframe(summary_df, current_row)

        # Resultados (Results)
        write_section_header(translations[language]["results"])
        results_data = df_display.reset_index()
        results_data.columns = [translations[language]["category"], translations[language]["score"], translations[language]["percent"], translations[language]["priority"]]
        write_dataframe(results_data, current_row)

        # Hallazgos (Findings)
        write_section_header(translations[language]["findings"])
        findings_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            score = df.loc[cat, translations[language]["percent"]]
            if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                findings_data.append({
                    translations[language]["category"]: display_cat,
                    translations[language]["score"]: f"{score:.1f}%",
                    translations[language]["priority"]: translations[language]["high_priority"] if score < SCORE_THRESHOLDS["CRITICAL"] else translations[language]["medium_priority"]
                })
        if findings_data:
            findings_df = pd.DataFrame(findings_data)
            write_dataframe(findings_df, current_row)
        else:
            worksheet.write(current_row, 0, "No critical findings.", cell_format)
            current_row += 1

        # Perspectivas Accionables (Actionable Insights)
        write_section_header(translations[language]["actionable_insights"])
        insights_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            score = df.loc[cat, translations[language]["percent"]]
            if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                insights_data.append({
                    translations[language]["category"]: display_cat,
                    translations[language]["action_plan"]: f"Focus on improving {display_cat}." if language == "English" else f"Concéntrese en mejorar {display_cat}."
                })
        if insights_data:
            insights_df = pd.DataFrame(insights_data)
            write_dataframe(insights_df, current_row)
        else:
            worksheet.write(current_row, 0, "All categories are performing well.", cell_format)
            current_row += 1

        # Gráficos Accionables (Actionable Charts)
        write_section_header(translations[language]["actionable_charts"])
        chart_data = df_display[[translations[language]["percent"]]].reset_index()
        chart_data.columns = [translations[language]["category"], translations[language]["percent"]]
        chart_data.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 15, percent_format)
        current_row += len(chart_data) + 1

        # Contacto (Contact)
        write_section_header(translations[language]["contact"])
        contact_df = pd.DataFrame({
            translations[language]["metric"]: ["Email" if language == "English" else "Correo", "Website" if language == "English" else "Sitio Web"],
            translations[language]["value"]: [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)
        current_row += len(contact_df) + 1
        worksheet.merge_range(current_row, 0, current_row, 6, translations[language]["marketing_message"], wrap_format)
        current_row += 2

        # Verify single worksheet
        worksheets = workbook.worksheets()
        if len(worksheets) > 1:
            logger.error("Multiple worksheets detected: %s", [ws.get_name() for ws in worksheets])
            raise ValueError(f"Multiple worksheets created: {[ws.get_name() for ws in worksheets]}")
        logger.debug("Verified single worksheet: %s", sheet_name)

    excel_output.seek(0)
    logger.debug("Excel report generated successfully with single worksheet")
    return excel_output

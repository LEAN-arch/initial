import pandas as pd
import io
import xlsxwriter
from typing import Dict
import numpy as np
import logging
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
    Generate a single-sheet Excel report with all content on one worksheet,
    ensuring a persuasive narrative to contract LEAN 2.0 Institute services.

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
        io.BytesIO: Excel file buffer
    """
    logger.debug("Starting Excel report generation with language: %s", language)

    # Validate inputs
    if language not in ["English", "Español"]:
        logger.error("Unsupported language: %s", language)
        raise ValueError(f"Unsupported language: {language}. Must be 'English' or 'Español'")
    
    if not isinstance(df, pd.DataFrame) or not isinstance(df_display, pd.DataFrame):
        logger.error("df and df_display must be pandas DataFrames")
        raise ValueError("df and df_display must be pandas DataFrames")
    
    if df.empty or df_display.empty or not responses:
        logger.error("Input data is empty")
        raise ValueError("Input data is empty")
    
    percent_col = "Percentage" if language == "English" else "Porcentaje"
    required_columns = ["Score", percent_col, "Priority"] if language == "English" else ["Puntuación", percent_col, "Prioridad"]
    if not all(col in df.columns for col in required_columns):
        logger.error("Required columns missing in df: %s", required_columns)
        raise ValueError(f"Required columns missing in df: {required_columns}")
    
    if not df.index.is_unique:
        logger.error("DataFrame index contains duplicates")
        raise ValueError("DataFrame index must be unique")
    
    if not all(cat in df.index for cat in questions.keys()):
        logger.error("Not all question categories are present in DataFrame index")
        raise ValueError("Not all question categories are present in DataFrame index")
    
    if not pd.api.types.is_numeric_dtype(df[percent_col]):
        logger.error("Column %s must contain numeric values", percent_col)
        raise ValueError(f"Column {percent_col} must contain numeric values")
    
    if df[percent_col].isna().any():
        logger.error("Column %s contains missing values", percent_col)
        raise ValueError(f"Column {percent_col} contains missing values")
    
    if not (df[percent_col] >= 0).all() or not (df[percent_col] <= 100).all():
        logger.error("Column %s contains values outside 0-100 range", percent_col)
        raise ValueError(f"Column {percent_col} must contain values between 0 and 100")
    
    if len(df) != len(df_display):
        logger.error("Mismatch in row counts between df (%d) and df_display (%d)", len(df), len(df_display))
        raise ValueError(f"Mismatch in row counts between df ({len(df)}) and df_display ({len(df_display)})")
    
    for cat in responses:
        if not responses[cat] or any(score is None or pd.isna(score) for score in responses[cat]):
            logger.error("Invalid or missing responses in category %s", cat)
            raise ValueError(f"Invalid or missing responses in category {cat}")

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
            "critical_categories": "Critical Categories",
            "overall_score": "Overall Score",
            "grade": "Grade",
            "question": "Question",
            "recommendation": "Recommendation",
            "chart_title": "Category Performance",
            "marketing_message": (
                "Partner with LEAN 2.0 Institute to transform your workplace! "
                "Contact us today for ethical, inclusive, and sustainable solutions."
            ),
            "date_format": "%m/%d/%Y",
            "metric": "Metric",
            "value": "Value",
            "action_plan": "Action Plan",
            "effort": "Effort",
            "type": "Type",
            "category_type": "Category",
            "question_type": "Question",
            "prepared_by": "Prepared by: LEAN 2.0 Institute",
            "mission": "Mission: Building ethical, inclusive, sustainable workplaces."
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
            "critical_categories": "Categorías Críticas",
            "overall_score": "Puntuación General",
            "grade": "Calificación",
            "question": "Pregunta",
            "recommendation": "Recomendación",
            "chart_title": "Desempeño por Categoría",
            "marketing_message": (
                "¡Asóciese con el Instituto LEAN 2.0 para transformar su lugar de trabajo! "
                "Contáctenos hoy para soluciones éticas, inclusivas y sostenibles."
            ),
            "date_format": "%d/%m/%Y",
            "metric": "Métrica",
            "value": "Valor",
            "action_plan": "Plan de Acción",
            "effort": "Esfuerzo",
            "type": "Tipo",
            "category_type": "Categoría",
            "question_type": "Pregunta",
            "prepared_by": "Preparado por: Instituto LEAN 2.0",
            "mission": "Misión: Construir lugares de trabajo éticos, inclusivos y sostenibles."
        }
    }

    # Validate translation keys
    required_keys = [
        "report_title", "summary", "results", "chart", "contact", "category", "score", "percent",
        "priority", "high_priority", "medium_priority", "low_priority", "critical_categories",
        "overall_score", "grade", "question", "recommendation", "chart_title", "marketing_message",
        "date_format", "metric", "value", "action_plan", "effort", "type", "category_type",
        "question_type", "prepared_by", "mission"
    ]
    missing_keys = [key for key in required_keys if key not in translations[language]]
    if missing_keys:
        logger.error("Missing translation keys for %s: %s", language, missing_keys)
        raise ValueError(f"Missing translation keys for {language}: {missing_keys}")

    # Format date
    try:
        report_date_formatted = datetime.strptime(REPORT_DATE, "%Y-%m-%d").strftime(translations[language]["date_format"])
        logger.debug("Formatted report date: %s", report_date_formatted)
    except ValueError:
        logger.warning("Invalid date format for REPORT_DATE: %s. Using current date.", REPORT_DATE)
        report_date_formatted = datetime.now().strftime(translations[language]["date_format"])

    # Initialize Excel output
    excel_output = io.BytesIO()
    sheet_name = "Informe de Auditoría" if language == "Español" else "Audit Report"
    logger.debug("Creating single worksheet: %s", sheet_name)

    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        logger.debug("Worksheet %s created", sheet_name)

        # Define formats
        title_format = workbook.add_format({
            'bold': True, 'font_size': 16, 'font_name': 'Arial', 'align': 'center',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        subtitle_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'font_name': 'Arial', 'align': 'left',
            'font_color': '#424242', 'border': 1
        })
        header_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_name': 'Arial', 'align': 'center',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        cell_format = workbook.add_format({
            'font_size': 10, 'font_name': 'Arial', 'border': 1, 'text_wrap': True
        })
        percent_format = workbook.add_format({
            'num_format': '0.0%', 'font_size': 10, 'font_name': 'Arial', 'border': 1
        })
        center_format = workbook.add_format({
            'align': 'center', 'font_size': 10, 'font_name': 'Arial', 'border': 1
        })
        wrap_format = workbook.add_format({
            'text_wrap': True, 'font_size': 10, 'font_name': 'Arial', 'border': 1
        })
        critical_format = workbook.add_format({
            'bg_color': '#D32F2F', 'font_color': 'white', 'border': 1, 'font_name': 'Arial'
        })
        improvement_format = workbook.add_format({
            'bg_color': '#FFD54F', 'font_color': '#212121', 'border': 1, 'font_name': 'Arial'
        })
        good_format = workbook.add_format({
            'bg_color': '#43A047', 'font_color': 'white', 'border': 1, 'font_name': 'Arial'
        })
        alt_row_format = workbook.add_format({
            'bg_color': '#F5F5F5', 'font_size': 10, 'font_name': 'Arial', 'border': 1
        })

        # Cover Section (Rows 1-5)
        logger.debug("Writing Cover Section to %s", sheet_name)
        worksheet.merge_range('A1:G1', translations[language]["report_title"], title_format)
        worksheet.merge_range('A2:G2', translations[language]["prepared_by"], subtitle_format)
        worksheet.merge_range('A3:G3', f"Date: {report_date_formatted}", subtitle_format)
        worksheet.merge_range('A4:G4', translations[language]["mission"], subtitle_format)
        worksheet.merge_range('A5:G5', "Logo Placeholder: Insert LEAN 2.0 Institute logo (PNG, 200x100)", cell_format)
        worksheet.set_column('A:G', 20)

        # Executive Summary Section (Rows 7-12)
        logger.debug("Writing Executive Summary Section to %s", sheet_name)
        worksheet.merge_range('A7:G7', translations[language]["summary"], title_format)
        critical_count = len(df[df[translations[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
        summary_data = {
            translations[language]["metric"]: [
                translations[language]["overall_score"],
                translations[language]["grade"],
                translations[language]["critical_categories"]
            ],
            translations[language]["value"]: [
                f"{overall_score:.1f}%",
                grade,
                critical_count
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name=sheet_name, startrow=8, startcol=0, index=False)
        logger.debug("Wrote Executive Summary to %s at row 8", sheet_name)
        worksheet.write('A9', translations[language]["metric"], header_format)
        worksheet.write('B9', translations[language]["value"], header_format)
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 20)
        for row in range(9, 9 + len(summary_df)):
            worksheet.write(row, 0, summary_df[translations[language]["metric"]][row-9], cell_format)
            worksheet.write(row, 1, summary_df[translations[language]["value"]][row-9], center_format)
            worksheet.write(row, 0, "", alt_row_format if (row-9) % 2 else cell_format)
        worksheet.conditional_format('B10:B10', {
            'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100, 'format': critical_format
        })
        worksheet.conditional_format('B10:B10', {
            'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
            'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-0.01)/100, 'format': improvement_format
        })
        worksheet.conditional_format('B10:B10', {
            'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100, 'format': good_format
        })

        # Results & Action Plan Section (Rows 14-40)
        logger.debug("Writing Results & Action Plan Section to %s", sheet_name)
        worksheet.merge_range('A14:G14', translations[language]["results"], title_format)
        action_plan_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            try:
                score = df.loc[cat, translations[language]["percent"]].item()
                logger.debug("Retrieved score for category %s: %s", cat, score)
            except (ValueError, KeyError) as e:
                logger.error("Failed to retrieve score for category %s: %s", cat, str(e))
                raise ValueError(f"Failed to retrieve score for category {cat}: {str(e)}")
            priority = (
                translations[language]["high_priority"] if score < SCORE_THRESHOLDS["CRITICAL"] else
                translations[language]["medium_priority"] if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
                translations[language]["low_priority"]
            )
            effort = (
                "Alto" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                "Medio" if score < SCORE_THRESHOLDS["CRITICAL"] else
                "Bajo" if language == "Español" else
                "High" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                "Medium" if score < SCORE_THRESHOLDS["CRITICAL"] else
                "Low"
            )
            action_plan_data.append({
                translations[language]["category"]: display_cat,
                translations[language]["score"]: score / 100,
                translations[language]["priority"]: priority,
                translations[language]["effort"]: effort,
                translations[language]["action_plan"]: (
                    f"Mejorar {display_cat} con acciones inmediatas." if language == "Español" else
                    f"Improve {display_cat} with immediate actions."
                ),
                translations[language]["type"]: translations[language]["category_type"]
            })
            for idx, q_score in enumerate(responses[cat]):
                if q_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] and len(action_plan_data) < 20:
                    question, _, rec = questions[cat][language][idx]
                    q_effort = (
                        "Alto" if q_score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                        "Medio" if q_score < SCORE_THRESHOLDS["CRITICAL"] else
                        "Bajo" if language == "Español" else
                        "High" if q_score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                        "Medium" if q_score < SCORE_THRESHOLDS["CRITICAL"] else
                        "Low"
                    )
                    action_plan_data.append({
                        translations[language]["category"]: "",
                        translations[language]["score"]: "",
                        translations[language]["priority"]: "",
                        translations[language]["effort"]: q_effort,
                        translations[language]["action_plan"]: (
                            f"{translations[language]['question']}: {question[:50]}{'...' if len(question) > 50 else ''}\n"
                            f"{translations[language]['recommendation']}: {rec}"
                        ),
                        translations[language]["type"]: translations[language]["question_type"]
                    })
        action_plan_df = pd.DataFrame(action_plan_data).sort_values(translations[language]["score"])
        action_plan_df.to_excel(writer, sheet_name=sheet_name, startrow=15, startcol=0, index=False)
        logger.debug("Wrote Results & Action Plan to %s at row 15", sheet_name)
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 15, percent_format)
        worksheet.set_column('C:C', 15, center_format)
        worksheet.set_column('D:D', 15, center_format)
        worksheet.set_column('E:E', 50, wrap_format)
        worksheet.set_column('F:F', 15, center_format)
        headers = [
            translations[language]["category"], translations[language]["score"], translations[language]["priority"],
            translations[language]["effort"], translations[language]["action_plan"], translations[language]["type"]
        ]
        for col, header in enumerate(headers):
            worksheet.write(15, col, header, header_format)
        for row in range(16, 16 + len(action_plan_df)):
            worksheet.write(row, 0, action_plan_df[translations[language]["category"]][row-16], 
                           cell_format if (row-16) % 2 else alt_row_format)
            if action_plan_df[translations[language]["score"]][row-16] != "":
                worksheet.conditional_format(f'B{row}:B{row}', {
                    'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100, 'format': critical_format
                })
                worksheet.conditional_format(f'B{row}:B{row}', {
                    'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
                    'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-0.01)/100, 'format': improvement_format
                })
                worksheet.conditional_format(f'B{row}:B{row}', {
                    'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100, 'format': good_format
                })
            worksheet.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 'value': translations[language]["high_priority"], 'format': critical_format
            })
            worksheet.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 'value': translations[language]["medium_priority"], 'format': improvement_format
            })
        worksheet.freeze_panes(16, 0)

        # Performance Overview Section (Rows 42-55)
        logger.debug("Writing Performance Overview Section to %s", sheet_name)
        worksheet.merge_range('A42:G42', translations[language]["chart"], title_format)
        chart_data = df_display[[translations[language]["percent"]]].reset_index()
        chart_data[translations[language]["score"]] = chart_data[translations[language]["percent"]] / 100
        chart_data.to_excel(writer, sheet_name=sheet_name, startrow=43, startcol=0, index=False)
        logger.debug("Wrote chart data to %s at row 43", sheet_name)
        worksheet.write('A44', translations[language]["category"], header_format)
        worksheet.write('B44', translations[language]["score"], header_format)
        bar_chart = workbook.add_chart({'type': 'bar'})
        bar_chart.add_series({
            'name': translations[language]["score"],
            'categories': f"='{sheet_name}'!$A$45:$A${44 + len(chart_data)}",
            'values': f"='{sheet_name}'!$B$45:$B${44 + len(chart_data)}",
            'fill': {'color': '#1E88E5'},
            'data_labels': {'value': True, 'num_format': '0.0%'}
        })
        bar_chart.set_title({'name': translations[language]["chart_title"]})
        bar_chart.set_x_axis({'name': translations[language]["score"], 'num_format': '0%'})
        bar_chart.set_y_axis({'name': translations[language]["category"], 'reverse': True})
        bar_chart.set_size({'width': 400, 'height': 150})
        worksheet.insert_chart('D45', bar_chart)
        logger.debug("Inserted bar chart to %s at D45", sheet_name)

        # Contact Section (Rows 57-62)
        logger.debug("Writing Contact Section to %s", sheet_name)
        worksheet.merge_range('A57:G57', translations[language]["contact"], title_format)
        contact_df = pd.DataFrame({
            translations[language]["metric"]: ["Email" if language == "English" else "Correo", "Website" if language == "English" else "Sitio Web"],
            translations[language]["value"]: [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=sheet_name, startrow=58, startcol=0, index=False)
        logger.debug("Wrote Contact to %s at row 58", sheet_name)
        worksheet.write('A59', translations[language]["metric"], header_format)
        worksheet.write('B59', translations[language]["value"], header_format)
        for row in range(59, 59 + len(contact_df)):
            worksheet.write(row, 0, contact_df[translations[language]["metric"]][row-59], cell_format)
        worksheet.merge_range('A61:G62', translations[language]["marketing_message"], wrap_format)

        # Verify single worksheet
        worksheets = workbook.worksheets()
        logger.debug("Total worksheets: %d", len(worksheets))
        if len(worksheets) > 1:
            logger.error("Multiple worksheets detected: %s", [ws.get_name() for ws in worksheets])
            raise ValueError(f"Multiple worksheets created: {[ws.get_name() for ws in worksheets]}")

    logger.debug("Excel report generation completed with single worksheet: %s", sheet_name)
    excel_output.seek(0)
    return excel_output

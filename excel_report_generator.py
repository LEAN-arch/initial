import pandas as pd
import io
import xlsxwriter
from typing import Dict, Any
import numpy as np
import logging

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
    TRANSLATIONS: Dict,
    SCORE_THRESHOLDS: Dict,
    CONFIG: Dict,
    overall_score: float,
    grade: str,
    REPORT_DATE: str
) -> io.BytesIO:
    """
    Generate a professional Excel report with accurate English and Spanish translations,
    advanced visualizations, contact information on the cover sheet, and actionable insights.
    
    Args:
        df: DataFrame with category scores and priorities
        df_display: DataFrame with display-friendly category names
        questions: Dictionary of questions by category and language
        responses: Dictionary of user responses
        language: Selected language ("Español" or "English")
        category_mapping: Mapping of display categories to internal categories
        TRANSLATIONS: Translation dictionary
        SCORE_THRESHOLDS: Thresholds for score categories
        CONFIG: Configuration dictionary with contact info
        overall_score: Overall audit score
        grade: Overall grade
        REPORT_DATE: Report generation date
    
    Returns:
        io.BytesIO: Excel file buffer
    """
    logger.debug("Starting Excel report generation with language: %s", language)

    # Validate language input
    if language not in ["English", "Español"]:
        logger.error("Unsupported language: %s", language)
        raise ValueError(f"Unsupported language: {language}. Must be 'English' or 'Español'")

    # Comprehensive translation dictionary
    TRANSLATIONS = {
        "English": {
            "report_title": "LEAN 2.0 Workplace Audit Report",
            "summary": "Executive Summary",
            "results": "Results",
            "actionable_charts": "Visualizations",
            "contact": "Contact",
            "category": "Category",
            "score": "Score",
            "score_percent": "Score (%)",
            "percent": "Percentage",
            "priority": "Priority",
            "high_priority": "High Priority",
            "medium_priority": "Medium Priority",
            "low_priority": "Low Priority",
            "high_priority_categories": "High Priority Categories",
            "overall_score": "Overall Score",
            "grade": "Grade",
            "question": "Question",
            "suggestion": "Suggestion",
            "chart_title": "Category Performance Overview",
            "marketing_message": "Contact LEAN 2.0 Institute to transform your workplace with ethical and sustainable practices.",
            "date_format": "%m/%d/%Y"
        },
        "Español": {
            "report_title": "Informe de Auditoría del Lugar de Trabajo LEAN 2.0",
            "summary": "Resumen Ejecutivo",
            "results": "Resultados",
            "actionable_charts": "Visualizaciones",
            "contact": "Contacto",
            "category": "Categoría",
            "score": "Puntuación",
            "score_percent": "Puntuación (%)",
            "percent": "Porcentaje",
            "priority": "Prioridad",
            "high_priority": "Alta Prioridad",
            "medium_priority": "Prioridad Media",
            "low_priority": "Baja Prioridad",
            "high_priority_categories": "Categorías de Alta Prioridad",
            "overall_score": "Puntuación General",
            "grade": "Calificación",
            "question": "Pregunta",
            "suggestion": "Sugerencia",
            "chart_title": "Resumen de Desempeño por Categoría",
            "marketing_message": "Contacte al Instituto LEAN 2.0 para transformar su lugar de trabajo con prácticas éticas y sostenibles.",
            "date_format": "%d/%m/%Y"
        }
    }

    # Format date according to language
    from datetime import datetime
    report_date_formatted = datetime.strptime(REPORT_DATE, "%Y-%m-%d").strftime(TRANSLATIONS[language]["date_format"])
    logger.debug("Formatted report date: %s", report_date_formatted)

    # Enhanced Data Validation
    logger.debug("Validating input data")
    if df.empty or df_display.empty or not responses:
        logger.error("Input data is empty or invalid")
        raise ValueError("Input data is empty or invalid")
    required_columns = [TRANSLATIONS[language]["score"], TRANSLATIONS[language]["percent"], TRANSLATIONS[language]["priority"]]
    if not all(col in df.columns for col in required_columns):
        logger.error("Required columns missing in df: %s", required_columns)
        raise ValueError(f"Required columns missing in df: {required_columns}")
    if not df.index.is_unique:
        logger.error("DataFrame index contains duplicates")
        raise ValueError("DataFrame index must be unique for categories")
    if not all(cat in df.index for cat in questions.keys()):
        logger.error("Not all question categories are present in DataFrame index")
        raise ValueError("Not all question categories are present in DataFrame index")
    if not pd.api.types.is_numeric_dtype(df[TRANSLATIONS[language]["percent"]]):
        logger.error("Column %s must contain numeric values", TRANSLATIONS[language]["percent"])
        raise ValueError(f"Column {TRANSLATIONS[language]['percent']} must contain numeric values")
    if df[TRANSLATIONS[language]["percent"]].isnull().any():
        logger.error("Column %s contains missing values", TRANSLATIONS[language]["percent"])
        raise ValueError(f"Column {TRANSLATIONS[language]['percent']} contains missing values")
    if not (df[TRANSLATIONS[language]["percent"]] >= 0).all() or not (df[TRANSLATIONS[language]["percent"]] <= 100).all():
        logger.error("Column %s contains values outside 0-100 range", TRANSLATIONS[language]["percent"])
        raise ValueError(f"Column {TRANSLATIONS[language]['percent']} must contain values between 0 and 100")
    if len(df) != len(df_display):
        logger.error("Mismatch in row counts between df (%d) and df_display (%d)", len(df), len(df_display))
        raise ValueError(f"Mismatch in row counts between df ({len(df)}) and df_display ({len(df_display)})")
    for cat in responses:
        if not responses[cat] or any(score is None for score in responses[cat]):
            logger.error("Invalid or missing responses in category %s", cat)
            raise ValueError(f"Invalid or missing responses in category {cat}")
    # Validate translation keys
    required_keys = ["report_title", "summary", "results", "actionable_charts", "contact", "category",
                    "score", "score_percent", "percent", "priority", "high_priority", "medium_priority",
                    "low_priority", "high_priority_categories", "overall_score", "grade", "question",
                    "suggestion", "chart_title", "marketing_message", "date_format"]
    if not all(key in TRANSLATIONS[language] for key in required_keys):
        missing = [key for key in required_keys if key not in TRANSLATIONS[language]]
        logger.error("Missing translation keys for %s: %s", language, missing)
        raise ValueError(f"Missing translation keys for {language}: {missing}")

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Define formats
        title_format = workbook.add_format({
            'bold': True, 'font_size': 16, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        subtitle_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'font_name': 'Arial', 'align': 'left', 'valign': 'vcenter',
            'font_color': '#424242', 'border': 1
        })
        header_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        cell_format = workbook.add_format({
            'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top', 'text_wrap': True
        })
        percent_format = workbook.add_format({
            'num_format': '0.0%', 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
        })
        number_format = workbook.add_format({
            'num_format': '0.00', 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
        })
        bold_format = workbook.add_format({
            'bold': True, 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
        })
        wrap_format = workbook.add_format({
            'text_wrap': True, 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
        })
        center_format = workbook.add_format({
            'align': 'center', 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
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
            'bg_color': '#F5F5F5', 'font_size': 10, 'font_name': 'Arial', 'border': 1, 'valign': 'top'
        })

        # Cover Sheet
        logger.debug("Generating Cover Sheet")
        cover_sheet = workbook.add_worksheet("Cover")
        cover_sheet.merge_range('A1:F1', TRANSLATIONS[language]["report_title"], title_format)
        cover_sheet.merge_range('A2:F2', "Auditoría del Lugar de Trabajo Ético" if language == "Español" else "Ethical Lean Workplace Audit", subtitle_format)
        cover_sheet.merge_range('A3:F3', f"Generado el: {report_date_formatted}" if language == "Español" else f"Generated on: {report_date_formatted}", subtitle_format)
        cover_sheet.merge_range('A4:F4', "Preparado por: Instituto LEAN 2.0" if language == "Español" else "Prepared by: LEAN 2.0 Institute", subtitle_format)
        cover_sheet.merge_range('A5:F5', "Contáctenos:" if language == "Español" else "Contact Us:", bold_format)
        cover_sheet.merge_range('A6:F6', f"Correo: {CONFIG['contact']['email']} | Sitio Web: {CONFIG['contact']['website']}" if language == "Español" else f"Email: {CONFIG['contact']['email']} | Website: {CONFIG['contact']['website']}", cell_format)
        cover_sheet.merge_range('A7:F7', "Misión: Transformar los lugares de trabajo con prácticas éticas, inclusivas y sostenibles." if language == "Español" else "Mission: Transforming workplaces through ethical, inclusive, and sustainable practices.", wrap_format)
        cover_sheet.write('A9', "Tabla de Contenidos:" if language == "Español" else "Table of Contents:", bold_format)
        toc_links = [
            ("Resumen Ejecutivo" if language == "Español" else "Executive Summary", TRANSLATIONS[language]["summary"]),
            ("Perspectivas Estadísticas" if language == "Español" else "Statistical Insights", "Statistical Insights"),
            ("Resultados" if language == "Español" else "Results", TRANSLATIONS[language]["results"]),
            ("Plan de Acción" if language == "Español" else "Action Plan", "Action Plan"),
            ("Visualizaciones" if language == "Español" else "Visualizations", TRANSLATIONS[language]["actionable_charts"]),
            ("Contacto" if language == "Español" else "Contact", TRANSLATIONS[language]["contact"])
        ]
        for idx, (name, sheet) in enumerate(toc_links, start=10):
            cover_sheet.write_url(f'A{idx}', f"internal:'{sheet}'!A1", 
                               string=name, cell_format=workbook.add_format({
                                   'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'
                               }))
        cover_sheet.write('A17', "Placeholder para el Logo: Inserte el logo del Instituto LEAN 2.0" if language == "Español" else "Logo Placeholder: Insert LEAN 2.0 Institute logo", bold_format)
        cover_sheet.set_column('A:F', 25)

        # Executive Summary Sheet
        logger.debug("Generating Executive Summary Sheet")
        critical_count = len(df[df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
        improvement_count = len(df[
            (df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["CRITICAL"]) & 
            (df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"])
        ])
        good_count = len(df[df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]])
        summary_data = {
            TRANSLATIONS[language]["overall_score"]: [f"{overall_score:.1f}%"],
            TRANSLATIONS[language]["grade"]: [grade],
            TRANSLATIONS[language]["high_priority_categories"]: [critical_count],
            "Categorías que Necesitan Mejora" if language == "Español" else "Categories Needing Improvement": [improvement_count],
            "Categorías con Buen Desempeño" if language == "Español" else "Categories Performing Well": [good_count],
            "Próximos Pasos" if language == "Español" else "Next Steps": ["Revise el Plan de Acción para recomendaciones priorizadas." if language == "Español" else "Review the Action Plan for prioritized recommendations."]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["summary"], index=False, startrow=4)
        worksheet_summary = writer.sheets[TRANSLATIONS[language]["summary"]]
        worksheet_summary.merge_range('A1:F1', TRANSLATIONS[language]["summary"], title_format)
        worksheet_summary.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_summary.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                                  cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_summary.set_column('A:A', 30, cell_format)
        worksheet_summary.set_column('B:E', 20, center_format)
        worksheet_summary.set_column('F:F', 60, wrap_format)
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet_summary.write(4, col_num, value, header_format)
        worksheet_summary.write('A4', "Métrica" if language == "Español" else "Metric", header_format)
        for row in range(5, 5 + len(summary_df)):
            worksheet_summary.write(row, 0, summary_df.columns[row-5], bold_format)
            worksheet_summary.write(row, 0, row-5+1, center_format)
            worksheet_summary.write(row, 0, "", alt_row_format if (row-5) % 2 else cell_format)
        worksheet_summary.conditional_format('B6:B6', {
            'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100,
            'format': critical_format
        })
        worksheet_summary.conditional_format('B6:B6', {
            'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
            'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1)/100, 'format': improvement_format
        })
        worksheet_summary.conditional_format('B6:B6', {
            'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100,
            'format': good_format
        })
        worksheet_summary.freeze_panes(5, 0)

        # Statistical Insights Sheet
        logger.debug("Generating Statistical Insights Sheet")
        stats_data = {
            TRANSLATIONS[language]["category"]: df_display.index,
            TRANSLATIONS[language]["score"]: df_display[TRANSLATIONS[language]["percent"]],
            "Desviación Estándar" if language == "Español" else "Standard Deviation": [np.std(responses[cat]) for cat in responses],
            "Puntuación Mínima" if language == "Español" else "Min Score": [min(responses[cat]) for cat in responses],
            "Puntuación Máxima" if language == "Español" else "Max Score": [max(responses[cat]) for cat in responses]
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name="Perspectivas Estadísticas" if language == "Español" else "Statistical Insights", index=False, startrow=3)
        worksheet_stats = writer.sheets["Perspectivas Estadísticas" if language == "Español" else "Statistical Insights"]
        worksheet_stats.merge_range('A1:E1', "Perspectivas Estadísticas" if language == "Español" else "Statistical Insights", title_format)
        worksheet_stats.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_stats.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                                cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_stats.set_column('A:A', 35, cell_format)
        worksheet_stats.set_column('B:B', 15, percent_format)
        worksheet_stats.set_column('C:E', 15, number_format)
        for col_num, value in enumerate(stats_df.columns.values):
            worksheet_stats.write(3, col_num, value, header_format)
        for row in range(4, 4 + len(stats_df)):
            worksheet_stats.write(row, 0, stats_df[TRANSLATIONS[language]["category"]][row-4], cell_format if (row-4) % 2 else alt_row_format)
            worksheet_stats.conditional_format(f'B{row}:B{row}', {
                'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100,
                'format': critical_format
            })
            worksheet_stats.conditional_format(f'B{row}:B{row}', {
                'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
                'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1)/100, 'format': improvement_format
            })
            worksheet_stats.conditional_format(f'B{row}:B{row}', {
                'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100,
                'format': good_format
            })
        worksheet_stats.freeze_panes(4, 0)

        # Results Sheet
        logger.debug("Generating Results Sheet")
        results_df = df_display.copy()
        results_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["results"], startrow=3)
        worksheet_results = writer.sheets[TRANSLATIONS[language]["results"]]
        worksheet_results.merge_range('A1:D1', TRANSLATIONS[language]["results"], title_format)
        worksheet_results.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_results.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                                  cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_results.set_column('A:A', 35, cell_format)
        worksheet_results.set_column('B:D', 15, center_format)
        for col_num, value in enumerate(results_df.columns.values):
            worksheet_results.write(3, col_num + 1, value, header_format)
        worksheet_results.write(3, 0, TRANSLATIONS[language]["category"], header_format)
        for row in range(4, 4 + len(results_df)):
            worksheet_results.write(row, 0, results_df.index[row-4], cell_format if (row-4) % 2 else alt_row_format)
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100,
                'format': critical_format
            })
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
                'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1)/100, 'format': improvement_format
            })
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100,
                'format': good_format
            })
        worksheet_results.freeze_panes(4, 0)

        # Action Plan Sheet
        logger.debug("Generating Action Plan Sheet")
        action_plan_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            try:
                score = df.loc[cat, TRANSLATIONS[language]["percent"]].item()
                logger.debug("Retrieved score for category %s: %s", cat, score)
            except (ValueError, KeyError) as e:
                logger.error("Failed to retrieve scalar score for category %s: %s", cat, str(e))
                raise ValueError(f"Failed to retrieve scalar score for category {cat}: {str(e)}")
            priority = (
                TRANSLATIONS[language]["high_priority"] if score < SCORE_THRESHOLDS["CRITICAL"] else
                TRANSLATIONS[language]["medium_priority"] if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
                TRANSLATIONS[language]["low_priority"]
            )
            if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                impact_score = (100 - score) / 100
                feasibility_score = score / 100
                priority_score = impact_score * 0.7 + feasibility_score * 0.3
                effort = (
                    "Alto" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                    "Medio" if score < SCORE_THRESHOLDS["CRITICAL"] else
                    "Bajo" if language == "Español" else
                    "High" if score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                    "Medium" if score < SCORE_THRESHOLDS["CRITICAL"] else
                    "Low"
                )
                action_plan_data.append({
                    "Category": display_cat,
                    "Score": f"{score:.1f}%",
                    "Priority": priority,
                    "Priority Score": priority_score,
                    "Estimated Effort": effort,
                    "Action Plan": f"Se requiere acción inmediata para abordar el bajo desempeño en {display_cat}." if language == "Español" else f"Immediate action required to address low performance in {display_cat}.",
                    "Type": "Categoría" if language == "Español" else "Category"
                })
                for idx, q_score in enumerate(responses[cat]):
                    if q_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        question, _, rec = questions[cat][language][idx]
                        q_impact_score = (100 - q_score) / 100
                        q_priority_score = q_impact_score * 0.7 + feasibility_score * 0.3
                        q_effort = (
                            "Alto" if q_score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                            "Medio" if q_score < SCORE_THRESHOLDS["CRITICAL"] else
                            "Bajo" if language == "Español" else
                            "High" if q_score < SCORE_THRESHOLDS["CRITICAL"]/2 else
                            "Medium" if q_score < SCORE_THRESHOLDS["CRITICAL"] else
                            "Low"
                        )
                        action_plan_data.append({
                            "Category": "",
                            "Score": "",
                            "Priority": "",
                            "Priority Score": q_priority_score,
                            "Estimated Effort": q_effort,
                            "Action Plan": f"{TRANSLATIONS[language]['question']}: {question[:100]}{'...' if len(question) > 100 else ''}\n"
                                         f"{TRANSLATIONS[language]['suggestion']}: {rec}",
                            "Type": "Pregunta" if language == "Español" else "Question"
                        })
        action_plan_df = pd.DataFrame(action_plan_data).sort_values(["Priority", "Priority Score"], ascending=[True, False])
        action_plan_df.to_excel(writer, sheet_name="Plan de Acción" if language == "Español" else "Action Plan", index=False, startrow=3)
        worksheet_action = writer.sheets["Plan de Acción" if language == "Español" else "Action Plan"]
        worksheet_action.merge_range('A1:F1', "Plan de Acción" if language == "Español" else "Action Plan", title_format)
        worksheet_action.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_action.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                                cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_action.set_column('A:A', 35, cell_format)
        worksheet_action.set_column('B:C', 15, center_format)
        worksheet_action.set_column('D:D', 15, number_format)
        worksheet_action.set_column('E:E', 15, center_format)
        worksheet_action.set_column('F:F', 80, wrap_format)
        action_plan_columns = [
            TRANSLATIONS[language]["category"], TRANSLATIONS[language]["score"], TRANSLATIONS[language]["priority"],
            "Puntuación de Prioridad" if language == "Español" else "Priority Score",
            "Esfuerzo Estimado" if language == "Español" else "Estimated Effort",
            "Plan de Acción" if language == "Español" else "Action Plan"
        ]
        for col_num, value in enumerate(action_plan_columns):
            worksheet_action.write(3, col_num, value, header_format)
        for row in range(4, 4 + len(action_plan_df)):
            worksheet_action.write(row, 0, action_plan_df["Category"][row-4], cell_format if (row-4) % 2 else alt_row_format)
            worksheet_action.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 
                'value': TRANSLATIONS[language]["high_priority"], 'format': critical_format
            })
            worksheet_action.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 
                'value': TRANSLATIONS[language]["medium_priority"], 'format': improvement_format
            })
        worksheet_action.freeze_panes(4, 0)

        # Visualizations Sheet
        logger.debug("Generating Visualizations Sheet")
        worksheet_viz = workbook.add_worksheet(TRANSLATIONS[language]["actionable_charts"])
        worksheet_viz.merge_range('A1:I1', TRANSLATIONS[language]["actionable_charts"], title_format)
        worksheet_viz.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_viz.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                              cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))

        # Stacked Bar Chart
        logger.debug("Generating Stacked Bar Chart")
        stacked_data = df_display[[TRANSLATIONS[language]["percent"]]].reset_index()
        stacked_data["Critical Gap"] = stacked_data[TRANSLATIONS[language]["percent"]].apply(
            lambda x: max(0, SCORE_THRESHOLDS["CRITICAL"]/100 - x)
        )
        stacked_data["Improvement Gap"] = stacked_data[TRANSLATIONS[language]["percent"]].apply(
            lambda x: max(0, min((SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-SCORE_THRESHOLDS["CRITICAL"])/100, SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100 - x - stacked_data["Critical Gap"]))
        )
        stacked_data["Achieved"] = stacked_data[TRANSLATIONS[language]["percent"]]/100
        stacked_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=0, index=False)
        worksheet_viz.set_column('A:A', 35, cell_format)
        worksheet_viz.set_column('B:D', 15, percent_format)
        worksheet_viz.write('A5', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('B5', "Puntuación Alcanzada" if language == "Español" else "Achieved Score", header_format)
        worksheet_viz.write('C5', "Brecha Crítica (<50%)" if language == "Español" else "Critical Gap (<50%)", header_format)
        worksheet_viz.write('D5', "Brecha de Mejora (50-69%)" if language == "Español" else "Improvement Gap (50-69%)", header_format)
        stacked_chart = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})
        stacked_chart.add_series({
            'name': TRANSL subcribe[language]["score_percent"],
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$A$6:$A${5 + len(stacked_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$B$6:$B${5 + len(stacked_data)}",
            'fill': {'color': '#43A047'},
            'data_labels': {'value': True, 'num_format': '0.0%', 'font': {'size': 10, 'name': 'Arial'}}
        })
        stacked_chart.add_series({
            'name': "Brecha Crítica (<50%)" if language == "Español" else "Critical Gap (<50%)",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$C$6:$C${5 + len(stacked_data)}",
            'fill': {'color': '#D32F2F'},
            'data_labels': {'value': False}
        })
        stacked_chart.add_series({
            'name': "Brecha de Mejora (50-69%)" if language == "Español" else "Improvement Gap (50-69%)",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$D$6:$D${5 + len(stacked_data)}",
            'fill': {'color': '#FFD54F'},
            'data_labels': {'value': False}
        })
        stacked_chart.set_title({'name': TRANSLATIONS[language]["chart_title"], 'name_font': {'size': 12, 'name': 'Arial'}})
        stacked_chart.set_x_axis({
            'name': TRANSLATIONS[language]["score_percent"],
            'min': 0, 'max': 1, 'num_format': '0%',
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        stacked_chart.set_y_axis({
            'name': TRANSLATIONS[language]["category"],
            'reverse': True,
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        stacked_chart.set_legend({'position': 'bottom', 'font': {'size': 10, 'name': 'Arial'}})
        stacked_chart.set_size({'width': 600, 'height': 300})
        worksheet_viz.insert_chart('G5', stacked_chart)

        # Visual Summary Table
        logger.debug("Generating Visual Summary Table")
        summary_stats = {
            "Métrica" if language == "Español" else "Metric": [
                "Puntuación Media" if language == "Español" else "Mean Score",
                "Varianza" if language == "Español" else "Variance",
                "Categorías Críticas" if language == "Español" else "Critical Categories"
            ],
            "Valor" if language == "Español" else "Value": [
                f"{df_display[TRANSLATIONS[language]['percent']].mean().item():.1f}%",
                f"{df_display[TRANSLATIONS[language]['percent']].var().item():.2f}",
                critical_count
            ]
        }
        summary_stats_df = pd.DataFrame(summary_stats)
        summary_stats_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=4, index=False)
        worksheet_viz.set_column('E:F', 20, cell_format)
        worksheet_viz.write('E5', "Métrica" if language == "Español" else "Metric", header_format)
        worksheet_viz.write('F5', "Valor" if language == "Español" else "Value", header_format)
        for row in range(5, 5 + len(summary_stats_df)):
            worksheet_viz.write(row, 4, summary_stats_df["Métrica" if language == "Español" else "Metric"][row-5], cell_format if (row-5) % 2 else alt_row_format)

        # Box Plot
        logger.debug("Generating Box Plot")
        box_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            for q_score in responses[cat]:
                box_data.append({
                    TRANSLATIONS[language]["category"]: display_cat,
                    "Score": q_score / 100
                })
        box_df = pd.DataFrame(box_data)
        box_df_pivot = box_df.pivot(columns=TRANSLATIONS[language]["category"], values="Score")
        box_df_pivot.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=20, startcol=0, index=False)
        worksheet_viz.set_column('A:G', 15, percent_format)
        for col_num, value in enumerate(box_df_pivot.columns, 1):
            worksheet_viz.write(20, col_num, value, header_format)
        box_chart = workbook.add_chart({'type': 'box'})
        for col_idx, cat in enumerate(box_df_pivot.columns, 1):
            box_chart.add_series({
                'name': cat,
                'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!${chr(65+col_idx)}$21:${chr(65+col_idx)}${20 + len(box_df_pivot)}",
                'fill': {'color': '#1E88E5'},
                'line': {'color': '#1E88E5'},
                'whisker': {'color': '#1E88E5'},
                'median': {'color': '#D32F2F'},
                'mean': {'color': '#FFD54F', 'type': 'diamond', 'size': 6}
            })
        box_chart.set_title({'name': "Distribución de Puntuaciones por Pregunta por Categoría" if language == "Español" else "Question Score Distribution by Category", 'name_font': {'size': 12, 'name': 'Arial'}})
        box_chart.set_x_axis({
            'name': TRANSLATIONS[language]["category"],
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        box_chart.set_y_axis({
            'name': TRANSLATIONS[language]["score_percent"],
            'min': 0, 'max': 1, 'num_format': '0%',
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        box_chart.set_size({'width': 600, 'height': 300})
        worksheet_viz.insert_chart('G21', box_chart)

        # Scatter Plot for Prioritization
        logger.debug("Generating Scatter Plot")
        scatter_data = action_plan_df[["Category", "Priority Score", "Estimated Effort"]].copy()
        scatter_data["Effort Value"] = scatter_data["Estimated Effort"].map({
            "Low": 1, "Medium": 2, "High": 3, "Bajo": 1, "Medio": 2, "Alto": 3
        })
        scatter_data = scatter_data[scatter_data["Category"] != ""]
        scatter_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=36, startcol=0, index=False)
        worksheet_viz.set_column('A:A', 35, cell_format)
        worksheet_viz.set_column('B:B', 15, number_format)
        worksheet_viz.set_column('C:D', 15, center_format)
        worksheet_viz.write('A37', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('B37', "Puntuación de Prioridad" if language == "Español" else "Priority Score", header_format)
        worksheet_viz.write('C37', "Esfuerzo Estimado" if language == "Español" else "Estimated Effort", header_format)
        worksheet_viz.write('D37', "Valor de Esfuerzo" if language == "Español" else "Effort Value", header_format)
        scatter_chart = workbook.add_chart({'type': 'scatter'})
        scatter_chart.add_series({
            'name': "Prioridad vs. Esfuerzo" if language == "Español" else "Priority vs Effort",
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$A$38:$A${37 + len(scatter_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$B$38:$B${37 + len(scatter_data)}",
            'y2_values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$D$38:$D${37 + len(scatter_data)}",
            'marker': {
                'type': 'circle',
                'size': 8,
                'fill': {'color': '#1E88E5'},
                'border': {'color': '#1E88E5'}
            }
        })
        scatter_chart.set_title({'name': "Análisis de Prioridad vs. Esfuerzo" if language == "Español" else "Priority vs. Effort Analysis", 'name_font': {'size': 12, 'name': 'Arial'}})
        scatter_chart.set_x_axis({
            'name': "Puntuación de Prioridad" if language == "Español" else "Priority Score",
            'min': 0, 'max': 1,
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        scatter_chart.set_y_axis({
            'name': "Esfuerzo (1=Bajo, 3=Alto)" if language == "Español" else "Effort (1=Low, 3=High)",
            'min': 0.5, 'max': 3.5,
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        scatter_chart.set_size({'width': 600, 'height': 300})
        worksheet_viz.insert_chart('G37', scatter_chart)

        # Heatmap
        logger.debug("Generating Heatmap")
        question_scores = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            for idx, (q, _, _) in enumerate(questions[cat][language]):
                question_scores.append({
                    TRANSLATIONS[language]["category"]: display_cat,
                    TRANSLATIONS[language]["question"]: q[:50] + ("..." if len(q) > 50 else ""),
                    TRANSLATIONS[language]["score"]: responses[cat][idx] / 100
                })
        heatmap_df = pd.DataFrame(question_scores)
        heatmap_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=10, index=False)
        worksheet_viz.set_column('K:M', 30, cell_format)
        worksheet_viz.write('K5', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('L5', TRANSLATIONS[language]["question"], header_format)
        worksheet_viz.write('M5', TRANSLATIONS[language]["score_percent"], header_format)
        for row in range(5, 5 + len(heatmap_df)):
            worksheet_viz.write(row, 10, heatmap_df[TRANSLATIONS[language]["category"]][row-5], cell_format if (row-5) % 2 else alt_row_format)
            worksheet_viz.conditional_format(f'M{row}:M{row}', {
                'type': '3_color_scale',
                'min_color': '#D32F2F',
                'mid_color': '#FFD54F',
                'max_color': '#43A047'
            })

        # Contact Sheet
        logger.debug("Generating Contact Sheet")
        contact_df = pd.DataFrame({
            "Método de Contacto" if language == "Español" else "Contact Method": ["Correo" if language == "Español" else "Email", "Sitio Web" if language == "Español" else "Website"],
            "Detalles" if language == "Español" else "Details": [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["contact"], index=False, startrow=3)
        worksheet_contact = writer.sheets[TRANSLATIONS[language]["contact"]]
        worksheet_contact.merge_range('A1:B1', TRANSLATIONS[language]["contact"], title_format)
        worksheet_contact.write('A2', f"Fecha: {report_date_formatted}" if language == "Español" else f"Date: {report_date_formatted}", subtitle_format)
        worksheet_contact.write_url('A3', "internal:Cover!A1", string="Volver a la Portada" if language == "Español" else "Back to Cover", 
                                 cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_contact.set_column('A:A', 20, cell_format)
        worksheet_contact.set_column('B:B', 50, cell_format)
        for col_num, value in enumerate(contact_df.columns.values):
            worksheet_contact.write(3, col_num, value, header_format)
        for row in range(4, 4 + len(contact_df)):
            worksheet_contact.write(row, 0, contact_df["Método de Contacto" if language == "Español" else "Contact Method"][row-4], cell_format if (row-4) % 2 else alt_row_format)
        worksheet_contact.write('A7', "Colabore con Nosotros" if language == "Español" else "Collaborate with Us", bold_format)
        worksheet_contact.write('A8', TRANSLATIONS[language]["marketing_message"], wrap_format)

    logger.debug("Excel report generation completed")
    excel_output.seek(0)
    return excel_output

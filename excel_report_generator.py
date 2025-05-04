import pandas as pd
import io
import xlsxwriter
from typing import Dict, Any
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
    Generate a professional single-sheet Excel report with accurate translations,
    compelling visualizations, and a persuasive narrative to contract LEAN 2.0 Institute services.

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

    # Translation dictionary
    translations = {
        "English": {
            "report_title": "LEAN 2.0 Workplace Audit Report",
            "summary": "Executive Summary",
            "results": "Results & Action Plan",
            "actionable_charts": "Performance Overview",
            "contact": "Contact Us",
            "why_choose": "Why Choose LEAN 2.0 Institute?",
            "category": "Category",
            "score": "Score",
            "score_percent": "Score (%)",
            "percent": "Percentage",
            "priority": "Priority",
            "high_priority": "High",
            "medium_priority": "Medium",
            "low_priority": "Low",
            "high_priority_categories": "Critical Categories",
            "overall_score": "Overall Score",
            "grade": "Grade",
            "question": "Question",
            "suggestion": "Recommendation",
            "chart_title": "Category Performance",
            "marketing_message": (
                "Transform your workplace with LEAN 2.0 Institute! Our proven methodologies and tailored solutions "
                "empower organizations to achieve ethical, inclusive, and sustainable workplaces. Contact us today to start your journey."
            ),
            "why_choose_message": (
                "- Proven Expertise: Decades of experience in lean and ethical workplace transformations.\n"
                "- Tailored Solutions: Customized strategies to address your unique challenges.\n"
                "- Sustainable Impact: Long-term improvements in employee well-being and operational efficiency.\n"
                "- Global Reach: Supporting organizations worldwide with localized expertise."
            ),
            "date_format": "%m/%d/%Y",
            "table_of_contents": "Table of Contents",
            "metric": "Metric",
            "value": "Value",
            "action_plan": "Action Plan",
            "estimated_effort": "Effort",
            "type": "Type",
            "category_type": "Category",
            "question_type": "Question",
            "next_steps": "Next Steps",
            "prepared_by": "Prepared by: LEAN 2.0 Institute",
            "mission": "Mission: Fostering ethical, inclusive, and sustainable workplaces worldwide."
        },
        "Español": {
            "report_title": "Informe de Auditoría del Lugar de Trabajo LEAN 2.0",
            "summary": "Resumen Ejecutivo",
            "results": "Resultados y Plan de Acción",
            "actionable_charts": "Resumen de Desempeño",
            "contact": "Contáctenos",
            "why_choose": "¿Por qué Elegir el Instituto LEAN 2.0?",
            "category": "Categoría",
            "score": "Puntuación",
            "score_percent": "Puntuación (%)",
            "percent": "Porcentaje",
            "priority": "Prioridad",
            "high_priority": "Alta",
            "medium_priority": "Media",
            "low_priority": "Baja",
            "high_priority_categories": "Categorías Críticas",
            "overall_score": "Puntuación General",
            "grade": "Calificación",
            "question": "Pregunta",
            "suggestion": "Recomendación",
            "chart_title": "Desempeño por Categoría",
            "marketing_message": (
                "¡Transforme su lugar de trabajo con el Instituto LEAN 2.0! Nuestras metodologías probadas y soluciones personalizadas "
                "empoderan a las organizaciones para lograr lugares de trabajo éticos, inclusivos y sostenibles. Contáctenos hoy para comenzar."
            ),
            "why_choose_message": (
                "- Experiencia Comprobada: Décadas de experiencia en transformaciones lean y éticas.\n"
                "- Soluciones Personalizadas: Estrategias adaptadas a sus desafíos únicos.\n"
                "- Impacto Sostenible: Mejoras a largo plazo en el bienestar de los empleados y la eficiencia operativa.\n"
                "- Alcance Global: Apoyo a organizaciones en todo el mundo con experiencia localizada."
            ),
            "date_format": "%d/%m/%Y",
            "table_of_contents": "Tabla de Contenidos",
            "metric": "Métrica",
            "value": "Valor",
            "action_plan": "Plan de Acción",
            "estimated_effort": "Esfuerzo",
            "type": "Tipo",
            "category_type": "Categoría",
            "question_type": "Pregunta",
            "next_steps": "Próximos Pasos",
            "prepared_by": "Preparado por: Instituto LEAN 2.0",
            "mission": "Misión: Fomentar lugares de trabajo éticos, inclusivos y sostenibles en todo el mundo."
        }
    }

    # Validate translation keys
    required_keys = [
        "report_title", "summary", "results", "actionable_charts", "contact", "why_choose", "category",
        "score", "score_percent", "percent", "priority", "high_priority", "medium_priority", "low_priority",
        "high_priority_categories", "overall_score", "grade", "question", "suggestion", "chart_title",
        "marketing_message", "why_choose_message", "date_format", "table_of_contents", "metric", "value",
        "action_plan", "estimated_effort", "type", "category_type", "question_type", "next_steps",
        "prepared_by", "mission"
    ]
    if not all(key in translations[language] for key in required_keys):
        missing = [key for key in required_keys if key not in translations[language]]
        logger.error("Missing translation keys for %s: %s", language, missing)
        raise ValueError(f"Missing translation keys for {language}: {missing}")

    # Format date
    try:
        report_date_formatted = datetime.strptime(REPORT_DATE, "%Y-%m-%d").strftime(translations[language]["date_format"])
        logger.debug("Formatted report date: %s", report_date_formatted)
    except ValueError as e:
        logger.error("Invalid date format for REPORT_DATE: %s", REPORT_DATE)
        raise ValueError(f"Invalid date format for REPORT_DATE: {REPORT_DATE}. Expected YYYY-MM-DD")

    # Enhanced data validation
    logger.debug("Validating input data")
    if df.empty or df_display.empty or not responses:
        logger.error("Input data is empty or invalid")
        raise ValueError("Input data is empty or invalid")
    
    required_columns = [translations[language]["score"], translations[language]["percent"], translations[language]["priority"]]
    if not all(col in df.columns for col in required_columns):
        logger.error("Required columns missing in df: %s", required_columns)
        raise ValueError(f"Required columns missing in df: {required_columns}")
    
    if not df.index.is_unique:
        logger.error("DataFrame index contains duplicates")
        raise ValueError("DataFrame index must be unique for categories")
    
    if not all(cat in df.index for cat in questions.keys()):
        logger.error("Not all question categories are present in DataFrame index")
        raise ValueError("Not all question categories are present in DataFrame index")
    
    if not pd.api.types.is_numeric_dtype(df[translations[language]["percent"]]):
        logger.error("Column %s must contain numeric values", translations[language]["percent"])
        raise ValueError(f"Column {translations[language]['percent']} must contain numeric values")
    
    if df[translations[language]["percent"]].isna().any():
        logger.error("Column %s contains missing values", translations[language]["percent"])
        raise ValueError(f"Column {translations[language]['percent']} contains missing values")
    
    if not (df[translations[language]["percent"]] >= 0).all() or not (df[translations[language]["percent"]] <= 100).all():
        logger.error("Column %s contains values outside 0-100 range", translations[language]["percent"])
        raise ValueError(f"Column {translations[language]['percent']} must contain values between 0 and 100")
    
    if len(df) != len(df_display):
        logger.error("Mismatch in row counts between df (%d) and df_display (%d)", len(df), len(df_display))
        raise ValueError(f"Mismatch in row counts between df ({len(df)}) and df_display ({len(df_display)})")
    
    for cat in responses:
        if not responses[cat] or any(score is None or pd.isna(score) for score in responses[cat]):
            logger.error("Invalid or missing responses in category %s", cat)
            raise ValueError(f"Invalid or missing responses in category {cat}")

    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Informe de Auditoría" if language == "Español" else "Audit Report")

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
        link_format = workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'
        })

        # Cover Section (Rows 1-10)
        logger.debug("Generating Cover Section")
        worksheet.merge_range('A1:I1', translations[language]["report_title"], title_format)
        worksheet.merge_range('A2:I2', "Auditoría del Lugar de Trabajo Ético" if language == "Español" else "Ethical Lean Workplace Audit", subtitle_format)
        worksheet.merge_range('A3:I3', f"Generado el: {report_date_formatted}" if language == "Español" else f"Generated on: {report_date_formatted}", subtitle_format)
        worksheet.merge_range('A4:I4', translations[language]["prepared_by"], subtitle_format)
        worksheet.merge_range('A5:I5', translations[language]["mission"], bold_format)
        worksheet.write('A7', translations[language]["table_of_contents"], bold_format)
        toc = [
            (translations[language]["summary"], "A12"),
            (translations[language]["results"], "A22"),
            (translations[language]["actionable_charts"], "A50"),
            (translations[language]["why_choose"], "A70"),
            (translations[language]["contact"], "A80")
        ]
        for idx, (name, cell) in enumerate(toc, start=8):
            worksheet.write_url(f'A{idx}', f"internal:'Informe de Auditoría'!{cell}" if language == "Español" else f"internal:'Audit Report'!{cell}", 
                              string=name, cell_format=link_format)
        worksheet.merge_range('A10:I10', "Placeholder para el Logo: Inserte el logo del Instituto LEAN 2.0 en formato PNG, 200x100 píxeles" if language == "Español" else 
                            "Logo Placeholder: Insert LEAN 2.0 Institute logo in PNG format, 200x100 pixels", bold_format)
        worksheet.set_column('A:I', 20, cell_format)

        # Executive Summary Section (Rows 12-20)
        logger.debug("Generating Executive Summary Section")
        worksheet.merge_range('A12:I12', translations[language]["summary"], title_format)
        critical_count = len(df[df[translations[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
        summary_data = {
            translations[language]["metric"]: [
                translations[language]["overall_score"],
                translations[language]["grade"],
                translations[language]["high_priority_categories"],
                translations[language]["next_steps"]
            ],
            translations[language]["value"]: [
                f"{overall_score:.1f}%",
                grade,
                critical_count,
                "Consulte el Plan de Acción para recomendaciones priorizadas." if language == "Español" else "Review the Action Plan for prioritized recommendations."
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="Informe de Auditoría" if language == "Español" else "Audit Report", startrow=13, startcol=0, index=False)
        worksheet.write('A14', translations[language]["metric"], header_format)
        worksheet.write('B14', translations[language]["value"], header_format)
        worksheet.set_column('A:A', 30)
        worksheet.set_column('B:B', 40)
        for row in range(14, 14 + len(summary_df)):
            worksheet.write(row, 0, summary_df[translations[language]["metric"]][row-14], bold_format if row == 14 else cell_format)
            worksheet.write(row, 1, summary_df[translations[language]["value"]][row-14], center_format)
            worksheet.write(row, 0, "", alt_row_format if (row-14) % 2 else cell_format)
        worksheet.conditional_format('B15:B15', {
            'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"]/100, 'format': critical_format
        })
        worksheet.conditional_format('B15:B15', {
            'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"]/100,
            'maximum': (SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-0.01)/100, 'format': improvement_format
        })
        worksheet.conditional_format('B15:B15', {
            'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]/100, 'format': good_format
        })

        # Results & Action Plan Section (Rows 22-48)
        logger.debug("Generating Results & Action Plan Section")
        worksheet.merge_range('A22:I22', translations[language]["results"], title_format)
        action_plan_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            try:
                score = df.loc[cat, translations[language]["percent"]].item()
                logger.debug("Retrieved score for category %s: %s", cat, score)
            except (ValueError, KeyError) as e:
                logger.error("Failed to retrieve scalar score for category %s: %s", cat, str(e))
                raise ValueError(f"Failed to retrieve scalar score for category {cat}: {str(e)}")
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
                translations[language]["estimated_effort"]: effort,
                translations[language]["action_plan"]: (
                    f"Mejorar el desempeño en {display_cat} con acciones inmediatas." if language == "Español" else
                    f"Enhance performance in {display_cat} with immediate actions."
                ),
                translations[language]["type"]: translations[language]["category_type"]
            })
            for idx, q_score in enumerate(responses[cat]):
                if q_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
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
                        translations[language]["estimated_effort"]: q_effort,
                        translations[language]["action_plan"]: (
                            f"{translations[language]['question']}: {question[:100]}{'...' if len(question) > 100 else ''}\n"
                            f"{translations[language]['suggestion']}: {rec}"
                        ),
                        translations[language]["type"]: translations[language]["question_type"]
                    })
        action_plan_df = pd.DataFrame(action_plan_data).sort_values(
            [translations[language]["score"], translations[language]["priority"]], ascending=[True, False]
        )
        action_plan_df.to_excel(writer, sheet_name="Informe de Auditoría" if language == "Español" else "Audit Report", 
                               startrow=23, startcol=0, index=False)
        worksheet.set_column('A:A', 35)
        worksheet.set_column('B:B', 15, percent_format)
        worksheet.set_column('C:C', 15, center_format)
        worksheet.set_column('D:D', 15, center_format)
        worksheet.set_column('E:E', 80, wrap_format)
        worksheet.set_column('F:F', 15, center_format)
        action_plan_columns = [
            translations[language]["category"], translations[language]["score"], translations[language]["priority"],
            translations[language]["estimated_effort"], translations[language]["action_plan"], translations[language]["type"]
        ]
        for col_num, value in enumerate(action_plan_columns):
            worksheet.write(23, col_num, value, header_format)
        for row in range(24, 24 + len(action_plan_df)):
            worksheet.write(row, 0, action_plan_df[translations[language]["category"]][row-24], 
                           cell_format if (row-24) % 2 else alt_row_format)
            if action_plan_df[translations[language]["score"]][row-24] != "":
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
        worksheet.freeze_panes(24, 0)

        # Performance Overview Section (Rows 50-68)
        logger.debug("Generating Performance Overview Section")
        worksheet.merge_range('A50:I50', translations[language]["actionable_charts"], title_format)
        chart_data = df_display[[translations[language]["percent"]]].reset_index()
        chart_data[translations[language]["score"]] = chart_data[translations[language]["percent"]] / 100
        chart_data.to_excel(writer, sheet_name="Informe de Auditoría" if language == "Español" else "Audit Report", 
                           startrow=51, startcol=0, index=False)
        worksheet.write('A52', translations[language]["category"], header_format)
        worksheet.write('B52', translations[language]["score_percent"], header_format)
        bar_chart = workbook.add_chart({'type': 'bar'})
        bar_chart.add_series({
            'name': translations[language]["score_percent"],
            'categories': f"='Informe de Auditoría'!$A$53:$A${52 + len(chart_data)}" if language == "Español" else 
                         f"='Audit Report'!$A$53:$A${52 + len(chart_data)}",
            'values': f"='Informe de Auditoría'!$B$53:$B${52 + len(chart_data)}" if language == "Español" else 
                     f"='Audit Report'!$B$53:$B${52 + len(chart_data)}",
            'fill': {'color': '#1E88E5'},
            'data_labels': {'value': True, 'num_format': '0.0%', 'font': {'size': 10, 'name': 'Arial'}}
        })
        bar_chart.set_title({'name': translations[language]["chart_title"], 'name_font': {'size': 12, 'name': 'Arial'}})
        bar_chart.set_x_axis({
            'name': translations[language]["score_percent"],
            'min': 0, 'max': 1, 'num_format': '0%',
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        bar_chart.set_y_axis({
            'name': translations[language]["category"],
            'reverse': True,
            'name_font': {'size': 10, 'name': 'Arial'},
            'num_font': {'size': 10, 'name': 'Arial'}
        })
        bar_chart.set_size({'width': 600, 'height': 200})
        worksheet.insert_chart('D53', bar_chart)

        # Summary Stats Table
        summary_stats = {
            translations[language]["metric"]: [
                "Puntuación Media" if language == "Español" else "Mean Score",
                "Varianza" if language == "Español" else "Variance",
                translations[language]["high_priority_categories"]
            ],
            translations[language]["value"]: [
                f"{df_display[translations[language]['percent']].mean():.1f}%",
                f"{df_display[translations[language]['percent']].var():.2f}",
                critical_count
            ]
        }
        summary_stats_df = pd.DataFrame(summary_stats)
        summary_stats_df.to_excel(writer, sheet_name="Informe de Auditoría" if language == "Español" else "Audit Report", 
                                startrow=51, startcol=2, index=False)
        worksheet.set_column('C:D', 20)
        worksheet.write('C52', translations[language]["metric"], header_format)
        worksheet.write('D52', translations[language]["value"], header_format)
        for row in range(52, 52 + len(summary_stats_df)):
            worksheet.write(row, 2, summary_stats_df[translations[language]["metric"]][row-52], 
                           cell_format if (row-52) % 2 else alt_row_format)

        # Why Choose LEAN 2.0 Institute? Section (Rows 70-78)
        logger.debug("Generating Why Choose Section")
        worksheet.merge_range('A70:I70', translations[language]["why_choose"], title_format)
        worksheet.merge_range('A71:I78', translations[language]["why_choose_message"], wrap_format)

        # Contact Section (Rows 80-88)
        logger.debug("Generating Contact Section")
        worksheet.merge_range('A80:I80', translations[language]["contact"], title_format)
        contact_df = pd.DataFrame({
            translations[language]["metric"]: ["Correo" if language == "Español" else "Email", "Sitio Web" if language == "Español" else "Website"],
            translations[language]["value"]: [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name="Informe de Auditoría" if language == "Español" else "Audit Report", 
                           startrow=81, startcol=0, index=False)
        worksheet.write('A82', translations[language]["metric"], header_format)
        worksheet.write('B82', translations[language]["value"], header_format)
        for row in range(82, 82 + len(contact_df)):
            worksheet.write(row, 0, contact_df[translations[language]["metric"]][row-82], 
                           cell_format if (row-82) % 2 else alt_row_format)
        worksheet.merge_range('A85:I88', translations[language]["marketing_message"], wrap_format)

    logger.debug("Excel report generation completed")
    excel_output.seek(0)
    return excel_output

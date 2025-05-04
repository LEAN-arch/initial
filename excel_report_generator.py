import pandas as pd
import io
import xlsxwriter
from typing import Dict, Any

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
    Generate a comprehensive Excel report with formatted sheets for summary, results, findings, 
    actionable insights, charts, and contact information.
    
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
    excel_output = io.BytesIO()
    with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Define formats
        bold = workbook.add_format({'bold': True, 'font_size': 12})
        percent_format = workbook.add_format({'num_format': '0.0%', 'font_size': 11})
        wrap_format = workbook.add_format({'text_wrap': True, 'font_size': 11, 'valign': 'top'})
        border_format = workbook.add_format({'border': 1, 'font_size': 11})
        header_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#1E88E5', 
            'color': 'white', 
            'border': 1, 
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter'
        })
        title_format = workbook.add_format({
            'bold': True, 
            'font_size': 14, 
            'align': 'center', 
            'valign': 'vcenter'
        })
        cell_format = workbook.add_format({
            'font_size': 11, 
            'border': 1, 
            'valign': 'top'
        })

        # Summary Sheet
        critical_count = len(df[df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
        improvement_count = len(df[
            (df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["CRITICAL"]) & 
            (df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"])
        ])
        summary_df = pd.DataFrame({
            TRANSLATIONS[language]["overall_score"]: [f"{overall_score:.1f}%"],
            TRANSLATIONS[language]["grade"]: [grade],
            TRANSLATIONS[language]["findings_summary"]: [
                TRANSLATIONS[language]["findings_summary_text"].format(
                    critical_count,
                    SCORE_THRESHOLDS["CRITICAL"],
                    improvement_count,
                    SCORE_THRESHOLDS["CRITICAL"],
                    SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1,
                    overall_score
                )
            ]
        })
        summary_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["summary"], index=False, startrow=3)
        worksheet_summary = writer.sheets[TRANSLATIONS[language]["summary"]]
        worksheet_summary.merge_range('A1:C1', TRANSLATIONS[language]["report_title"], title_format)
        worksheet_summary.write('A2', f"Date: {REPORT_DATE}", bold)
        worksheet_summary.set_column('A:A', 25, cell_format)
        worksheet_summary.set_column('B:B', 15, cell_format)
        worksheet_summary.set_column('C:C', 80, wrap_format)
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet_summary.write(3, col_num, value, header_format)

        # Results Sheet
        df_display.to_excel(writer, sheet_name=TRANSLATIONS[language]["results"], float_format="%.1f", startrow=3)
        worksheet_results = writer.sheets[TRANSLATIONS[language]["results"]]
        worksheet_results.merge_range('A1:D1', TRANSLATIONS[language]["results"], title_format)
        worksheet_results.write('A2', f"Date: {REPORT_DATE}", bold)
        worksheet_results.set_column('A:A', 35, cell_format)
        worksheet_results.set_column('B:C', 15, cell_format)
        worksheet_results.set_column('D:D', 20, cell_format)
        for col_num, value in enumerate(df_display.columns.values):
            worksheet_results.write(3, col_num + 1, value, header_format)
        worksheet_results.write(3, 0, TRANSLATIONS[language]["category"], header_format)

        # Findings Sheet
        findings_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            if df.loc[cat, TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                findings_data.append([
                    display_cat,
                    f"{df.loc[cat, TRANSLATIONS[language]['percent']]:.1f}%",
                    TRANSLATIONS[language]["high_priority"] if df.loc[cat, TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else TRANSLATIONS[language]["medium_priority"],
                    TRANSLATIONS[language]["action_required"].format(
                        "Urgent" if df.loc[cat, TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"] else "Specific"
                    )
                ])
                for idx, score in enumerate(responses[cat]):
                    if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        question, _, rec = questions[cat][language][idx]
                        findings_data.append([
                            "", "", "", 
                            f"{TRANSLATIONS[language]['question']}: {question[:50]}... - {TRANSLATIONS[language]['suggestion']}: {rec}"
                        ])
        findings_df = pd.DataFrame(
            findings_data,
            columns=[
                TRANSLATIONS[language]["category"],
                TRANSLATIONS[language]["score"],
                TRANSLATIONS[language]["priority"],
                TRANSLATIONS[language]["findings_and_suggestions"]
            ]
        )
        findings_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["findings"], index=False, startrow=3)
        worksheet_findings = writer.sheets[TRANSLATIONS[language]["findings"]]
        worksheet_findings.merge_range('A1:D1', TRANSLATIONS[language]["findings"], title_format)
        worksheet_findings.write('A2', f"Date: {REPORT_DATE}", bold)
        worksheet_findings.set_column('A:A', 35, cell_format)
        worksheet_findings.set_column('B:C', 15, cell_format)
        worksheet_findings.set_column('D:D', 80, wrap_format)
        for col_num, value in enumerate(findings_df.columns.values):
            worksheet_findings.write(3, col_num, value, header_format)

        # Actionable Insights Sheet
        insights_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            score = df.loc[cat, TRANSLATIONS[language]["percent"]]
            if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                insights_data.append([display_cat, f"{score:.1f}%", "Focus on immediate improvements."])
        insights_df = pd.DataFrame(
            insights_data,
            columns=[
                TRANSLATIONS[language]["category"],
                TRANSLATIONS[language]["score"],
                TRANSLATIONS[language]["actionable_insights"]
            ]
        )
        insights_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_insights"], index=False, startrow=3)
        worksheet_insights = writer.sheets[TRANSLATIONS[language]["actionable_insights"]]
        worksheet_insights.merge_range('A1:C1', TRANSLATIONS[language]["actionable_insights"], title_format)
        worksheet_insights.write('A2', f"Date: {REPORT_DATE}", bold)
        worksheet_insights.set_column('A:A', 35, cell_format)
        worksheet_insights.set_column('B:B', 15, cell_format)
        worksheet_insights.set_column('C:C', 50, wrap_format)
        for col_num, value in enumerate(insights_df.columns.values):
            worksheet_insights.write(3, col_num, value, header_format)

        # Actionable Charts Sheet
        worksheet_charts = workbook.add_worksheet(TRANSLATIONS[language]["actionable_charts"])
        worksheet_charts.merge_range('A1:B1', TRANSLATIONS[language]["actionable_charts"], title_format)
        worksheet_charts.write('A2', f"Date: {REPORT_DATE}", bold)
        chart_data = df_display[[TRANSLATIONS[language]["percent"]]].reset_index()
        chart_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, index=False)
        worksheet_charts.set_column('A:A', 35, cell_format)
        worksheet_charts.set_column('B:B', 15, cell_format)
        worksheet_charts.write('A5', TRANSLATIONS[language]["category"], header_format)
        worksheet_charts.write('B5', TRANSLATIONS[language]["score_percent"], header_format)
        bar_chart = workbook.add_chart({'type': 'bar'})
        bar_chart.add_series({
            'name': TRANSLATIONS[language]["score_percent"],
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$A$6:$A${5 + len(chart_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$B$6:$B${5 + len(chart_data)}",
            'fill': {'color': '#1E88E5'},
            'data_labels': {'value': True, 'position': 'inside_end'}
        })
        bar_chart.set_title({'name': TRANSLATIONS[language]["chart_title"]})
        bar_chart.set_x_axis({'name': TRANSLATIONS[language]["score_percent"], 'min': 0, 'max': 100})
        bar_chart.set_y_axis({'name': TRANSLATIONS[language]["category"], 'reverse': True})
        bar_chart.set_size({'width': 720, 'height': 400})
        worksheet_charts.insert_chart('D5', bar_chart)

        # Contact Sheet
        contact_df = pd.DataFrame({
            "Contact Method": ["Email", "Website"],
            "Details": [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["contact"], index=False, startrow=3)
        worksheet_contact = writer.sheets[TRANSLATIONS[language]["contact"]]
        worksheet_contact.merge_range('A1:B1', TRANSLATIONS[language]["contact"], title_format)
        worksheet_contact.write('A2', f"Date: {REPORT_DATE}", bold)
        worksheet_contact.set_column('A:A', 20, cell_format)
        worksheet_contact.set_column('B:B', 50, cell_format)
        for col_num, value in enumerate(contact_df.columns.values):
            worksheet_contact.write(3, col_num, value, header_format)
        worksheet_contact.write('A7', "Collaborate with Us", bold)
        worksheet_contact.write('A8', TRANSLATIONS[language]["marketing_message"], wrap_format)

    excel_output.seek(0)
    return excel_output

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
    Generate a professional Excel report with enhanced formatting, executive summary, 
    actionable insights, branded cover sheet, and improved charts for maximum impact.
    
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
        title_format = workbook.add_format({
            'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        subtitle_format = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'left', 'valign': 'vcenter',
            'font_color': '#424242'
        })
        header_format = workbook.add_format({
            'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter',
            'bg_color': '#1E88E5', 'font_color': 'white', 'border': 1
        })
        cell_format = workbook.add_format({
            'font_size': 10, 'border': 1, 'valign': 'top', 'text_wrap': True
        })
        percent_format = workbook.add_format({
            'num_format': '0.0%', 'font_size': 10, 'border': 1, 'valign': 'top'
        })
        bold_format = workbook.add_format({
            'bold': True, 'font_size': 10, 'border': 1, 'valign': 'top'
        })
        wrap_format = workbook.add_format({
            'text_wrap': True, 'font_size': 10, 'border': 1, 'valign': 'top'
        })
        center_format = workbook.add_format({
            'align': 'center', 'font_size': 10, 'border': 1, 'valign': 'top'
        })
        # Conditional formats for priority
        critical_format = workbook.add_format({
            'bg_color': '#D32F2F', 'font_color': 'white', 'border': 1
        })
        improvement_format = workbook.add_format({
            'bg_color': '#FFD54F', 'font_color': '#212121', 'border': 1
        })
        good_format = workbook.add_format({
            'bg_color': '#43A047', 'font_color': 'white', 'border': 1
        })

        # Cover Sheet
        cover_sheet = workbook.add_worksheet("Cover")
        cover_sheet.merge_range('A1:E1', TRANSLATIONS[language]["report_title"], title_format)
        cover_sheet.merge_range('A2:E2', f"Generated on: {REPORT_DATE}", subtitle_format)
        cover_sheet.merge_range('A3:E3', "Prepared by: LEAN 2.0 Institute", subtitle_format)
        cover_sheet.merge_range('A4:E4', TRANSLATIONS[language]["marketing_message"], wrap_format)
        cover_sheet.write('A6', "Navigate to:", bold_format)
        cover_sheet.write('A7', 'Executive Summary', workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10
        }))
        cover_sheet.write('A8', 'Results', workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10
        }))
        cover_sheet.write('A9', 'Action Plan', workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10
        }))
        cover_sheet.write('A10', 'Charts', workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10
        }))
        cover_sheet.write('A11', 'Contact', workbook.add_format({
            'font_color': 'blue', 'underline': 1, 'font_size': 10
        }))
        # Placeholder for logo (assuming logo.png is available)
        cover_sheet.write('A13', "Logo Placeholder: Insert LEAN 2.0 Institute logo here", bold_format)
        cover_sheet.set_column('A:E', 25)

        # Executive Summary Sheet
        critical_count = len(df[df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]])
        improvement_count = len(df[
            (df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["CRITICAL"]) & 
            (df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"])
        ])
        summary_data = {
            TRANSLATIONS[language]["overall_score"]: [f"{overall_score:.1f}%"],
            TRANSLATIONS[language]["grade"]: [grade],
            TRANSLATIONS[language]["high_priority_categories"]: [critical_count],
            "Categories Needing Improvement": [improvement_count],
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
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["summary"], index=False, startrow=3)
        worksheet_summary = writer.sheets[TRANSLATIONS[language]["summary"]]
        worksheet_summary.merge_range('A1:E1', TRANSLATIONS[language]["summary"], title_format)
        worksheet_summary.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_summary.set_column('A:A', 25, cell_format)
        worksheet_summary.set_column('B:E', 20, cell_format)
        worksheet_summary.set_column('E:E', 60, wrap_format)
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet_summary.write(3, col_num, value, header_format)
        # Add conditional formatting for scores
        worksheet_summary.conditional_format('A5:A5', {
            'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"],
            'format': critical_format
        })
        worksheet_summary.conditional_format('A5:A5', {
            'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"],
            'maximum': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1, 'format': improvement_format
        })
        worksheet_summary.conditional_format('A5:A5', {
            'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"],
            'format': good_format
        })

        # Results Sheet
        results_df = df_display.copy()
        results_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["results"], startrow=3)
        worksheet_results = writer.sheets[TRANSLATIONS[language]["results"]]
        worksheet_results.merge_range('A1:D1', TRANSLATIONS[language]["results"], title_format)
        worksheet_results.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_results.set_column('A:A', 35, cell_format)
        worksheet_results.set_column('B:D', 15, center_format)
        for col_num, value in enumerate(results_df.columns.values):
            worksheet_results.write(3, col_num + 1, value, header_format)
        worksheet_results.write(3, 0, TRANSLATIONS[language]["category"], header_format)
        # Conditional formatting for priority
        for row in range(4, 4 + len(results_df)):
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': '<', 'value': SCORE_THRESHOLDS["CRITICAL"],
                'format': critical_format
            })
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': 'between', 'minimum': SCORE_THRESHOLDS["CRITICAL"],
                'maximum': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]-1, 'format': improvement_format
            })
            worksheet_results.conditional_format(f'C{row}:C{row}', {
                'type': 'cell', 'criteria': '>=', 'value': SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"],
                'format': good_format
            })

        # Action Plan Sheet
        action_plan_data = []
        for cat in questions.keys():
            display_cat = next(k for k, v in category_mapping[language].items() if v == cat)
            score = df.loc[cat, TRANSLATIONS[language]["percent"]]
            priority = (
                TRANSLATIONS[language]["high_priority"] if score < SCORE_THRESHOLDS["CRITICAL"] else
                TRANSLATIONS[language]["medium_priority"] if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"] else
                TRANSLATIONS[language]["low_priority"]
            )
            if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                action_plan_data.append([
                    display_cat, f"{score:.1f}%", priority, 
                    f"Immediate action required to address low performance in {display_cat}."
                ])
                for idx, score in enumerate(responses[cat]):
                    if score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        question, _, rec = questions[cat][language][idx]
                        action_plan_data.append([
                            "", "", "",
                            f"{TRANSLATIONS[language]['question']}: {question[:100]}{'...' if len(question) > 100 else ''}\n"
                            f"{TRANSLATIONS[language]['suggestion']}: {rec}"
                        ])
        action_plan_df = pd.DataFrame(
            action_plan_data,
            columns=[
                TRANSLATIONS[language]["category"],
                TRANSLATIONS[language]["score"],
                TRANSLATIONS[language]["priority"],
                "Action Plan"
            ]
        )
        action_plan_df.to_excel(writer, sheet_name="Action Plan", index=False, startrow=3)
        worksheet_action = writer.sheets["Action Plan"]
        worksheet_action.merge_range('A1:D1', "Action Plan", title_format)
        worksheet_action.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_action.set_column('A:A', 35, cell_format)
        worksheet_action.set_column('B:C', 15, center_format)
        worksheet_action.set_column('D:D', 80, wrap_format)
        for col_num, value in enumerate(action_plan_df.columns.values):
            worksheet_action.write(3, col_num, value, header_format)
        # Conditional formatting
        for row in range(4, 4 + len(action_plan_df)):
            worksheet_action.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 
                'value': TRANSLATIONS[language]["high_priority"], 'format': critical_format
            })
            worksheet_action.conditional_format(f'C{row}:C{row}', {
                'type': 'text', 'criteria': 'containing', 
                'value': TRANSLATIONS[language]["medium_priority"], 'format': improvement_format
            })

        # Charts Sheet
        worksheet_charts = workbook.add_worksheet(TRANSLATIONS[language]["actionable_charts"])
        worksheet_charts.merge_range('A1:D1', TRANSLATIONS[language]["actionable_charts"], title_format)
        worksheet_charts.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        chart_data = df_display[[TRANSLATIONS[language]["percent"]]].reset_index()
        chart_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, index=False)
        worksheet_charts.set_column('A:A', 35, cell_format)
        worksheet_charts.set_column('B:B', 15, percent_format)
        worksheet_charts.write('A5', TRANSLATIONS[language]["category"], header_format)
        worksheet_charts.write('B5', TRANSLATIONS[language]["score_percent"], header_format)

        # Bar Chart
        bar_chart = workbook.add_chart({'type': 'bar'})
        bar_chart.add_series({
            'name': TRANSLATIONS[language]["score_percent"],
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$A$6:$A${5 + len(chart_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$B$6:$B${5 + len(chart_data)}",
            'fill': {'color': '#1E88E5'},
            'data_labels': {'value': True, 'position': 'inside_end', 'num_format': '0.0%'}
        })
        bar_chart.set_title({'name': TRANSLATIONS[language]["chart_title"]})
        bar_chart.set_x_axis({'name': TRANSLATIONS[language]["score_percent"], 'min': 0, 'max': 1, 'num_format': '0%'})
        bar_chart.set_y_axis({'name': TRANSLATIONS[language]["category"], 'reverse': True})
        bar_chart.set_size({'width': 720, 'height': 300})
        worksheet_charts.insert_chart('D5', bar_chart)

        # Pie Chart for Priority Distribution
        priority_counts = [
            len(df[df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["CRITICAL"]]),
            len(df[
                (df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["CRITICAL"]) & 
                (df[TRANSLATIONS[language]["percent"]] < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"])
            ]),
            len(df[df[TRANSLATIONS[language]["percent"]] >= SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]])
        ]
        priority_df = pd.DataFrame({
            TRANSLATIONS[language]["priority"]: [
                TRANSLATIONS[language]["high_priority"],
                TRANSLATIONS[language]["medium_priority"],
                TRANSLATIONS[language]["low_priority"]
            ],
            "Count": priority_counts
        })
        priority_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=3, index=False)
        worksheet_charts.write('D5', TRANSLATIONS[language]["priority"], header_format)
        worksheet_charts.write('E5', "Count", header_format)
        pie_chart = workbook.add_chart({'type': 'pie'})
        pie_chart.add_series({
            'name': 'Priority Distribution',
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$D$6:$D$8",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$E$6:$E$8",
            'data_labels': {'percentage': True, 'category': True},
            'points': [
                {'fill': {'color': '#D32F2F'}},
                {'fill': {'color': '#FFD54F'}},
                {'fill': {'color': '#43A047'}}
            ]
        })
        pie_chart.set_title({'name': "Priority Distribution"})
        pie_chart.set_size({'width': 360, 'height': 300})
        worksheet_charts.insert_chart('D20', pie_chart)

        # Contact Sheet
        contact_df = pd.DataFrame({
            "Contact Method": ["Email", "Website"],
            "Details": [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["contact"], index=False, startrow=3)
        worksheet_contact = writer.sheets[TRANSLATIONS[language]["contact"]]
        worksheet_contact.merge_range('A1:B1', TRANSLATIONS[language]["contact"], title_format)
        worksheet_contact.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_contact.set_column('A:A', 20, cell_format)
        worksheet_contact.set_column('B:B', 50, cell_format)
        for col_num, value in enumerate(contact_df.columns.values):
            worksheet_contact.write(3, col_num, value, header_format)
        worksheet_contact.write('A7', "Collaborate with Us", bold_format)
        worksheet_contact.write('A8', TRANSLATIONS[language]["marketing_message"], wrap_format)

        # Add hyperlinks to Cover Sheet
        cover_sheet.write_url('A7', f"internal:'{TRANSLATIONS[language]['summary']}'!A1", 
                            string='Executive Summary')
        cover_sheet.write_url('A8', f"internal:'{TRANSLATIONS[language]['results']}'!A1", 
                            string='Results')
        cover_sheet.write_url('A9', f"internal:'Action Plan'!A1", string='Action Plan')
        cover_sheet.write_url('A10', f"internal:'{TRANSLATIONS[language]['actionable_charts']}'!A1", 
                            string='Charts')
        cover_sheet.write_url('A11', f"internal:'{TRANSLATIONS[language]['contact']}'!A1", 
                            string='Contact')

    excel_output.seek(0)
    return excel_output

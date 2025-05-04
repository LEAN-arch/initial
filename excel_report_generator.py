import pandas as pd
import io
import xlsxwriter
from typing import Dict, Any
import numpy as np

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
    Generate a professional Excel report with enhanced content, contact information on the cover sheet,
    refined visualizations, and actionable insights for maximum impact.
    
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

        # Data Validation
        if df.empty or df_display.empty or not responses:
            raise ValueError("Input data is empty or invalid")
        if not all(col in df.columns for col in [TRANSLATIONS[language]["score"], TRANSLATIONS[language]["percent"], TRANSLATIONS[language]["priority"]]):
            raise ValueError("Required columns missing in df")
        for cat in responses:
            if not responses[cat] or any(score is None for score in responses[cat]):
                raise ValueError(f"Invalid or missing responses in category {cat}")

        # Cover Sheet
        cover_sheet = workbook.add_worksheet("Cover")
        cover_sheet.merge_range('A1:F1', TRANSLATIONS[language]["report_title"], title_format)
        cover_sheet.merge_range('A2:F2', "Ethical Lean Workplace Audit", subtitle_format)
        cover_sheet.merge_range('A3:F3', f"Generated on: {REPORT_DATE}", subtitle_format)
        cover_sheet.merge_range('A4:F4', "Prepared by: LEAN 2.0 Institute", subtitle_format)
        cover_sheet.merge_range('A5:F5', "Contact Us:", bold_format)
        cover_sheet.merge_range('A6:F6', f"Email: {CONFIG['contact']['email']} | Website: {CONFIG['contact']['website']}", cell_format)
        cover_sheet.merge_range('A7:F7', "Mission: Transforming workplaces through ethical, inclusive, and sustainable practices.", wrap_format)
        cover_sheet.write('A9', "Table of Contents:", bold_format)
        toc_links = [
            ("Executive Summary", TRANSLATIONS[language]["summary"]),
            ("Statistical Insights", "Statistical Insights"),
            ("Results", TRANSLATIONS[language]["results"]),
            ("Action Plan", "Action Plan"),
            ("Visualizations", TRANSLATIONS[language]["actionable_charts"]),
            ("Contact", TRANSLATIONS[language]["contact"])
        ]
        for idx, (name, sheet) in enumerate(toc_links, start=10):
            cover_sheet.write_url(f'A{idx}', f"internal:'{sheet}'!A1", 
                               string=name, cell_format=workbook.add_format({
                                   'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'
                               }))
        cover_sheet.write('A17', "Logo Placeholder: Insert LEAN 2.0 Institute logo", bold_format)
        cover_sheet.set_column('A:F', 25)

        # Executive Summary Sheet
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
            "Categories Needing Improvement": [improvement_count],
            "Categories Performing Well": [good_count],
            "Next Steps": ["Review the Action Plan for prioritized recommendations."]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["summary"], index=False, startrow=4)
        worksheet_summary = writer.sheets[TRANSLATIONS[language]["summary"]]
        worksheet_summary.merge_range('A1:F1', TRANSLATIONS[language]["summary"], title_format)
        worksheet_summary.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_summary.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
                                  cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_summary.set_column('A:A', 25, cell_format)
        worksheet_summary.set_column('B:E', 20, center_format)
        worksheet_summary.set_column('F:F', 60, wrap_format)
        for col_num, value in enumerate(summary_df.columns.values):
            worksheet_summary.write(4, col_num, value, header_format)
        worksheet_summary.write('A4', "Metric", header_format)
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
        stats_data = {
            TRANSLATIONS[language]["category"]: df_display.index,
            TRANSLATIONS[language]["score"]: df_display[TRANSLATIONS[language]["percent"]],
            "Standard Deviation": [np.std(responses[cat]) for cat in responses],
            "Min Score": [min(responses[cat]) for cat in responses],
            "Max Score": [max(responses[cat]) for cat in responses]
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name="Statistical Insights", index=False, startrow=3)
        worksheet_stats = writer.sheets["Statistical Insights"]
        worksheet_stats.merge_range('A1:E1', "Statistical Insights", title_format)
        worksheet_stats.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_stats.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
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
        results_df = df_display.copy()
        results_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["results"], startrow=3)
        worksheet_results = writer.sheets[TRANSLATIONS[language]["results"]]
        worksheet_results.merge_range('A1:D1', TRANSLATIONS[language]["results"], title_format)
        worksheet_results.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_results.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
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

        # Action Plan Sheet with Enhanced Prioritization
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
                impact_score = (100 - score) / 100
                feasibility_score = score / 100
                priority_score = impact_score * 0.7 + feasibility_score * 0.3
                effort = (
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
                    "Action Plan": f"Immediate action required to address low performance in {display_cat}.",
                    "Type": "Category"
                })
                for idx, q_score in enumerate(responses[cat]):
                    if q_score < SCORE_THRESHOLDS["NEEDS_IMPROVEMENT"]:
                        question, _, rec = questions[cat][language][idx]
                        q_impact_score = (100 - q_score) / 100
                        q_priority_score = q_impact_score * 0.7 + feasibility_score * 0.3
                        q_effort = (
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
                            "Type": "Question"
                        })
        action_plan_df = pd.DataFrame(action_plan_data).sort_values(["Priority", "Priority Score"], ascending=[True, False])
        action_plan_df.to_excel(writer, sheet_name="Action Plan", index=False, startrow=3)
        worksheet_action = writer.sheets["Action Plan"]
        worksheet_action.merge_range('A1:F1', "Action Plan", title_format)
        worksheet_action.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_action.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
                                cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_action.set_column('A:A', 35, cell_format)
        worksheet_action.set_column('B:C', 15, center_format)
        worksheet_action.set_column('D:D', 15, number_format)
        worksheet_action.set_column('E:E', 15, center_format)
        worksheet_action.set_column('F:F', 80, wrap_format)
        action_plan_columns = [
            TRANSLATIONS[language]["category"], TRANSLATIONS[language]["score"], TRANSLATIONS[language]["priority"],
            "Priority Score", "Estimated Effort", "Action Plan"
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
        worksheet_viz = workbook.add_worksheet(TRANSLATIONS[language]["actionable_charts"])
        worksheet_viz.merge_range('A1:F1', TRANSLATIONS[language]["actionable_charts"], title_format)
        worksheet_viz.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_viz.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
                              cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))

        # Bar Chart
        chart_data = df_display[[TRANSLATIONS[language]["percent"]]].reset_index()
        chart_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, index=False)
        worksheet_viz.set_column('A:A', 35, cell_format)
        worksheet_viz.set_column('B:B', 15, percent_format)
        worksheet_viz.write('A5', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('B5', TRANSLATIONS[language]["score_percent"], header_format)
        bar_chart = workbook.add_chart({'type': 'bar'})
        bar_chart.add_series({
            'name': TRANSLATIONS[language]["score_percent"],
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$A$6:$A${5 + len(chart_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$B$6:$B${5 + len(chart_data)}",
            'fill': {'color': '#1E88E5'},
            'data_labels': {'value': True, 'num_format': '0.0%'}
        })
        bar_chart.set_title({'name': TRANSLATIONS[language]["chart_title"]})
        bar_chart.set_x_axis({'name': TRANSLATIONS[language]["score_percent"], 'min': 0, 'max': 1, 'num_format': '0%'})
        bar_chart.set_y_axis({'name': TRANSLATIONS[language]["category"], 'reverse': True})
        bar_chart.set_size({'width': 600, 'height': 300})
        worksheet_viz.insert_chart('E5', bar_chart)

        # Line Chart
        line_data = df_display[[TRANSLATIONS[language]["percent"]]].reset_index()
        line_data.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=3, index=False)
        worksheet_viz.write('D5', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('E5', TRANSLATIONS[language]["score_percent"], header_format)
        line_chart = workbook.add_chart({'type': 'line'})
        line_chart.add_series({
            'name': TRANSLATIONS[language]["score_percent"],
            'categories': f"='{TRANSLATIONS[language]['actionable_charts']}'!$D$6:$D${5 + len(line_data)}",
            'values': f"='{TRANSLATIONS[language]['actionable_charts']}'!$E$6:$E${5 + len(line_data)}",
            'line': {'color': '#1E88E5', 'width': 2},
            'marker': {'type': 'circle', 'size': 6, 'fill': {'color': '#1E88E5'}, 'border': {'color': '#1E88E5'}},
            'data_labels': {'value': True, 'num_format': '0.0%'}
        })
        line_chart.set_title({'name': "Score Trend Across Categories"})
        line_chart.set_x_axis({'name': TRANSLATIONS[language]["category"]})
        line_chart.set_y_axis({'name': TRANSLATIONS[language]["score_percent"], 'min': 0, 'max': 1, 'num_format': '0%'})
        line_chart.set_size({'width': 600, 'height': 300})
        worksheet_viz.insert_chart('E20', line_chart)

        # Heatmap
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
        heatmap_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["actionable_charts"], startrow=4, startcol=7, index=False)
        worksheet_viz.set_column('H:J', 25, cell_format)
        worksheet_viz.write('H5', TRANSLATIONS[language]["category"], header_format)
        worksheet_viz.write('I5', TRANSLATIONS[language]["question"], header_format)
        worksheet_viz.write('J5', TRANSLATIONS[language]["score_percent"], header_format)
        for row in range(5, 5 + len(heatmap_df)):
            worksheet_viz.write(row, 7, heatmap_df[TRANSLATIONS[language]["category"]][row-5], cell_format if (row-5) % 2 else alt_row_format)
            worksheet_viz.conditional_format(f'J{row}:J{row}', {
                'type': '3_color_scale',
                'min_color': '#D32F2F',
                'mid_color': '#FFD54F',
                'max_color': '#43A047'
            })

        # Contact Sheet
        contact_df = pd.DataFrame({
            "Contact Method": ["Email", "Website"],
            "Details": [CONFIG["contact"]["email"], CONFIG["contact"]["website"]]
        })
        contact_df.to_excel(writer, sheet_name=TRANSLATIONS[language]["contact"], index=False, startrow=3)
        worksheet_contact = writer.sheets[TRANSLATIONS[language]["contact"]]
        worksheet_contact.merge_range('A1:B1', TRANSLATIONS[language]["contact"], title_format)
        worksheet_contact.write('A2', f"Date: {REPORT_DATE}", subtitle_format)
        worksheet_contact.write_url('A3', "internal:Cover!A1", string="Back to Cover", 
                                 cell_format=workbook.add_format({'font_color': 'blue', 'underline': 1, 'font_size': 10, 'font_name': 'Arial'}))
        worksheet_contact.set_column('A:A', 20, cell_format)
        worksheet_contact.set_column('B:B', 50, cell_format)
        for col_num, value in enumerate(contact_df.columns.values):
            worksheet_contact.write(3, col_num, value, header_format)
        for row in range(4, 4 + len(contact_df)):
            worksheet_contact.write(row, 0, contact_df["Contact Method"][row-4], cell_format if (row-4) % 2 else alt_row_format)
        worksheet_contact.write('A7',"¡Trabajemos juntos! | Let's work together!", bold_format)
        worksheet_contact.write('A8', TRANSLATIONS[language]["marketing_message"], wrap_format)

    excel_output.seek(0)
    return excel_output

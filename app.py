import json
import streamlit.components.v1 as components
import streamlit as st
import pandas as pd
import os
import grade_logic
from datetime import datetime

# Page Config
st.set_page_config(page_title="Student Result System", page_icon="🎓", layout="wide")

# Session State Initialization
if 'all_sheets_data' not in st.session_state:
    st.session_state.all_sheets_data = {}
if 'current_file' not in st.session_state:
    st.session_state.current_file = None
if 'pending_changes' not in st.session_state:
    st.session_state.pending_changes = {}
if 'pending_deletes' not in st.session_state:
    st.session_state.pending_deletes = {}
if 'theme' not in st.session_state:
    st.session_state.theme = "Light"

# Theme Definitions
themes = {
    "Light": {
        "main_header": "#1E88E5",
        "sub_header": "#424242",
        "card_bg": "#f9f9f9",
        "card_text": "#333333",
        "metric_bg": "#e3f2fd",
        "metric_text": "#333333",
        "metric_value": "#1565c0",
        "metric_label": "#555",
        "body_bg": "#ffffff",
        "body_text": "#333333"
    },
    "Dark": {
        "main_header": "#90caf9",
        "sub_header": "#e0e0e0",
        "card_bg": "#424242",
        "card_text": "#ffffff",
        "metric_bg": "#616161",
        "metric_text": "#ffffff",
        "metric_value": "#90caf9",
        "metric_label": "#bdbdbd",
        "body_bg": "#303030",
        "body_text": "#e0e0e0"
    }
}

current_theme = themes[st.session_state.theme]

# Custom CSS
st.markdown(f"""
<style>
    /* Main App Background */
    .stApp {{
        background-color: {current_theme['body_bg']};
        color: {current_theme['body_text']};
    }}
    
    /* Sidebar Styling */
    [data-testid="stSidebar"] {{
        background-color: {current_theme['card_bg']};
    }}
    [data-testid="stSidebar"] .stMarkdown {{
        color: {current_theme['card_text']};
    }}
    
    /* Input Fields */
    .stSelectbox label, .stTextInput label {{
        color: {current_theme['body_text']} !important;
    }}
    .stSelectbox div[data-baseweb="select"] > div {{
        background-color: {current_theme['card_bg']};
        color: {current_theme['card_text']};
    }}
    
    /* Buttons */
    .stButton > button {{
        background-color: {current_theme['metric_bg']};
        color: {current_theme['metric_text']};
        border: 1px solid {current_theme['metric_value']};
    }}
    .stButton > button:hover {{
        background-color: {current_theme['metric_value']};
        color: white;
    }}
    
    /* Dataframe/Data Editor */
    [data-testid="stDataFrame"], [data-testid="stDataEditor"] {{
        background-color: {current_theme['card_bg']};
        color: {current_theme['card_text']};
    }}
    
    /* Info/Warning/Success Messages */
    .stAlert {{
        background-color: {current_theme['card_bg']};
        color: {current_theme['card_text']};
    }}
    
    /* Custom Classes */
    .main-header {{
        font-size: 2.5rem;
        font-weight: bold;
        color: {current_theme['main_header']};
        margin-bottom: 1rem;
    }}
    .sub-header {{
        font-size: 1.5rem;
        font-weight: bold;
        color: {current_theme['sub_header']};
        margin-top: 1rem;
    }}
    .card {{
        background-color: {current_theme['card_bg']};
        padding: 0.8rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin-bottom: 0.5rem;
        color: {current_theme['card_text']};
    }}
    .metric-box {{
        text-align: center;
        padding: 8px;
        background-color: {current_theme['metric_bg']};
        border-radius: 8px;
        border: 1px solid #90caf9;
        color: {current_theme['metric_text']};
    }}
    .metric-value {{
        font-size: 1.5rem;
        font-weight: bold;
        color: {current_theme['metric_value']};
    }}
    .metric-label {{
        font-size: 0.9rem;
        color: {current_theme['metric_label']};
    }}
    .block-container {{
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }}
</style>
""", unsafe_allow_html=True)

# Main Header
st.markdown('<div class="main-header">🎓 Student Result System</div>', unsafe_allow_html=True)

# --- Sidebar ---
st.sidebar.title("Control Panel")

# File Selection
sheets_dir = "sheets"
if not os.path.exists(sheets_dir):
    os.makedirs(sheets_dir)

files = [f for f in os.listdir(sheets_dir) if f.endswith(('.xlsx', '.xls'))]
selected_file = st.sidebar.selectbox("Select Intake", ["Select a file..."] + files)

if selected_file != "Select a file..." and selected_file != st.session_state.current_file:
    file_path = os.path.join(sheets_dir, selected_file)
    with st.spinner(f"Loading {selected_file}..."):
        data, cols, credits = grade_logic.load_workbook_data(file_path)
        if data:
            st.session_state.all_sheets_data = data
            st.session_state.subject_columns_per_sheet = cols
            st.session_state.subject_credits_per_sheet = credits
            st.session_state.current_file = selected_file
            st.session_state.pending_changes = {}
            st.session_state.pending_deletes = {}
            st.success(f"Loaded {selected_file}")
        else:
            st.error("Failed to load workbook.")

if st.session_state.current_file:
    sheet_names = list(st.session_state.all_sheets_data.keys())
    selected_sheet = st.sidebar.selectbox("Select Batch", sheet_names)
    
    include_gpa = st.sidebar.checkbox("Include GPA & Class in Transcript", value=True)

    if selected_sheet:
        df = st.session_state.all_sheets_data[selected_sheet]

        # Search via Selectbox
        # Identify Name Column
        name_col = next((c for c in df.columns if "name" in c.lower()), None)
        
        # Create Search Options: "RegNo - Name"
        if name_col:
            # Create a dictionary for mapping: "RegNo - Name" -> RegNo
            # We use a list comprehension to handle potential missing values gracefully
            search_map = {}
            for idx, row in df.iterrows():
                reg = str(row['Registration Number'])
                name = str(row[name_col]) if pd.notna(row[name_col]) else "Unknown"
                display_str = f"{reg} - {name}"
                search_map[display_str] = reg
            
            search_options = ["Select a student..."] + sorted(list(search_map.keys()))
        else:
            # Fallback if no name column found
            all_regs = df['Registration Number'].dropna().astype(str).unique().tolist()
            search_options = ["Select a student..."] + sorted(all_regs)
            search_map = {r: r for r in all_regs}

        selected_option = st.selectbox("Search Student", search_options, index=0)
        
        student_row = None
        student_idx = None
        
        if selected_option != "Select a student...":
            # Extract RegNo from selection
            if name_col:
                selected_reg = search_map[selected_option]
            else:
                selected_reg = selected_option
                
            matches = df[df['Registration Number'].astype(str) == selected_reg]
            
            if len(matches) == 0:
                st.warning("Student not found.")
            elif len(matches) > 1:
                st.info(f"Multiple entries found for {selected_reg}. Select one below.")
                cols_to_show = ['Registration Number']
                if len(matches.columns) > 1:
                    second_col = matches.columns[1]
                    if second_col != 'Registration Number':
                        cols_to_show.append(second_col)
                
                event = st.dataframe(
                    matches[cols_to_show],
                    on_select="rerun",
                    selection_mode="single-row",
                    width="stretch",
                    hide_index=True
                )
                
                if len(event.selection.rows) > 0:
                    selected_row_idx = event.selection.rows[0]
                    student_idx = matches.index[selected_row_idx]
                    student_row = matches.loc[student_idx]
            else:
                student_idx = matches.index[0]
                student_row = matches.iloc[0]

        if student_row is not None:
            # Student Info
            name_col = next((c for c in df.columns if "name" in c.lower()), None)
            name = student_row[name_col] if name_col else "Unknown"
            reg_no = student_row['Registration Number']
            
            # Prepare Data for Editor
            valid_subjects = st.session_state.subject_columns_per_sheet.get(selected_sheet, [])
            subject_credits = st.session_state.subject_credits_per_sheet.get(selected_sheet, {})
            
            editor_data = []
            for sub in valid_subjects:
                grade = student_row[sub]
                # Ensure grade is a scalar value (not a Series)
                if isinstance(grade, pd.Series):
                    grade = grade.iloc[0] if len(grade) > 0 else ""
                grade_display = grade if pd.notna(grade) and str(grade).strip() != "" else ""
                
                # Check for pending changes
                if selected_sheet in st.session_state.pending_changes:
                    if reg_no in st.session_state.pending_changes[selected_sheet]:
                        if sub in st.session_state.pending_changes[selected_sheet][reg_no]:
                            grade_display = st.session_state.pending_changes[selected_sheet][reg_no][sub]

                editor_data.append({"Subject": sub, "Grade": grade_display})
                
            editor_df = pd.DataFrame(editor_data)
            editor_df.index = editor_df.index + 1
            
            # Calculate GPA (Initial)
            current_grades = [row['Grade'] for row in editor_data]
            current_credits = [subject_credits.get(row['Subject'], None) for row in editor_data]
            gpa = grade_logic.calculate_gpa(current_grades, current_credits)
            class_awarded = grade_logic.calculate_class(gpa)

            # Layout: Info & Metrics Side-by-Side
            col_info, col_metrics = st.columns([2, 1])
            
            with col_info:
                st.markdown(f"""
                <div class="card">
                    <div style="font-size: 1.2rem; font-weight: bold; margin-bottom: 5px;">{name}</div>
                    <p style="margin: 0; font-size: 0.9rem;"><b>Reg:</b> {reg_no} | <b>Sheet:</b> {selected_sheet}</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_metrics:
                st.markdown(f"""
                <div class="metric-box">
                    <div class="metric-label">GPA</div>
                    <div class="metric-value">{gpa:.2f}</div>
                    <div class="metric-label" style="margin-top:5px;">{class_awarded}</div>
                </div>
                """, unsafe_allow_html=True)
            
            st.markdown('<div class="sub-header">Results</div>', unsafe_allow_html=True)
            
            edited_df = st.data_editor(
                editor_df, 
                column_config={
                    "Subject": st.column_config.TextColumn("Subject", disabled=True),
                    "Grade": st.column_config.TextColumn("Grade")
                },
                width="stretch",
                key="grade_editor"
            )
            
            # Process Edits & Recalculate GPA if needed
            changes_detected = False
            recalc_needed = False
            
            for index, row in edited_df.iterrows():
                sub = row['Subject']
                new_grade = row['Grade']
                original_grade = next((item['Grade'] for item in editor_data if item['Subject'] == sub), "")
                
                if new_grade != original_grade:
                    changes_detected = True
                    recalc_needed = True
                    # Update Pending Changes
                    if selected_sheet not in st.session_state.pending_changes:
                        st.session_state.pending_changes[selected_sheet] = {}
                    if reg_no not in st.session_state.pending_changes[selected_sheet]:
                        st.session_state.pending_changes[selected_sheet][reg_no] = {}
                    
                    st.session_state.pending_changes[selected_sheet][reg_no][sub] = new_grade
            
            if recalc_needed:
                # Recalculate GPA with new values for display update (requires rerun or dynamic update)
                # Since we can't easily update the metrics above without a rerun, we'll rely on the next rerun.
                # However, st.data_editor triggers a rerun on edit, so the values above will update automatically!
                pass

            if changes_detected:
                st.info("Changes detected. Click 'Save Changes to Excel' in the sidebar to commit.")
            
            # Actions in Sidebar
            st.sidebar.markdown("---")
            st.sidebar.markdown("### Actions")
            
            if st.sidebar.button("Print Transcript"):
                # Generate HTML
                date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                table_rows = ""
                for idx, (_, row) in enumerate(edited_df.iterrows()):
                    table_rows += f"<tr><td style='padding:8px; border-bottom:1px solid #ddd;'>{idx+1}</td><td style='padding:8px; border-bottom:1px solid #ddd;'>{row['Subject']}</td><td style='padding:8px; border-bottom:1px solid #ddd; text-align:center;'><b>{row['Grade']}</b></td></tr>"

                summary_rows = ""
                if include_gpa:
                    summary_rows += f"<tr><td><b>GPA:</b></td><td>{gpa:.2f}</td></tr>"
                    summary_rows += f"<tr><td><b>Class Awarded:</b></td><td>{class_awarded}</td></tr>"
                
                html_content = f"""
                <html>
                <head>
                    <title>Result Sheet - {reg_no}</title>
                    <style>
                        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 25px 35px; color: #333; }}
                        h1,h2 {{ text-align: center; margin: 0; padding: 0; }}
                        h3 {{ text-align: center; color: #666; margin: 5px 0; }}
                        .header-box {{ display: flex; justify-content: space-between; border-top: 2px solid #333; border-bottom: 2px solid #333; padding: 10px 0; margin: 10px 0; }}
                        .header-left, .header-right {{ width: 48%; }}
                        .info-row {{ margin-bottom: 4px; font-size: 14px; }}
                        table {{ width: 100%; border-collapse: collapse; margin-top: 5px; }}
                        th {{ background-color: #f2f2f2; text-align: left; padding: 8px; border-bottom: 2px solid #aaa; font-size: 14px; }}
                        td {{ padding: 6px 8px; font-size: 13px; }}
                        .summary-box {{ margin-top: 15px; padding: 12px; background-color: #f9f9f9; border: 1px solid #ddd; }}
                        
                        /* Footer styling for print */
                        .page-footer {{
                            text-align: center;
                            font-size: 11px;
                            color: #666;
                            margin-top: 25px;
                        }}
                        
                        @media print {{
                            body {{ -webkit-print-color-adjust: exact; margin: 15px 25px; }}
                            
                            /* Minimize spacing for print to fit table on first page */
                            h1, h2 {{ margin: 0; padding: 0; font-size: 18px; }}
                            h3 {{ margin: 3px 0; font-size: 14px; }}
                            .header-box {{ padding: 8px 0; margin: 8px 0; }}
                            .info-row {{ margin-bottom: 2px; font-size: 12px; }}
                            table {{ margin-top: 5px; }}
                            th {{ padding: 6px; font-size: 12px; }}
                            td {{ padding: 5px 6px; font-size: 11px; }}
                            
                            /* Fixed footer on every page */
                            .page-footer {{
                                position: fixed;
                                bottom: 0;
                                left: 0;
                                right: 0;
                                text-align: center;
                                font-size: 9px;
                                color: #888;
                                padding: 8px 0;
                                border-top: 1px solid #ddd;
                                background: white;
                                height: 35px;
                            }}
                            
                            /* Minimal space for footer */
                            body {{
                                margin-bottom: 45px;
                            }}
                            
                            /* Ensure tables don't break awkwardly */
                            .summary-box {{
                                page-break-inside: avoid;
                            }}
                        }}
                    </style>
                </head>
                <body>
                    <h2>SAB Campus of Chartered Accountants Sri Lanka </h2>
                    <h1>Student Result Sheet</h1>
                    <h3>{selected_sheet}</h3>
                    <div class="header-box">
                        <div class="header-left">
                            <div class="info-row"><b>Name:</b> {name}</div>
                            <div class="info-row"><b>Registration No:</b> {reg_no}</div>
                        </div>
                        <div class="header-right">
                            <div class="info-row"><b>Date Issued:</b> {date_str}</div>
                        </div>
                    </div>
                    <table>
                        <thead>
                            <tr>
                                <th width="10%">#</th>
                                <th width="70%">Subject</th>
                                <th width="20%" style="text-align:center;">Grade</th>
                            </tr>
                        </thead>
                        <tbody>
                            {table_rows}
                        </tbody>
                    </table>
                    <div class="summary-box">
                        <table style="margin-top:0; width:50%">
                            {summary_rows}
                        </table>
                    </div>
                    
                    <!-- Footer appears on every page when printed -->
                    <div class="page-footer">
                        <p style="margin: 5px 0;">Generated by SAB Campus - Student Results System</p>
                        <p style="position: absolute; right: 20px; bottom: 10px; margin: 0; font-size: 9px; color: #aaa;">Dev@Salinda</p>
                    </div>
                </body>
                </html>
                """
                
                # JavaScript to open window and print
                js_code = f"""
                <script>
                    var printWindow = window.open('', '_blank');
                    printWindow.document.write({json.dumps(html_content)});
                    printWindow.document.close();
                    printWindow.focus();
                    printWindow.print();
                </script>
                """
                components.html(js_code, height=0, width=0)
            
            # Download Student Results as Excel
            # Create a DataFrame with the student's results
            download_df = edited_df.copy()
            download_df.insert(0, 'Student Name', name)
            download_df.insert(1, 'Registration Number', reg_no)
            download_df.insert(2, 'Batch', selected_sheet)
            if include_gpa:
                download_df['GPA'] = gpa
                download_df['Class'] = class_awarded
            
            # Convert to Excel
            from io import BytesIO
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                download_df.to_excel(writer, index=True, sheet_name='Results')
            excel_buffer.seek(0)
            
            st.sidebar.download_button(
                label="📥 Download Excel",
                data=excel_buffer,
                file_name=f"{reg_no}_{name}_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )





else:
    st.info("Please select a workbook from the sidebar to begin.")

# Theme Selection (Always visible at bottom of sidebar)
st.sidebar.markdown("---")
col1, col2 = st.sidebar.columns([3, 1])
with col1:
    st.sidebar.markdown("**Theme**")
with col2:
    # Toggle button with icons
    current_icon = "☀️" if st.session_state.theme == "Light" else "🌙"
    if st.sidebar.button(current_icon, key="theme_toggle"):
        st.session_state.theme = "Dark" if st.session_state.theme == "Light" else "Light"
        st.rerun()

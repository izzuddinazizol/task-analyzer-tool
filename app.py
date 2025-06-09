import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
import holidays
import io

# =================================================================================
# CONFIGURATION
# =================================================================================

st.set_page_config(
    page_title="Task & Productivity Analyzer",
    page_icon="📊",
    layout="wide"
)

DEFAULT_WORK_START = time(9, 30)
DEFAULT_WORK_END = time(18, 30)

MALAYSIA_STATES = {
    'JHR': 'Johor', 'KDH': 'Kedah', 'KTN': 'Kelantan', 'MLK': 'Melaka',
    'NSN': 'Negeri Sembilan', 'PHG': 'Pahang', 'PNG': 'Pulau Pinang',
    'PRK': 'Perak', 'PLS': 'Perlis', 'SBH': 'Sabah', 'SGR': 'Selangor',
    'SWK': 'Sarawak', 'TRG': 'Terengganu', 'KUL': 'W.P. Kuala Lumpur',
    'LBN': 'W.P. Labuan', 'PJY': 'W.P. Putrajaya'
}

# =================================================================================
# CORE FUNCTIONS
# =================================================================================

def load_css(file_name):
    """Loads a CSS file and injects it into the Streamlit app."""
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.error(f"CSS file '{file_name}' not found. Please make sure it's in the same folder as the app.")

@st.cache_data
def load_data(uploaded_file, sheet_name):
    """Load data from a specific sheet in an Excel file and clean column headers."""
    if uploaded_file is not None and sheet_name is not None:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            st.error(f"Error loading data from sheet '{sheet_name}': {e}")
            return None
    return None

def get_date_range(preset):
    """Calculate date range based on preset selection."""
    today = datetime.now().date()
    if preset == "Today": return today, today
    if preset == "Yesterday":
        yesterday = today - timedelta(days=1)
        return yesterday, yesterday
    if preset == "This Week":
        start = today - timedelta(days=today.weekday())
        return start, today
    if preset == "Last Week":
        start = today - timedelta(days=today.weekday() + 7)
        end = start + timedelta(days=6)
        return start, end
    if preset == "This Month":
        start = today.replace(day=1)
        return start, today
    if preset == "Last Month":
        end = today.replace(day=1) - timedelta(days=1)
        start = end.replace(day=1)
        return start, end
    return None, None

def calculate_working_days(start_timestamp_str, end_timestamp_str, work_start_time, work_end_time, public_holidays):
    """Calculates the number of working days between two dates considering business hours and holidays."""
    if pd.isna(start_timestamp_str) or pd.isna(end_timestamp_str): return np.nan
    start_str, end_str = str(start_timestamp_str).strip(), str(end_timestamp_str).strip()
    formats_to_try = [
        '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %I:%M %p', '%Y-%m-%d %H:%M',
        '%d/%m/%Y %H:%M', '%m/%d/%Y %I:%M %p', '%m/%d/%Y %H:%M',
        '%Y/%m/%d %H:%M', '%Y/%m/%d %I:%M %p'
    ]
    
    start_datetime, end_datetime = None, None
    for fmt in formats_to_try:
        try:
            if not start_datetime: start_datetime = datetime.strptime(start_str, fmt)
        except (ValueError, TypeError): continue
    for fmt in formats_to_try:
        try:
            if not end_datetime: end_datetime = datetime.strptime(end_str, fmt)
        except (ValueError, TypeError): continue
    if start_datetime is None or end_datetime is None: return np.nan
    if start_datetime >= end_datetime: return 0.0

    while True:
        if start_datetime.date() in public_holidays or start_datetime.weekday() >= 5:
            start_datetime = (start_datetime + timedelta(days=1)).replace(hour=work_start_time.hour, minute=work_start_time.minute, second=0, microsecond=0)
            continue
        if start_datetime.time() < work_start_time:
            start_datetime = start_datetime.replace(hour=work_start_time.hour, minute=work_start_time.minute)
        elif start_datetime.time() >= work_end_time:
            start_datetime = (start_datetime + timedelta(days=1)).replace(hour=work_start_time.hour, minute=work_start_time.minute)
            continue
        break

    while True:
        if end_datetime.date() in public_holidays or end_datetime.weekday() >= 5:
            end_datetime = (end_datetime - timedelta(days=1)).replace(hour=work_end_time.hour, minute=work_end_time.minute, second=0, microsecond=0)
            continue
        if end_datetime.time() > work_end_time:
            end_datetime = end_datetime.replace(hour=work_end_time.hour, minute=work_end_time.minute)
        elif end_datetime.time() <= work_start_time:
            end_datetime = (end_datetime - timedelta(days=1)).replace(hour=work_end_time.hour, minute=work_end_time.minute)
            continue
        break
        
    if start_datetime >= end_datetime: return 0.0

    total_working_seconds = 0
    current_process_time = start_datetime
    while current_process_time < end_datetime:
        if current_process_time.weekday() < 5 and current_process_time.date() not in public_holidays:
            day_start = datetime.combine(current_process_time.date(), work_start_time)
            day_end = datetime.combine(current_process_time.date(), work_end_time)
            total_working_seconds += (min(end_datetime, day_end) - max(current_process_time, day_start)).total_seconds()
        current_process_time = (current_process_time.replace(hour=0, minute=0) + timedelta(days=1)).replace(hour=work_start_time.hour, minute=work_start_time.minute)

    working_seconds_per_day = (datetime.combine(datetime.min, work_end_time) - datetime.combine(datetime.min, work_start_time)).total_seconds()
    return total_working_seconds / working_seconds_per_day if working_seconds_per_day > 0 else 0.0

def calculate_total_working_hours(start_date, end_date, work_start_time, work_end_time, public_holidays):
    """Calculates total available working hours in a date range, excluding weekends and holidays."""
    working_hours_per_day = ((datetime.combine(datetime.min, work_end_time) - datetime.combine(datetime.min, work_start_time)).total_seconds() / 3600)
    if working_hours_per_day <= 0 or end_date is None or start_date is None or end_date < start_date: return 0
    total_hours, current_date = 0, start_date
    while current_date <= end_date:
        if current_date.weekday() < 5 and current_date not in public_holidays:
            total_hours += working_hours_per_day
        current_date += timedelta(days=1)
    return total_hours

# =================================================================================
# MAIN APP LOGIC
# =================================================================================

def main():
    load_css("style.css") # Apply custom styles
    st.title("📊 Task & Productivity Analyzer")
    
    with st.sidebar:
        st.header("Global Filters & Settings")
        uploaded_file = st.file_uploader("Upload Task Management Excel File", type=['xlsx'])
        
        df, sheet_names = None, []
        if uploaded_file:
            try: sheet_names = pd.ExcelFile(uploaded_file).sheet_names
            except Exception as e: st.error(f"Could not read Excel file. Error: {e}")
        
        if sheet_names:
            selected_sheet = st.selectbox("Select the sheet to analyze", sheet_names)
            df = load_data(uploaded_file, selected_sheet)
        
        if df is not None:
            date_preset = st.selectbox("Select Date Range", ["This Month", "Last Month", "This Week", "Last Week", "Today", "Yesterday", "Custom Range"])
            start_date, end_date = get_date_range(date_preset)
            if date_preset == "Custom Range":
                start_date = st.date_input("Start Date", value=datetime.now().date())
                end_date = st.date_input("End Date", value=datetime.now().date())

            with st.expander("Column Names & Exclusions"):
                created_date_col, done_timestamp_col, assigned_person_col, ticket_category_col, category_to_exclude = (
                    st.text_input("Created Date Column", value="Created Date"),
                    st.text_input("Done Timestamp Column", value="Done Timestamp"),
                    st.text_input("Assigned Person Column", value="Person"),
                    st.text_input("Ticket Category Column", value="Ticket Category"),
                    st.text_input("Exclude this Category", value="Renewal - Account Renewal")
                )

            player_options = sorted([str(p) for p in df[assigned_person_col].dropna().unique()]) if assigned_person_col in df.columns else []
            selected_players = st.multiselect("Filter by Players", player_options, default=player_options)
            
            category_options = sorted([str(c) for c in df[ticket_category_col].dropna().unique()]) if ticket_category_col in df.columns else []
            selected_categories = st.multiselect("Filter by Categories", category_options, default=category_options)
            
            with st.expander("Business Rules", expanded=True):
                work_start, work_end = st.time_input("Working Day Start Time", value=DEFAULT_WORK_START), st.time_input("Working Day End Time", value=DEFAULT_WORK_END)
                
                state_name_options = ["National (Federal Only)"] + sorted(list(MALAYSIA_STATES.values()))
                try: default_index = state_name_options.index("Selangor")
                except ValueError: default_index = 0
                
                selected_state_name = st.selectbox("Select State", state_name_options, index=default_index)

                holiday_dict = {}
                if start_date and end_date:
                    analysis_years = range(start_date.year, end_date.year + 1)
                    if selected_state_name == "National (Federal Only)": holiday_dict = holidays.MY(years=analysis_years)
                    else:
                        state_code = [code for code, name in MALAYSIA_STATES.items() if name == selected_state_name][0]
                        holiday_dict = holidays.MY(subdiv=state_code, years=analysis_years)
                
                holiday_options = [f"{day.strftime('%Y-%m-%d')}: {name}" for day, name in sorted(holiday_dict.items()) if start_date <= day <= end_date]
                selected_holidays_str = st.multiselect(
                    "Select Public Holidays to Exclude", options=holiday_options, default=holiday_options,
                    help="Deselect any day you want to treat as a normal workday."
                )

    if df is None:
        st.info("👋 Welcome! Please upload your Excel file using the sidebar to begin.")
        return

    public_holidays = {datetime.strptime(s.split(':')[0], '%Y-%m-%d').date() for s in selected_holidays_str}
    
    essential_cols = [created_date_col, done_timestamp_col, assigned_person_col, ticket_category_col]
    if not all(col in df.columns for col in essential_cols):
        st.error(f"One or more essential columns are missing. Check 'Column Names' settings. Required: {essential_cols}")
        return

    df[created_date_col] = pd.to_datetime(df[created_date_col], errors='coerce')
    df_filtered = df.dropna(subset=[created_date_col])
    if category_to_exclude: df_filtered = df_filtered[df_filtered[ticket_category_col] != category_to_exclude]

    mask = ( (df_filtered[created_date_col].dt.date >= start_date) & (df_filtered[created_date_col].dt.date <= end_date) &
             (df_filtered[assigned_person_col].isin(selected_players)) & (df_filtered[ticket_category_col].isin(selected_categories)) )
    df_final = df_filtered[mask].copy()

    if df_final.empty:
        st.warning("No data found for the selected filters.")
        return
        
    tab1, tab2 = st.tabs(["Resolution Time Analysis", "Player Productivity Analysis"])

    with st.spinner("Analyzing data... Please wait."):
        df_final['Resolution Time (WD)'] = df_final.apply(lambda row: calculate_working_days(row[created_date_col], row[done_timestamp_col], work_start, work_end, public_holidays), axis=1)
        df_with_valid_res_time = df_final.dropna(subset=['Resolution Time (WD)'])

    with tab1:
        st.header("Resolution Time Analysis")
        if df_with_valid_res_time.empty:
            st.warning("Could not calculate resolution time for any tasks in the selected range.")
        else:
            avg_res_time = df_with_valid_res_time.groupby(assigned_person_col)['Resolution Time (WD)'].mean()
            team_avg_res_time = df_with_valid_res_time['Resolution Time (WD)'].mean()
            task_counts = df_final.groupby([assigned_person_col, ticket_category_col]).size().unstack(fill_value=0)
            task_counts['Total Tasks (Player)'] = task_counts.sum(axis=1)
            consolidated_table = task_counts.copy()
            consolidated_table['Average Resolution Time (WD)'] = avg_res_time
            team_total_row = pd.DataFrame(consolidated_table.sum(axis=0)).T
            team_total_row.index = ['Team Total']
            team_total_row['Average Resolution Time (WD)'] = team_avg_res_time
            final_table = pd.concat([consolidated_table, team_total_row])
            st.metric("Overall Team Average Resolution Time", f"{team_avg_res_time:.2f} working days")
            st.dataframe(final_table.fillna(0).round(2))
            csv = final_table.to_csv().encode('utf-8')
            st.download_button("Download Resolution Analysis (CSV)", data=csv, file_name="resolution_time_analysis.csv")

    with tab2:
        st.header("Player Productivity Analysis")
        calc_method = st.radio("Select Calculation Method", ["Use actual calculated resolution time", "Use standard time estimates per category"], key="prod_method")
        
        standard_times = {}
        if calc_method == "Use standard time estimates per category":
            st.subheader("Standard Time Estimates (in MINUTES)")
            for category in sorted(df_final[ticket_category_col].unique()):
                standard_times[category] = st.number_input(f"Time for '{category}'", min_value=0.0, value=60.0, step=15.0, key=f"time_{category}")

        available_hours = calculate_total_working_hours(start_date, end_date, work_start, work_end, public_holidays)
        available_minutes = available_hours * 60
        
        working_hours_per_day = ((datetime.combine(datetime.min, work_end) - datetime.combine(datetime.min, work_start)).total_seconds() / 3600)
        
        productivity_data = []
        for player in sorted(df_final[assigned_person_col].unique()):
            player_df = df_with_valid_res_time[df_with_valid_res_time[assigned_person_col] == player]
            player_all_tasks_df = df_final[df_final[assigned_person_col] == player]
            total_tasks = len(player_all_tasks_df)
            total_time_spent_minutes = 0

            if calc_method == "Use actual calculated resolution time":
                total_time_spent_minutes = player_df['Resolution Time (WD)'].sum() * working_hours_per_day * 60
            else:
                total_time_spent_minutes = sum(standard_times.get(cat, 0) for cat in player_all_tasks_df[ticket_category_col])
            
            productivity = (total_time_spent_minutes / available_minutes * 100) if available_minutes > 0 else 0
            
            assessment = "Needs Improvement"
            if productivity >= 80: assessment = "Productive"
            if productivity >= 95: assessment = "Excellent"
            
            productivity_data.append({
                'Person': player,
                'Total Tasks Completed': total_tasks,
                'Total Time Spent (Minutes)': total_time_spent_minutes,
                'Available Working Minutes': available_minutes,
                'Productivity (%)': productivity,
                'Assessment': assessment
            })
        
        if productivity_data:
            productivity_df = pd.DataFrame(productivity_data)
            st.dataframe(productivity_df.round(2))
            
            team_total_available_minutes = len(player_options) * available_minutes
            overall_productivity = (productivity_df['Total Time Spent (Minutes)'].sum() / team_total_available_minutes * 100) if team_total_available_minutes > 0 else 0
            st.metric("Overall Team Productivity", f"{overall_productivity:.2f}%")
            
            csv_prod = productivity_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download Productivity Analysis (CSV)", data=csv_prod, file_name="productivity_analysis.csv")

if __name__ == "__main__":
    main()
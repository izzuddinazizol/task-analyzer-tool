import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, time
import holidays
import io
import plotly.express as px

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
# CORE FUNCTIONS (Cached for Performance)
# =================================================================================

def load_css(file_name):
    """Loads a CSS file and injects it into the Streamlit app."""
    try:
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except FileNotFoundError:
        st.warning(f"CSS file '{file_name}' not found. App will use default styling.")

@st.cache_data
def load_data(uploaded_file, sheet_name):
    """Load data from a specific sheet in an Excel file."""
    if uploaded_file and sheet_name:
        try:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            st.error(f"Error loading sheet '{sheet_name}': {e}")
    return None

def get_date_range(preset):
    """Calculate date range based on preset selection."""
    today = datetime.now().date()
    # ... (date range logic remains the same)
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
    return datetime.now().date(), datetime.now().date()

def calculate_working_days(row, created_col, done_col, work_start_time, work_end_time, public_holidays):
    """Calculates resolution time in working days for a row of a DataFrame."""
    start_timestamp_str = row[created_col]
    end_timestamp_str = row[done_col]
    
    # ... (The core calculation logic remains the same)
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
        if start_datetime.time() < work_start_time: start_datetime = start_datetime.replace(hour=work_start_time.hour, minute=work_start_time.minute)
        elif start_datetime.time() >= work_end_time:
            start_datetime = (start_datetime + timedelta(days=1)).replace(hour=work_start_time.hour, minute=work_start_time.minute)
            continue
        break

    while True:
        if end_datetime.date() in public_holidays or end_datetime.weekday() >= 5:
            end_datetime = (end_datetime - timedelta(days=1)).replace(hour=work_end_time.hour, minute=work_end_time.minute, second=0, microsecond=0)
            continue
        if end_datetime.time() > work_end_time: end_datetime = end_datetime.replace(hour=work_end_time.hour, minute=work_end_time.minute)
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
    """Calculates total available working hours in a date range."""
    working_hours_per_day = ((datetime.combine(datetime.min, work_end_time) - datetime.combine(datetime.min, work_start_time)).total_seconds() / 3600)
    if working_hours_per_day <= 0 or end_date is None or start_date is None or end_date < start_date: return 0
    total_hours = 0
    current_date = start_date
    while current_date <= end_date:
        if current_date.weekday() < 5 and current_date not in public_holidays:
            total_hours += working_hours_per_day
        current_date += timedelta(days=1)
    return total_hours

# =================================================================================
# MAIN APP LOGIC
# =================================================================================
def main():
    load_css("style.css")
    st.title("📊 Task & Productivity Analyzer")

    # --- NEW LAYOUT: Uploader is in the main area ---
    uploaded_file = st.file_uploader(
        "Upload Your Task Management Excel File to Begin",
        type=['xlsx']
    )

    if uploaded_file is None:
        st.info("👋 Welcome! Please upload a file to start the analysis.")
        return

    # --- All filters and logic now run only AFTER a file is uploaded ---
    sheet_names = pd.ExcelFile(uploaded_file).sheet_names
    
    # Initialize session state for widgets to remember their state
    if 'selected_sheet' not in st.session_state:
        st.session_state.selected_sheet = sheet_names[0]

    # --- SIDEBAR CONTROLS ---
    with st.sidebar:
        st.header("Global Filters & Settings")
        selected_sheet = st.selectbox("1. Select the sheet to analyze", sheet_names, key='selected_sheet')
        
        df = load_data(uploaded_file, selected_sheet)
        if df is None:
            st.warning("Could not load data from the selected sheet.")
            return

        date_preset = st.selectbox("2. Select Date Range", ["This Month", "Last Month", "This Week", "Last Week", "Today", "Yesterday", "Custom Range"])
        start_date, end_date = get_date_range(date_preset)
        if date_preset == "Custom Range":
            start_date = st.date_input("Start Date", value=start_date)
            end_date = st.date_input("End Date", value=end_date)

        with st.expander("Column Names & Exclusions"):
            created_date_col, done_timestamp_col, assigned_person_col, ticket_category_col, category_to_exclude = (
                st.text_input("Created Date Column", value="Created Date"),
                st.text_input("Done Timestamp Column", value="Done Timestamp"),
                st.text_input("Assigned Person Column", value="Person"),
                st.text_input("Ticket Category Column", value="Ticket Category"),
                st.text_input("Exclude this Category", value="Renewal - Account Renewal")
            )
        
        player_options = sorted([str(p) for p in df[assigned_person_col].dropna().unique()]) if assigned_person_col in df.columns else []
        selected_players = st.multiselect("3. Filter by Players", player_options, default=player_options)
        
        category_options = sorted([str(c) for c in df[ticket_category_col].dropna().unique()]) if ticket_category_col in df.columns else []
        selected_categories = st.multiselect("4. Filter by Categories", category_options, default=category_options)
        
        with st.expander("Business Rules", expanded=True):
            work_start, work_end = st.time_input("Working Day Start Time", value=DEFAULT_WORK_START), st.time_input("Working Day End Time", value=DEFAULT_WORK_END)
            state_name_options = ["National (Federal Only)"] + sorted(list(MALAYSIA_STATES.values()))
            try: default_index = state_name_options.index("Selangor")
            except ValueError: default_index = 0
            selected_state_name = st.selectbox("Select State", state_name_options, index=default_index)

    # --- DATA PROCESSING (Main Page) ---
    public_holidays = {}
    if start_date:
        analysis_years = range(start_date.year, end_date.year + 1)
        if selected_state_name == "National (Federal Only)": public_holidays = holidays.MY(years=analysis_years)
        else:
            state_code = [code for code, name in MALAYSIA_STATES.items() if name == selected_state_name][0]
            public_holidays = holidays.MY(subdiv=state_code, years=analysis_years)
    
    # Filter and prepare the final DataFrame
    df[created_date_col] = pd.to_datetime(df[created_date_col], errors='coerce')
    df_filtered = df.dropna(subset=[created_date_col])
    if category_to_exclude: df_filtered = df_filtered[df_filtered[ticket_category_col] != category_to_exclude]
    mask = ((df_filtered[created_date_col].dt.date >= start_date) & (df_filtered[created_date_col].dt.date <= end_date) &
            (df_filtered[assigned_person_col].isin(selected_players)) & (df_filtered[ticket_category_col].isin(selected_categories)))
    df_final = df_filtered[mask].copy()

    if df_final.empty:
        st.warning("No data found for the selected filters. Please adjust your selections.")
        return

    # --- TABS & DASHBOARD DISPLAY ---
    tab1, tab2 = st.tabs(["Resolution Time Dashboard", "Productivity Dashboard"])

    with st.spinner("Analyzing data... Please wait."):
        df_final['Resolution Time (WD)'] = df_final.apply(calculate_working_days, args=(created_date_col, done_timestamp_col, work_start, work_end, public_holidays), axis=1)
        df_with_valid_res_time = df_final.dropna(subset=['Resolution Time (WD)'])

    with tab1:
        st.header("Resolution Time Dashboard")
        if not df_with_valid_res_time.empty:
            team_avg_res_time = df_with_valid_res_time['Resolution Time (WD)'].mean()
            
            # --- Key Metrics ---
            col1, col2, col3 = st.columns(3)
            col1.metric("Overall Avg. Resolution Time", f"{team_avg_res_time:.2f} WD")
            col2.metric("Total Tasks Analyzed", f"{len(df_final)}")
            col3.metric("Tasks with Valid Time", f"{len(df_with_valid_res_time)}")

            # --- Charts ---
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Avg. Resolution Time per Player")
                avg_res_time_per_player = df_with_valid_res_time.groupby(assigned_person_col)['Resolution Time (WD)'].mean().sort_values()
                fig = px.bar(avg_res_time_per_player, x=avg_res_time_per_player.values, y=avg_res_time_per_player.index, orientation='h', labels={'x':'Avg. Resolution Time (Working Days)', 'y':''})
                st.plotly_chart(fig, use_container_width=True)
            with col2:
                st.subheader("Task Count by Category")
                category_counts = df_final[ticket_category_col].value_counts()
                fig = px.pie(category_counts, values=category_counts.values, names=category_counts.index, hole=0.3)
                st.plotly_chart(fig, use_container_width=True)

            # --- Data Table ---
            st.markdown("---")
            st.subheader("Detailed Resolution Time Report")
            # ... (your existing table generation logic can go here)
        else:
            st.warning("Could not calculate resolution time for any tasks in the selected range.")

    with tab2:
        st.header("Productivity Dashboard")
        # ... (your existing productivity logic, now with charts)
        st.info("Productivity analysis is under construction.")

if __name__ == "__main__":
    main()
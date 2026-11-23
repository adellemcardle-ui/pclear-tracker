import streamlit as st
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
import io
import base64
import pytz
import os
import json
from datetime import datetime
from streamlit_float import *
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
from st_aggrid.shared import GridUpdateMode

st.set_page_config(layout="wide", page_title="Monash Medicine pCLEAR Assessment Tracker")
float_init()

def get_categorical_grade(percentage):

    thresholds = (
        st.session_state.grade_thresholds
    )

    if percentage >= thresholds["Commendation"]:

        return "Commendation"

    elif percentage >= thresholds["At Standard"]:

        return "At Standard"

    elif percentage >= thresholds["Not Yet At Standard"]:

        return "Not Yet At Standard"

    else:

        return (
            "Concerning / "
            "Unacceptable Behaviour"
        )
GRADE_CODE_MAP = {
    "Commendation": "C",
    "At Standard": "AS",
    "Not Yet At Standard": "NY",
    "Concerning / Unacceptable Behaviour": "CU",
}

def get_categorical_grade_code(percentage: float) -> str:
    full_label = get_categorical_grade(percentage)
    return GRADE_CODE_MAP.get(full_label, full_label)

SCORE_COLUMNS = [
    "Engagement",
    "Teamwork",
    "Written Tasks",
    "Patience/Respect",
    "Accountability",
    "Uncertainty",
    "Punctuality"
]

def calculate_domain_totals(
    student_df,
    domain,
    start_week,
    end_week
):

    actual = 0
    possible = 0

    for _, record in student_df.iterrows():
        try:
            week = int(float(record["Week"]))
        except (ValueError, TypeError):
            continue

        if week < start_week or week > end_week:
            continue

        value = str(
            record[domain]
        ).strip()

        if value.isdigit():

            actual += int(value)
            possible += 5

    return actual, possible

def calculate_week_totals(
    row,
    score_columns
):

    weekly_total = 0
    weekly_possible = 0

    for col in score_columns:

        value = str(row[col]).strip()

        if value.isdigit():

            weekly_total += int(value)
            weekly_possible += 5

    return weekly_total, weekly_possible

SEMESTER_1_WEEKS = list(range(1, 16))   # Weeks 1–15
SEMESTER_2_WEEKS = list(range(16, 29))  # Weeks 16–32

def get_semester_weeks(period: str):
    """Return the list of valid teaching weeks for a given reporting period."""
    if period == "Semester 1":
        return SEMESTER_1_WEEKS
    elif period == "Semester 2":
        return SEMESTER_2_WEEKS
    return None  # "Full Year" or "All"

def filter_by_period(df: pd.DataFrame, period: str, week_col: str = "Week") -> pd.DataFrame:
    """Filter any DataFrame (assessments or attendance) to the selected semester."""
    valid_weeks = get_semester_weeks(period)
    if valid_weeks is None or df.empty or week_col not in df.columns:
        return df
    numeric_weeks = pd.to_numeric(df[week_col], errors="coerce")
    return df[numeric_weeks.isin(valid_weeks)]

def save_ui_state():
    pd.to_pickle(
        {
            "selected_user": st.session_state.get("last_user", "Coordinator"),
            "page": st.session_state.get("last_page", "Assessment Analytics"),
            "saved_analytics_tab": st.session_state.get("saved_analytics_tab", "Achievement Matrix"),
            "saved_period": st.session_state.get("saved_period", "Semester 1"),
            "saved_group": st.session_state.get("saved_group", "All Groups"),
            "saved_sort_by": st.session_state.get("saved_sort_by", "Last Name (A → Z)"),
            "saved_view_mode": st.session_state.get("saved_view_mode", "Full Detail (All Weeks)"),
            # Mail Merge states
            "saved_mm_sort_by": st.session_state.get("saved_mm_sort_by", "Last Name (A → Z)"),
            "saved_mm_search": st.session_state.get("saved_mm_search", ""),
            # Validation states
            "saved_val_student": st.session_state.get("saved_val_student", None),
            "saved_val_week": st.session_state.get("saved_val_week", "All Weeks"),
            "saved_val_filter": st.session_state.get("saved_val_filter", "All Records"),
        },
        "ui_state.pkl"
    )

def save_analytics_tab(widget_key: str):
    st.session_state.saved_analytics_tab = st.session_state[widget_key]
    save_ui_state()

def save_period(widget_key: str):
    st.session_state.saved_period = st.session_state[widget_key]
    save_ui_state()

def save_group(widget_key: str):
    val = st.session_state[widget_key]
    if isinstance(val, str) and val.startswith("Group "):
        raw = val.replace("Group ", "", 1).strip()
        if not st.session_state.student_data.empty and "Group" in st.session_state.student_data.columns:
            sample_groups = st.session_state.student_data["Group"].dropna().tolist()
            if sample_groups and isinstance(sample_groups[0], (int, float)) and raw.isdigit():
                val = int(raw)
            else:
                val = raw
        else:
            val = int(raw) if raw.isdigit() else raw
    st.session_state.saved_group = val
    save_ui_state()

def save_sort_by(widget_key: str):
    st.session_state.saved_sort_by = st.session_state[widget_key]
    save_ui_state()

def save_view_mode(widget_key: str):
    st.session_state.saved_view_mode = st.session_state[widget_key]
    save_ui_state()

def save_mm_sort_by(widget_key: str):
    st.session_state.saved_mm_sort_by = st.session_state[widget_key]
    save_ui_state()

def save_mm_search(widget_key: str):
    st.session_state.saved_mm_search = st.session_state[widget_key]
    save_ui_state()

def save_val_student(widget_key: str):
    st.session_state.saved_val_student = st.session_state[widget_key]
    save_ui_state()

def save_val_week(widget_key: str):
    st.session_state.saved_val_week = st.session_state[widget_key]
    save_ui_state()

def save_val_filter(widget_key: str):
    st.session_state.saved_val_filter = st.session_state[widget_key]
    save_ui_state()

def format_weekly_total_cell(week_record: pd.DataFrame) -> str:
    """Return the weekly total score as a string (skipping N/A and I/O), or '' if no numeric scores exist."""
    if week_record.empty:
        return ""
    r = week_record.iloc[0]
    total = 0
    has_numeric = False
    for col in SCORE_COLUMNS:
        raw = str(r.get(col, "")).strip()
        if raw.upper() in ("", "NAN", "NONE", "---", "N/A", "I/O"):
            continue
        try:
            v = float(raw)
            if v.is_integer() and 0 <= int(v) <= 5:
                total += int(v)
                has_numeric = True
        except ValueError:
            pass
    return str(total) if has_numeric else ""

# Helper to calculate overall % across the active reporting period for sorting
def get_student_overall_pct(student_name: str, assessment_df: pd.DataFrame, reporting_period: str) -> float:
    s_df = assessment_df[assessment_df["Name"] == student_name]
    if reporting_period == "Semester 1":
        s_df = s_df[s_df["Week"].isin(SEMESTER_1_WEEKS)]
    elif reporting_period == "Semester 2":
        s_df = s_df[s_df["Week"].isin(SEMESTER_2_WEEKS)]
    else:
        s_df = s_df[s_df["Week"].isin(SEMESTER_1_WEEKS + SEMESTER_2_WEEKS)]

    total_actual = 0
    total_possible = 0
    for _, r in s_df.iterrows():
        act, poss = calculate_week_totals(r, SCORE_COLUMNS)
        total_actual += act
        total_possible += poss
    return (total_actual / total_possible * 100) if total_possible > 0 else 999.0

def build_template_matrix_excel(
    student_df: pd.DataFrame,
    assessment_df: pd.DataFrame,
    selected_group,
) -> bytes:
    cohort = student_df.copy()

    # Build Full Name reliably from preferred name + family name (or existing Name column)
    if "preferred name" in cohort.columns or "family name" in cohort.columns:
        pref = cohort["preferred name"].fillna("").astype(str).str.strip() if "preferred name" in cohort.columns else ""
        fam = cohort["family name"].fillna("").astype(str).str.strip() if "family name" in cohort.columns else ""
        cohort["_FullName"] = (pref + " " + fam).str.strip()
    elif "Name" in cohort.columns:
        cohort["_FullName"] = cohort["Name"].fillna("").astype(str).str.strip()
    else:
        cohort["_FullName"] = ""

    # Filter cohort if a specific group is selected (handles both "Group 1" and "1")
    if selected_group != "All Groups" and "Group" in cohort.columns:
        clean_grp = str(selected_group).replace("Group ", "").strip()
        cohort = cohort[
            cohort["Group"].astype(str).str.replace("Group ", "", regex=False).str.strip() == clean_grp
        ]

    # Look up Student ID, First Name, Surname columns automatically
    id_col = next((c for c in cohort.columns if c.lower() == "id" or "student id" in c.lower()), None)
    fn_col = next((c for c in cohort.columns if "preferred" in c.lower() or "first" in c.lower()), None)
    sn_col = next((c for c in cohort.columns if "family" in c.lower() or "last" in c.lower() or "sur" in c.lower()), None)

    domain_headers = {
        "Engagement": "Engagement",
        "Teamwork": "Teamwork",
        "Written Tasks": "Written Tasks",
        "Patience/Respect": "Patience, Courtesy, Respect & Reflection",
        "Accountability": "Accountability",
        "Uncertainty": "Dealing with Learning Ambiguity",
        "Punctuality": "Punctuality",
    }

    sem1_weeks = list(range(1, 16))   # Weeks 1–14
    sem2_weeks = list(range(16, 29))  # Weeks 15–28

    assess_copy = assessment_df.copy()
    if "Name" in assess_copy.columns:
        assess_copy["_MatchName"] = assess_copy["Name"].fillna("").astype(str).str.strip()
    else:
        assess_copy["_MatchName"] = ""
    assess_copy["Week_num"] = pd.to_numeric(assess_copy["Week"], errors="coerce")

    export_rows = []

    for _, stu in cohort.iterrows():
        full_name = stu.get("_FullName", "")
        if not full_name:
            continue

        stu_id = str(stu.get(id_col, "")).strip() if id_col else ""
        first_name = str(stu.get(fn_col, "")).strip() if fn_col else full_name.split(" ")[0]
        surname = str(stu.get(sn_col, "")).strip() if sn_col else " ".join(full_name.split(" ")[1:])

        # Match by ID first if available, otherwise by Full Name
        if stu_id and "ID" in assess_copy.columns:
            s_assess = assess_copy[assess_copy["ID"].astype(str).str.strip() == stu_id].copy()
            if s_assess.empty:
                s_assess = assess_copy[assess_copy["_MatchName"] == full_name].copy()
        else:
            s_assess = assess_copy[assess_copy["_MatchName"] == full_name].copy()

        def make_row(label, values_dict, is_first_row=False):
            row = {
                "Student ID": stu_id if is_first_row else "",
                "Student Firstname": first_name if is_first_row else "",
                "Student surname": surname if is_first_row else "",
                "DOMAINS": label,
            }
            for raw_col, disp_col in domain_headers.items():
                row[disp_col] = values_dict.get(raw_col, "")
            row["Total Score"] = values_dict.get("Total Score", "")
            return row

        def build_week_rows(weeks_list, mark_first=False):
            for idx, wk in enumerate(weeks_list):
                wk_rows = s_assess[s_assess["Week_num"] == wk]
                vals = {}
                row_total = 0
                has_any = False
                for col in SCORE_COLUMNS:
                    if not wk_rows.empty:
                        raw_v = str(wk_rows.iloc[0].get(col, "")).strip()
                        try:
                            v_float = float(raw_v)
                            if v_float.is_integer():
                                v_int = int(v_float)
                                vals[col] = v_int
                                row_total += v_int
                                has_any = True
                            else:
                                vals[col] = ""
                        except ValueError:
                            vals[col] = ""
                    else:
                        vals[col] = ""
                vals["Total Score"] = row_total if has_any else ""
                export_rows.append(
                    make_row(f"Week {wk}", vals, is_first_row=(mark_first and idx == 0))
                )

        def build_summary_block(weeks_list, prefix):
            sub_df = s_assess[s_assess["Week_num"].isin(weeks_list)]
            subtotals, pcts, grades, max_scores = {}, {}, {}, {}
            tot_act, tot_poss = 0, 0

            for col in SCORE_COLUMNS:
                act, poss = calculate_domain_totals(sub_df, col, min(weeks_list), max(weeks_list))
                pct = (act / poss * 100) if poss > 0 else 0.0
                subtotals[col] = act
                pcts[col] = f"{round(pct)}%"
                grades[col] = get_categorical_grade(pct)
                max_scores[col] = poss
                tot_act += act
                tot_poss += poss

            tot_pct = (tot_act / tot_poss * 100) if tot_poss > 0 else 0.0
            subtotals["Total Score"] = tot_act
            pcts["Total Score"] = f"{round(tot_pct)}%"
            grades["Total Score"] = get_categorical_grade(tot_pct)
            max_scores["Total Score"] = tot_poss

            export_rows.append(make_row(f"{prefix} - Student Subtotal", subtotals))
            export_rows.append(make_row(f"{prefix} - Student Percentage", pcts))
            export_rows.append(make_row(f"{prefix} - Categorical grade", grades))
            return max_scores

        # Semester 1 (Weeks 1–14) + Summary
        build_week_rows(sem1_weeks, mark_first=True)
        sem1_max = build_summary_block(sem1_weeks, "SEM 1")

        # Semester 2 (Weeks 15–28) + Summary
        build_week_rows(sem2_weeks, mark_first=False)
        sem2_max = build_summary_block(sem2_weeks, "SEM 2")

        # Full Year Summary
        full_max = build_summary_block(sem1_weeks + sem2_weeks, "Fullyear")

        # Maximum Possible Score rows
        export_rows.append(make_row("SEM 1 - Maximum Possible Score", sem1_max))
        export_rows.append(make_row("SEM 2 - Maximum Possible Score", sem2_max))
        export_rows.append(make_row("Fullyear - Maximum Possible Score", full_max))

    out_df = pd.DataFrame(export_rows)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Achievement Matrix")
    return buf.getvalue()

# --- 1. PERSISTENT MEMORY ---
if 'student_data' not in st.session_state:
    st.session_state.student_data = pd.DataFrame()
if 'image_library' not in st.session_state:
    st.session_state.image_library = {}
if 'calendar_data' not in st.session_state:
    st.session_state.calendar_data = {}
if 'assessments' not in st.session_state:
    st.session_state.assessments = pd.DataFrame()

    # AUTO-LOAD SAVED DATA

if os.path.exists("cohort.pkl"):
    st.session_state.student_data = pd.read_pickle("cohort.pkl")

if os.path.exists("calendar.pkl"):
    st.session_state.calendar_data = pd.read_pickle(
        "calendar.pkl"
    ).to_dict("records")
if os.path.exists("images.pkl"):
    st.session_state.image_library = pd.read_pickle(
        "images.pkl"
    )
if os.path.exists("assessments.pkl"):
    st.session_state.assessments = pd.read_pickle(
        "assessments.pkl"
    )

if not st.session_state.assessments.empty and "Week" in st.session_state.assessments.columns:
    st.session_state.assessments["Week"] = pd.to_numeric(
        st.session_state.assessments["Week"], errors="coerce"
    )
    st.session_state.assessments["ID"] = st.session_state.assessments["ID"].astype(str).str.strip()
    st.session_state.assessments = st.session_state.assessments.drop_duplicates(
        subset=["ID", "Week"], keep="last"
    ).reset_index(drop=True)

if os.path.exists("grade_thresholds.pkl"):
    st.session_state.grade_thresholds = pd.read_pickle(
        "grade_thresholds.pkl"
    )
else:
    st.session_state.grade_thresholds = {
        "Commendation": 81,
        "At Standard": 60,
        "Not Yet At Standard": 21
    }

if os.path.exists("ui_state.pkl"):
    ui_state = pd.read_pickle("ui_state.pkl")
    st.session_state.last_user = ui_state.get("selected_user", "Coordinator")
    st.session_state.last_page = ui_state.get("page", "Assessment Analytics")
    for key in [
        "saved_analytics_tab",
        "saved_period",
        "saved_group",
        "saved_sort_by",
        "saved_view_mode",
        "saved_mm_sort_by",
        "saved_mm_search",
        "saved_val_student",
        "saved_val_week",
        "saved_val_filter",
    ]:
        if key not in st.session_state and key in ui_state:
            st.session_state[key] = ui_state[key]

if "saved_analytics_tab" not in st.session_state:
    st.session_state.saved_analytics_tab = "Achievement Matrix"

if "saved_period" not in st.session_state:
    max_week = pd.to_numeric(
        st.session_state.assessments.get("Week", pd.Series(dtype="float")),
        errors="coerce"
    ).max()
    st.session_state.saved_period = (
        "Semester 2" if pd.notna(max_week) and max_week >= min(SEMESTER_2_WEEKS) else "Semester 1"
    )

if "saved_group" not in st.session_state:
    st.session_state.saved_group = "All Groups"

if "saved_sort_by" not in st.session_state:
    st.session_state.saved_sort_by = "Last Name (A → Z)"

if "saved_view_mode" not in st.session_state:
    st.session_state.saved_view_mode = "Full Detail (All Weeks)"

if "saved_mm_sort_by" not in st.session_state:
    st.session_state.saved_mm_sort_by = "Last Name (A → Z)"

if "saved_mm_search" not in st.session_state:
    st.session_state.saved_mm_search = ""

if "saved_val_student" not in st.session_state:
    st.session_state.saved_val_student = None

if "saved_val_week" not in st.session_state:
    st.session_state.saved_val_week = "All Weeks"

if "saved_val_filter" not in st.session_state:
    st.session_state.saved_val_filter = "All Records"

# --- 2. SIDEBAR NAVIGATION & SECURE LOGIN ---
with st.sidebar:
    st.title("System Menu")
    
    # USER ROLES AND GROUP ASSIGNMENTS

    USERS = {

        "CAPs Group 1": {
            "role": "Tutor",
            "groups": [1]
        },

        "CAPs Group 2": {
            "role": "Tutor",
            "groups": [2]
        },

        "CAPs Group 3": {
            "role": "Tutor",
            "groups": [3]
        },

        "CAPs Group 4": {
            "role": "Tutor",
            "groups": [4]
        },

        "CAPs Group 5": {
            "role": "Tutor",
            "groups": [5]
        },

        "CAPs Group 6": {
            "role": "Tutor",
            "groups": [6]
        },

        "CAPs Group 7": {
            "role": "Tutor",
            "groups": [7]
        },
        "CAPs Group 8": {
                "role": "Tutor",
                "groups": [8]
            },
        "CAPs Group 9": {
                    "role": "Tutor",
                    "groups": [9]
                },

        "Relief Tutor": {
            "role": "Tutor",
            "groups": [1, 2, 3, 4, 5, 6, 7, 8, 9]
        },

        "Assessment Team": {
            "role": "AssessmentTeam",
            "groups": ["ALL"]
        },

        "Coordinator": {
            "role": "Coordinator",
            "groups": ["ALL"]
        }
    }
        
    st.subheader("👤 User Access")

    if "last_user" not in st.session_state:
        st.session_state.last_user = "Coordinator"

    selected_user = st.selectbox(
        "Current Access",
        list(USERS.keys()),
        index=list(USERS.keys()).index(
            st.session_state.last_user
        )
    )

    st.session_state.last_user = selected_user

    st.session_state['active_user'] = selected_user
    st.session_state['role'] = USERS[selected_user]["role"]
    st.session_state['groups'] = USERS[selected_user]["groups"]

    st.success(
        f"Current Access: {selected_user}"
    )

    st.divider()


    # Navigation menu
    role = st.session_state.get("role")

    if role == "Tutor":

        pages = [
            "Weekly Assessment",
            "Dashboard"
        ]

    else:

        pages = [
            "Weekly Assessment",
            "Dashboard",
            "Assessment Analytics",
            "Setup Import"
        ]

    # First run defaults

    if "last_page" not in st.session_state:
        if role == "Tutor":
            st.session_state.last_page = "Dashboard"
        else:
            st.session_state.last_page = "Assessment Analytics"

    if "pending_page" in st.session_state:
        st.session_state.last_page = st.session_state.pop("pending_page")
        save_ui_state()

    # Guard against a saved page name that is no longer in `pages`
    if st.session_state.get("last_page") not in pages:
        st.session_state.last_page = pages[0]

    def on_nav_change():
        st.session_state.last_page = st.session_state.nav_page_select
        st.session_state.pop("jump_to_assessment", None)
        save_ui_state()

    page = st.selectbox(
        "Navigate",
        pages,
        index=pages.index(st.session_state.last_page),
        key="nav_page_select",
        on_change=on_nav_change,
    )

    st.divider()

    try:
        st.markdown(
            """
        **📝 Comment Guide**

        • **0 & 5:** Comment required  
        • **1, 2 & 4:** Comment preferred
        """
        )
        st.image(
            "scoring_guide.png",
            caption="Scoring Guide Reference",
            use_column_width=True
        )
        
    except:
        pass

# Global variable update for the rest of the app to use
active_user = st.session_state.get('active_user', None)
is_logged_in = True

# ALL THE COMBINED BITS TO CALL UPON.

def render_individual_student_summary(student_df):

    actual_total = 0
    possible_total = 0

    weekly_results = []


    for _, row in student_df.iterrows():

        attendance = str(row["Attendance"])

        weekly_total, weekly_possible = (
            calculate_week_totals(
                row,
                SCORE_COLUMNS
            )
        )

        actual_total += weekly_total
        possible_total += weekly_possible
        weekly_percentage = 0

        if weekly_possible > 0:
            weekly_percentage = (
                weekly_total
                / weekly_possible
            ) * 100

        attendance_icon = "✅"

        if attendance == "Absent - notified":
            attendance_icon = "🟣"

        elif attendance == "Absent - unnotified":
            attendance_icon = "🔴"

        weekly_results.append(
            {
                "Week": int(float(row["Week"])),
                "Attendance": attendance_icon,
                "Actual": weekly_total,
                "Possible": weekly_possible,
                "%": round(weekly_percentage, 1),
                "Comment": (
                    str(row["Comment"])
                    .encode()
                    .decode("unicode_escape")
                    .replace("\n", "<br>")
                )
            }
        )

    percentage = 0

    if possible_total > 0:
        percentage = (
            actual_total
            / possible_total
        ) * 100

    if possible_total == 0:

        status = "No Assessments Recorded"

    else:

        status = get_categorical_grade(
            percentage
        )

    c1, c2 = st.columns(2)

    if possible_total == 0:
        percentage_display = "N/A"
    else:
        percentage_display = f"{percentage:.1f}%"

    with c1:
        st.metric(
            "Overall Percentage",
            percentage_display
        )

    with c2:
        st.metric(
            "Current Status",
            status
        )

    weekly_results_df = pd.DataFrame(
        weekly_results
    )

    if weekly_results_df.empty:

        st.info(
            "No assessment records found for this student."
        )

    else:

        weekly_results_df = (
            weekly_results_df
            .sort_values("Week")
        )

        weekly_results_df["Week"] = (
            weekly_results_df["Week"]
            .astype(str)
        )

        weekly_results_df["Actual"] = (
            weekly_results_df["Actual"]
            .astype(str)
        )

        weekly_results_df["Possible"] = (
            weekly_results_df["Possible"]
            .astype(str)
        )

        weekly_results_df["%"] = (
            weekly_results_df["%"]
            .round(1)
            .astype(str)
        )

        html_table = weekly_results_df.to_html(
            index=False,
            escape=False
        )

        html_table = html_table.replace(
            "<th>Comment</th>",
            '<th style="text-align:left;">Comment</th>'
        )

        html_table = html_table.replace(
            "<th>%</th>",
            '<th style="text-align:left;">%</th>'
        )

        html_code = f"""
        <div style="overflow-x:auto;">
        {html_table}
        </div>
        """

        st.markdown(
            html_code,
            unsafe_allow_html=True
        )

# --- PAGE 1: IMPORT ---
if page == "Setup Import":

    if not is_logged_in:
        st.error("🔒 Please select a role first.")

    elif st.session_state.get('role') not in ["Coordinator", "AssessmentTeam"]:
        st.error("🔒 Setup Import is restricted to Coordinators and Assessment Team.")

    else:

        st.header("1. Bulk Setup Import")

        c1, c2 = st.columns(2)

        with c1:
            st.subheader("Cohort Upload")
            u_files = st.file_uploader(
                "Upload all Group Excels",
                type="xlsx",
                accept_multiple_files=True
            )

            if u_files:
                all_dfs = []
                new_imgs = {}

                for f in u_files:

                    df_temp = pd.read_excel(f)
                    df_temp.columns = [str(c).lower().strip() for c in df_temp.columns]

                    df_temp = df_temp.rename(columns={
                        'id': 'ID',
                        'monash email': 'Email',
                        'tutor': 'Tutor',
                        'group': 'Group',
                        'preferred name': 'preferred name',
                        'family name': 'family name'
                    })

                    all_dfs.append(df_temp)

                    pxl = openpyxl.load_workbook(f, data_only=True)
                    loader = SheetImageLoader(pxl.active)

                    for r in range(2, pxl.active.max_row + 1):
                        email = str(
                            pxl.active.cell(row=r, column=5).value
                        ).strip().lower()

                        try:
                            img = loader.get(f'F{r}')
                            buf = io.BytesIO()
                            img.save(buf, format="PNG")
                            new_imgs[email] = buf.getvalue()
                        except:
                            continue

                st.session_state.student_data = pd.concat(
                    all_dfs,
                    ignore_index=True
                ).drop_duplicates(subset=['ID'])

                st.session_state.student_data.to_pickle(
                    "cohort.pkl"
                )

                st.session_state.image_library.update(new_imgs)
                pd.to_pickle(
                    st.session_state.image_library,
                    "images.pkl"
                )

                st.success(
                    f"Loaded {len(st.session_state.student_data)} students across {len(u_files)} groups."
                )

        with c2:
            st.subheader("Calendar Config")

            c_file = st.file_uploader(
                "Upload Calendar (Commencing | Week | Session Date)",
                type="xlsx"
            )

            if c_file:
                cdf = pd.read_excel(c_file)

                st.session_state.calendar_data = cdf.to_dict('records')

                cdf.to_pickle("calendar.pkl")

                st.success(
                    f"Calendar synced with {len(cdf)} weeks."
                )

        st.divider()

        st.subheader(
            "Assessment Configuration"
        )

        st.caption(
            "Configure categorical grade thresholds used throughout the assessment system."
        )

        thresholds = st.session_state.grade_thresholds

        commendation = st.number_input(
            "Commendation ≥ %",
            min_value=0,
            max_value=100,
            value=thresholds["Commendation"]
        )

        at_standard = st.number_input(
            "At Standard ≥ %",
            min_value=0,
            max_value=100,
            value=thresholds["At Standard"]
        )

        not_yet = st.number_input(
            "Not Yet At Standard ≥ %",
            min_value=0,
            max_value=100,
            value=thresholds["Not Yet At Standard"]
        )

        if st.button(
            "Save Thresholds"
        ):

            st.session_state.grade_thresholds = {
                "Commendation": commendation,
                "At Standard": at_standard,
                "Not Yet At Standard": not_yet
            }

            pd.to_pickle(
                st.session_state.grade_thresholds,
                "grade_thresholds.pkl"
            )

            st.success(
                "Thresholds saved successfully."
            )

        st.divider()
        st.subheader("Cohort Archive & Reset")
        st.caption("Download a backup of the current cohort data before clearing the system for a new cohort.")

        import zipfile

        pkl_files = [
            "cohort.pkl",
            "assessments.pkl",
            "calendar.pkl",
            "images.pkl",
            "grade_thresholds.pkl",
            "ui_state.pkl",
        ]
        existing_pkls = [f for f in pkl_files if os.path.exists(f)]

        if existing_pkls:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname in existing_pkls:
                    zf.write(fname)
            zip_buf.seek(0)

            stamp = datetime.now(pytz.timezone("Australia/Melbourne")).strftime("%Y%m%d_%H%M")
            st.download_button(
                label="📦 Download Full Cohort Backup (.zip)",
                data=zip_buf.getvalue(),
                file_name=f"pCLEAR_Backup_{stamp}.zip",
                mime="application/zip",
                key="download_cohort_backup_zip",
            )

        confirm_reset = st.checkbox(
            "I have downloaded a backup and want to clear current cohort, images, calendar, and assessments.",
            key="confirm_cohort_reset_chk",
        )

        if st.button("🗑️ Reset App for New Cohort", type="primary", disabled=not confirm_reset):
            for fname in ["cohort.pkl", "assessments.pkl", "calendar.pkl", "images.pkl", "ui_state.pkl"]:
                if os.path.exists(fname):
                    os.remove(fname)

            st.session_state.student_data = pd.DataFrame()
            st.session_state.assessments = pd.DataFrame()
            st.session_state.calendar_data = []
            st.session_state.image_library = {}
            st.session_state.saved_group = "All Groups"
            st.session_state.saved_period = "Semester 1"
            st.success("✅ System reset complete. Upload the new cohort and calendar above.")
            st.rerun()

        # =====================================================================
        # 🛑 DANGER: DUMMY DATA GENERATOR - REMOVE BEFORE PRODUCTION DEPLOYMENT
        # =====================================================================
        st.divider()
        st.warning("⚠️ **DEVELOPMENT ONLY:** Dummy Data Generator Active. Remove before deployment.")
        st.subheader("🛠️ Developer Tools")
        
        if st.button("Generate Dummy Assessment Data"):
            import random
            
            if "student_data" not in st.session_state or st.session_state.student_data.empty:
                st.error("Please upload cohort data first so I know who to grade!")
            else:
                dummy_records = []
                cohort = st.session_state.student_data.copy()
                
                if "preferred name" not in cohort.columns:
                    cohort["preferred name"] = ""
                if "family name" not in cohort.columns:
                    cohort["family name"] = ""
                    
                cohort["Full Name"] = (cohort["preferred name"].fillna("") + " " + cohort["family name"].fillna("")).str.strip()
                
                # Find the ID column in student_data (e.g., 'ID', 'Student ID', or 'id')
                id_col = next((c for c in cohort.columns if c.strip().lower() in ["id", "student id"]), None)
                
                # Deduplicate students while keeping their ID
                valid_cohort = cohort[cohort["Full Name"] != ""].drop_duplicates(subset=["Full Name"])
                
                if valid_cohort.empty:
                    st.error("No valid student names found in the cohort data. Cannot generate assessments.")
                else:
                    grading_columns = [
                        "Engagement", "Teamwork", "Written Tasks", 
                        "Patience/Respect", "Accountability", 
                        "Uncertainty", "Punctuality"
                    ]
                    
                    for idx, row in valid_cohort.iterrows():
                        student_name = row["Full Name"]
                        student_id = row[id_col] if id_col else f"DUMMY_{idx}"
                        
                        for week in range(1, 29):
                            record = {
                                "ID": student_id,
                                "Name": student_name,
                                "Week": week,
                                "Attendance": random.choice(["Present", "Present", "Present", "Absent - notified"]),
                                "Comment": "Dummy comment" if random.random() > 0.8 else ""
                            }
                            for col in grading_columns:
                                record[col] = random.randint(2, 5)
                                
                            dummy_records.append(record)
                    
                    st.session_state.assessments = pd.DataFrame(dummy_records)
                    st.session_state.assessments.to_pickle("assessments.pkl")
                    
                    st.success(f"✅ Generated 28 weeks of dummy data for {len(valid_cohort)} students! Reloading...")
                    st.rerun()
        # =====================================================================

# --- PAGE 2: Assessment Analytics (LOCKED) ---
elif page == "Assessment Analytics":
    st.header("Assessment Analytics")

    col_period, col_label, col_group, col_reset = st.columns([3.2, 1.1, 1.8, 1.2])

    with col_period:
        period_options = ["Semester 1", "Semester 2", "Full Year"]
        reporting_period = st.radio(
            "**Reporting Period**",
            period_options,
            index=period_options.index(st.session_state.saved_period),
            horizontal=True,
            key="analytics_period_radio",
            on_change=save_period,
            args=("analytics_period_radio",)
        )

    with col_label:
        st.markdown(
            "<div style='padding-top: 8px; text-align: right; font-weight: 600;'>Filter by Group</div>",
            unsafe_allow_html=True
        )

    with col_group:
        if not st.session_state.student_data.empty and "Group" in st.session_state.student_data.columns:
            unique_groups = sorted(
                st.session_state.student_data["Group"].dropna().unique().tolist(),
                key=lambda x: int(x) if str(x).isdigit() else str(x)
            )
            group_options = ["All Groups"] + unique_groups
        else:
            group_options = ["All Groups"]

        if st.session_state.saved_group not in group_options:
            st.session_state.saved_group = "All Groups"

        selected_group = st.selectbox(
            "Filter by Group",
            group_options,
            index=group_options.index(st.session_state.saved_group),
            label_visibility="collapsed",
            key="analytics_group_select",
            on_change=save_group,
            args=("analytics_group_select",)
        )

    with col_reset:
        if st.button("🔄 Reset Filters", use_container_width=True):
            st.session_state.saved_period = "Full Year"
            st.session_state.saved_group = "All Groups"
            st.session_state.saved_sort_by = "Last Name (A → Z)"
            st.session_state.saved_view_mode = "Full Detail (All Weeks)"
            st.session_state.pop("matrix_sort_by", None)
            st.session_state.pop("matrix_view_mode", None)
            save_ui_state()
            st.rerun()

    analytics_tabs = ["Achievement Matrix", "Mail Merge", "Validation"]
    if st.session_state.get("saved_analytics_tab") not in analytics_tabs:
        st.session_state.saved_analytics_tab = analytics_tabs[0]

    st.markdown(
        """
        <style>
        div[data-testid="stRadio"] > div[role="radiogroup"] {
            gap: 1.5rem;
            border-bottom: 1px solid #E6EAF1;
            padding-bottom: 0.25rem;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label > div:first-child {
            display: none;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label {
            padding: 0.35rem 0.25rem;
            border-bottom: 2px solid transparent;
            cursor: pointer;
        }
        div[data-testid="stRadio"] > div[role="radiogroup"] label:has(input:checked) {
            border-bottom: 2px solid #FF4B4B;
            color: #FF4B4B;
            font-weight: 600;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    active_tab = st.radio(
        "Analytics Section",
        analytics_tabs,
        index=analytics_tabs.index(st.session_state.saved_analytics_tab),
        horizontal=True,
        key="analytics_tab_selector",
        on_change=save_analytics_tab,
        args=("analytics_tab_selector",),
        label_visibility="collapsed",
    )

    assessment_df = (
        st.session_state.assessments.copy()
    )

    if active_tab == "Achievement Matrix":

        st.subheader("Achievement Matrix")

        # -------------------------
        # Filters
        # -------------------------

        col_search, col_sort, col_view = st.columns([2, 1.5, 1.5])

        with col_search:
            student_search = st.text_input(
                "Find Student",
                placeholder="Type a student's name..."
            )

        with col_sort:
            sort_options = ["Last Name (A → Z)", "First Name (A → Z)", "Lowest Overall % First"]
            if st.session_state.saved_sort_by not in sort_options:
                st.session_state.saved_sort_by = "Last Name (A → Z)"
            sort_by = st.selectbox(
                "Sort Students By",
                sort_options,
                index=sort_options.index(st.session_state.saved_sort_by),
                key="matrix_sort_by",
                on_change=save_sort_by,
                args=("matrix_sort_by",)
            )

        with col_view:
            view_options = ["Full Detail (All Weeks)", "Summary Rows Only"]
            if st.session_state.saved_view_mode not in view_options:
                st.session_state.saved_view_mode = "Full Detail (All Weeks)"
            view_mode = st.selectbox(
                "Row Detail",
                view_options,
                index=view_options.index(st.session_state.saved_view_mode),
                key="matrix_view_mode",
                on_change=save_view_mode,
                args=("matrix_view_mode",)
            )

        # Ensure columns exist even if no assessments have been recorded yet
        if "Week" not in assessment_df.columns:
            assessment_df["Week"] = pd.Series(dtype="float")
        if "Name" not in assessment_df.columns:
            assessment_df["Name"] = pd.Series(dtype="str")

        assessment_df["Week"] = pd.to_numeric(
            assessment_df["Week"],
            errors="coerce"
        )

        matrix_rows = []

        cohort_df = st.session_state.student_data.copy()

        # Safely create placeholder columns if the database is empty
        if "preferred name" not in cohort_df.columns:
            cohort_df["preferred name"] = pd.Series(dtype="str")
        if "family name" not in cohort_df.columns:
            cohort_df["family name"] = pd.Series(dtype="str")

        cohort_df["Full Name"] = (
            cohort_df["preferred name"].fillna("")
            + " "
            + cohort_df["family name"].fillna("")
        ).str.strip()

        if selected_group != "All Groups":
            cohort_df = cohort_df[
                cohort_df["Group"] == selected_group
            ]

        unique_students = cohort_df["Full Name"].dropna().unique().tolist()

        if student_search:
            unique_students = [
                s for s in unique_students
                if student_search.lower() in s.lower()
            ]

        if sort_by == "First Name (A → Z)":
            students = sorted(unique_students, key=lambda x: x.lower())
        elif sort_by == "Lowest Overall % First":
            students = sorted(
                unique_students,
                key=lambda x: (get_student_overall_pct(x, assessment_df, reporting_period), x.split()[-1].lower())
            )
        else:  # "Last Name (A → Z)"
            students = sorted(unique_students, key=lambda x: x.split()[-1].lower())

        sem1_start, sem1_end = min(SEMESTER_1_WEEKS), max(SEMESTER_1_WEEKS)
        sem2_start, sem2_end = min(SEMESTER_2_WEEKS), max(SEMESTER_2_WEEKS)

        for student in students:

            student_df = (
                assessment_df[
                    assessment_df["Name"] == student
                ]
                .sort_values("Week")
            )

            matrix_rows.append({
                "Student": student.replace(" ", "\n", 1),
                "Row Type": "Student Header",
                "Engagement": "",
                "Teamwork": "",
                "Written Tasks": "",
                "Patience/Respect": "",
                "Accountability": "",
                "Uncertainty": "",
                "Punctuality": "",
                "Total Score": "",
            })

            # Always calculate totals so Full Year has both semesters ready
            semester_totals = {}
            semester_possible = {}
            for domain in SCORE_COLUMNS:
                actual, possible = calculate_domain_totals(
                    student_df, domain, sem1_start, sem1_end
                )
                semester_totals[domain] = actual
                semester_possible[domain] = possible

            semester2_totals = {}
            semester2_possible = {}
            for domain in SCORE_COLUMNS:
                actual, possible = calculate_domain_totals(
                    student_df, domain, sem2_start, sem2_end
                )
                semester2_totals[domain] = actual
                semester2_possible[domain] = possible

            # ----------------------------------
            # SEMESTER 1 ROWS
            # ----------------------------------
            if reporting_period in ["Semester 1", "Full Year"]:
                for week in SEMESTER_1_WEEKS:
                    week_record = student_df[student_df["Week"] == week]
                    row = {
                        "Student": "",
                        "Row Type": f"Week {week}"
                    }
                    for domain in SCORE_COLUMNS:
                        row[domain] = week_record.iloc[0][domain] if not week_record.empty else ""
                    row["Total Score"] = format_weekly_total_cell(week_record)
                    matrix_rows.append(row)

                subtotal_row = {"Student": "", "Row Type": "SEM 1 - Student Subtotal"}
                percentage_row = {"Student": "", "Row Type": "SEM 1 - Student Percentage"}
                grade_row = {"Student": "", "Row Type": "SEM 1 - Categorical Grade"}

                for domain in SCORE_COLUMNS:
                    subtotal_row[domain] = semester_totals[domain]
                    possible = semester_possible[domain]
                    if possible > 0:
                        pct = (semester_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                s1_tot_act = sum(semester_totals.values())
                s1_tot_poss = sum(semester_possible.values())
                subtotal_row["Total Score"] = s1_tot_act
                if s1_tot_poss > 0:
                    s1_tot_pct = (s1_tot_act / s1_tot_poss) * 100
                    percentage_row["Total Score"] = f"{s1_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(s1_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                matrix_rows.extend([subtotal_row, percentage_row, grade_row])

                if reporting_period == "Semester 1":
                    sem1_max_row = {"Student": "", "Row Type": "SEM 1 - Maximum Possible"}
                    for domain in SCORE_COLUMNS:
                        sem1_max_row[domain] = semester_possible[domain]
                    sem1_max_row["Total Score"] = s1_tot_poss
                    matrix_rows.append(sem1_max_row)

            # ----------------------------------
            # SEMESTER 2 ROWS
            # ----------------------------------
            if reporting_period in ["Semester 2", "Full Year"]:
                for week in SEMESTER_2_WEEKS:
                    week_record = student_df[student_df["Week"] == week]
                    row = {
                        "Student": "",
                        "Row Type": f"Week {week}"
                    }
                    for domain in SCORE_COLUMNS:
                        row[domain] = week_record.iloc[0][domain] if not week_record.empty else ""
                    row["Total Score"] = format_weekly_total_cell(week_record)
                    matrix_rows.append(row)

                subtotal_row = {"Student": "", "Row Type": "SEM 2 - Student Subtotal"}
                percentage_row = {"Student": "", "Row Type": "SEM 2 - Student Percentage"}
                grade_row = {"Student": "", "Row Type": "SEM 2 - Categorical Grade"}

                for domain in SCORE_COLUMNS:
                    subtotal_row[domain] = semester2_totals[domain]
                    possible = semester2_possible[domain]
                    if possible > 0:
                        pct = (semester2_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                s2_tot_act = sum(semester2_totals.values())
                s2_tot_poss = sum(semester2_possible.values())
                subtotal_row["Total Score"] = s2_tot_act
                if s2_tot_poss > 0:
                    s2_tot_pct = (s2_tot_act / s2_tot_poss) * 100
                    percentage_row["Total Score"] = f"{s2_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(s2_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                matrix_rows.extend([subtotal_row, percentage_row, grade_row])

                if reporting_period == "Semester 2":
                    sem2_max_row = {"Student": "", "Row Type": "SEM 2 - Maximum Possible"}
                    for domain in SCORE_COLUMNS:
                        sem2_max_row[domain] = semester2_possible[domain]
                    sem2_max_row["Total Score"] = s2_tot_poss
                    matrix_rows.append(sem2_max_row)

            # ----------------------------------
            # FULL YEAR SUMMARY & MAXIMUMS
            # ----------------------------------
            if reporting_period == "Full Year":
                full_year_totals = {}
                full_year_possible = {}

                for domain in SCORE_COLUMNS:
                    full_year_totals[domain] = semester_totals[domain] + semester2_totals[domain]
                    full_year_possible[domain] = semester_possible[domain] + semester2_possible[domain]

                full_year_total_row = {"Student": "", "Row Type": "Full Year - Student Total"}
                percentage_row = {"Student": "", "Row Type": "Full Year - Student Percentage"}
                grade_row = {"Student": "", "Row Type": "Full Year - Categorical Grade"}

                for domain in SCORE_COLUMNS:
                    full_year_total_row[domain] = full_year_totals[domain]
                    possible = full_year_possible[domain]
                    if possible > 0:
                        pct = (full_year_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                fy_tot_act = sum(full_year_totals.values())
                fy_tot_poss = sum(full_year_possible.values())
                full_year_total_row["Total Score"] = fy_tot_act
                if fy_tot_poss > 0:
                    fy_tot_pct = (fy_tot_act / fy_tot_poss) * 100
                    percentage_row["Total Score"] = f"{fy_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(fy_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                matrix_rows.extend([full_year_total_row, percentage_row, grade_row])

                sem1_max_row = {"Student": "", "Row Type": "SEM 1 - Maximum Possible"}
                sem2_max_row = {"Student": "", "Row Type": "SEM 2 - Maximum Possible"}
                full_year_max_row = {"Student": "", "Row Type": "Full Year - Maximum Possible"}

                for domain in SCORE_COLUMNS:
                    sem1_max_row[domain] = semester_possible[domain]
                    sem2_max_row[domain] = semester2_possible[domain]
                    full_year_max_row[domain] = full_year_possible[domain]

                sem1_max_row["Total Score"] = sum(semester_possible.values())
                sem2_max_row["Total Score"] = sum(semester2_possible.values())
                full_year_max_row["Total Score"] = fy_tot_poss

                matrix_rows.extend([sem1_max_row, sem2_max_row, full_year_max_row])

        matrix_rows.append({
            "Student": "END OF REPORT",
            "Row Type": "Student Header",
            "Engagement": "",
            "Teamwork": "",
            "Written Tasks": "",
            "Patience/Respect": "",
            "Accountability": "",
            "Uncertainty": "",
            "Punctuality": "",
            "Total Score": "",
        })
        
        # 1. Create the DataFrame ONCE
        if view_mode == "Summary Rows Only":
            matrix_rows = [
                r for r in matrix_rows
                if not str(r.get("Row Type", "")).startswith("Week ")
            ]
        matrix_df = pd.DataFrame(matrix_rows)

        # 2. Safely cast everything to strings so PyArrow doesn't crash
        for col in SCORE_COLUMNS + ["Total Score"]:
            matrix_df[col] = matrix_df[col].fillna("").astype(str)

        matrix_df["Student"] = matrix_df["Student"].fillna("").astype(str)
        matrix_df["Row Type"] = matrix_df["Row Type"].fillna("").astype(str)
        
        # 3. Build the Grid Options cleanly
        gb = GridOptionsBuilder.from_dataframe(matrix_df)

        gb.configure_default_column(
            resizable=True,
            sortable=False,
            filter=False,
            wrapHeaderText=True,
            autoHeaderHeight=True
        )

        # Configure Student column with clean 2-line wrapping (no mid-word splits)
        gb.configure_column(
            "Student",
            width=165,
            pinned="left",
            wrapText=True,
            autoHeight=True,
            cellStyle={
                "white-space": "pre-line",
                "word-break": "normal",
                "overflow-wrap": "normal",
                "line-height": "1.35",
                "padding-top": "6px",
                "padding-bottom": "6px"
            }
        )
        
        # Hide the literal text "Student Header" while keeping the value for blue row styling
        gb.configure_column(
            "Row Type", 
            header_name="", 
            width=235, 
            pinned="left",
            wrapText=True,
            autoHeight=True,
            valueFormatter=JsCode("""
                function(params) {
                    return params.value === 'Student Header' ? '' : params.value;
                }
            """)
        )
        
        gb.configure_column("Engagement")
        gb.configure_column("Teamwork")
        gb.configure_column("Written Tasks")
        gb.configure_column("Patience/Respect", header_name="Patience/ Respect")
        gb.configure_column("Accountability")
        gb.configure_column("Uncertainty")
        gb.configure_column("Punctuality")
        gb.configure_column(
            "Total Score",
            header_name="Total Score",
            cellStyle={"font-weight": "700", "background-color": "#F8F9FA"},
        )

        grid_options = gb.build()

        col_widths = {
            "Student": 130,
            "Row Type": 235,
            "Engagement": 134,
            "Teamwork": 120,
            "Written Tasks": 105,
            "Patience/Respect": 115,
            "Accountability": 148,
            "Uncertainty": 130,
            "Punctuality": 130,
            "Total Score": 115,
        }

        grade_tooltip_js = JsCode("""
            function(params) {
                const map = {
                    'C': 'Commendation',
                    'AS': 'At Standard',
                    'NY': 'Not Yet At Standard',
                    'CU': 'Concerning / Unacceptable Behaviour'
                };
                return map[params.value] || null;
            }
        """)

        for col_def in grid_options["columnDefs"]:
            field = col_def.get("field")
            if field in ["Student", "Row Type"]:
                col_def["pinned"] = "left"
            if field in col_widths:
                col_def["width"] = col_widths[field]
                col_def["minWidth"] = col_widths[field]
            if field in SCORE_COLUMNS + ["Total Score"]:
                col_def["tooltipValueGetter"] = grade_tooltip_js

        grid_options["tooltipShowDelay"] = 200

        # 4. Apply the styling rule for the headers, categorical grade rows, and END OF REPORT bar
        grid_options["rowClassRules"] = {
            "student-header": JsCode("""
                function(params) {
                    return params.data['Row Type'] === 'Student Header';
                }
            """),
            "categorical-grade-row": JsCode("""
                function(params) {
                    return params.data['Row Type'] && params.data['Row Type'].endsWith('Categorical Grade');
                }
            """)
        }

        # 5. Render directly (no outer container clipping) with a dynamic key so filters update immediately
        response = AgGrid(
            matrix_df,
            gridOptions=grid_options,
            allow_unsafe_jscode=True,
            custom_css={
                ".ag-row-odd, .ag-row-even": {
                    "background-color": "#FFFFFF !important"
                },
                ".student-header": {
                    "background-color": "#D9EAF7 !important",
                    "font-weight": "bold !important"
                },
                ".categorical-grade-row": {
                    "background-color": "#FEF9E7 !important",
                    "font-weight": "600 !important",
                    "border-top": "1px solid #F3E8C8 !important",
                    "border-bottom": "1px solid #E6D8B5 !important"
                },
                ".ag-body-viewport": {
                    "padding-bottom": "35px !important"
                }
            },
            height=800,
            theme="alpine",
            key=f"assessment_matrix_grid_{selected_group}_{student_search}_{reporting_period}_{sort_by}_{view_mode}"
        )

        matrix_excel_data = build_template_matrix_excel(
            st.session_state.student_data,
            assessment_df,
            selected_group,
        )
        period_slug = str(st.session_state.saved_period).replace(" ", "_")
        group_slug = str(st.session_state.saved_group).replace(" ", "_")

        st.download_button(
            label="📥 Download Achievement Matrix (Excel)",
            data=matrix_excel_data,
            file_name=f"Achievement_Matrix_{group_slug}_{period_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_matrix_excel",
        )

    elif active_tab == "Mail Merge":
        st.subheader("Mail Merge Export")
        st.caption(f"Showing **{reporting_period}** categorical grades for **{selected_group}**.")

        if st.session_state.student_data.empty:
            st.info("No student cohort data loaded yet.")
        else:
            mm_cohort = st.session_state.student_data.copy()

            # Detect the Student ID column name safely
            id_col = next(
                (c for c in mm_cohort.columns if c.strip().lower() in ["student id", "id", "student_id", "studentid"]),
                None
            )

            if "preferred name" not in mm_cohort.columns:
                mm_cohort["preferred name"] = ""
            if "family name" not in mm_cohort.columns:
                mm_cohort["family name"] = ""

            if selected_group != "All Groups" and "Group" in mm_cohort.columns:
                mm_cohort = mm_cohort[mm_cohort["Group"] == selected_group]

            # Sort alphabetically by surname then first name
            mm_cohort = mm_cohort.sort_values(
                by=["family name", "preferred name"],
                key=lambda s: s.fillna("").str.lower()
            )

            # Determine active week range from top Reporting Period selector
            if reporting_period == "Semester 1":
                mm_start, mm_end = min(SEMESTER_1_WEEKS), max(SEMESTER_1_WEEKS)
            elif reporting_period == "Semester 2":
                mm_start, mm_end = min(SEMESTER_2_WEEKS), max(SEMESTER_2_WEEKS)
            else:
                mm_start, mm_end = min(SEMESTER_1_WEEKS), max(SEMESTER_2_WEEKS)

            # Exact mapping required by the Mail Merge Excel template
            DOMAIN_TO_MAIL_MERGE_COL = {
                "Engagement": "Engagement - Categorical",
                "Teamwork": "Teamwork - Categorical",
                "Written Tasks": "Written Tesks - Categorical",
                "Patience/Respect": "Patience, Courtesy, respect and reflection - Categorical",
                "Accountability": "Accountability - Categorical",
                "Uncertainty": "Dealing with learning ambiguity - Categorical",
                "Punctuality": "Punctuality - Categorical",
            }

            # Clean display headers for the UI preview
            MAIL_MERGE_TO_DISPLAY_COL = {
                "Student ID": "Student ID",
                "Student Firstname": "First Name",
                "Student Surname": "Surname",
                "Engagement - Categorical": "Engagement",
                "Teamwork - Categorical": "Teamwork",
                "Written Tesks - Categorical": "Written Tasks",
                "Patience, Courtesy, respect and reflection - Categorical": "Patience/Respect",
                "Accountability - Categorical": "Accountability",
                "Dealing with learning ambiguity - Categorical": "Uncertainty",
                "Punctuality - Categorical": "Punctuality",
                "Total - Categorical": "Total",
            }

            mm_rows = []
            for _, stu in mm_cohort.iterrows():
                first_name = str(stu.get("preferred name", "") or "").strip()
                last_name = str(stu.get("family name", "") or "").strip()
                full_name = f"{first_name} {last_name}".strip()
                student_id = str(stu.get(id_col, "") or "").strip() if id_col else ""

                s_df = assessment_df[assessment_df["Name"] == full_name]

                row_out = {
                    "Student ID": student_id,
                    "Student Firstname": first_name,
                    "Student Surname": last_name,
                }

                overall_actual = 0
                overall_possible = 0

                for domain, target_col in DOMAIN_TO_MAIL_MERGE_COL.items():
                    actual, possible = calculate_domain_totals(s_df, domain, mm_start, mm_end)
                    overall_actual += actual
                    overall_possible += possible
                    if possible > 0:
                        pct = (actual / possible) * 100
                        row_out[target_col] = get_categorical_grade(pct)
                    else:
                        row_out[target_col] = ""

                if overall_possible > 0:
                    overall_pct = (overall_actual / overall_possible) * 100
                    row_out["Total - Categorical"] = get_categorical_grade(overall_pct)
                else:
                    row_out["Total - Categorical"] = ""

                mm_rows.append(row_out)

            # DataFrame with exact Mail Merge headers (used for Excel export)
            mm_export_df = pd.DataFrame(mm_rows)

            # DataFrame with clean UI headers (used for on-screen preview)
            mm_display_df = mm_export_df.rename(columns=MAIL_MERGE_TO_DISPLAY_COL)

            # Build AgGrid with a spanned "Categorical Grades" top header group
            mm_gb = GridOptionsBuilder.from_dataframe(mm_display_df)
            mm_gb.configure_default_column(
                resizable=True,
                sortable=True,
                filter=False,
                wrapHeaderText=True,
                autoHeaderHeight=True
            )

            mm_grid_options = mm_gb.build()
            mm_grid_options["columnDefs"] = [
                {
                    "headerName": "Student Details",
                    "children": [
                        {"field": "Student ID", "width": 125, "pinned": "left"},
                        {"field": "First Name", "width": 135, "pinned": "left"},
                        {"field": "Surname", "width": 140, "pinned": "left"},
                    ]
                },
                {
                    "headerName": "Categorical Grades",
                    "children": [
                        {"field": "Engagement", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Teamwork", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Written Tasks", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Patience/Respect", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Accountability", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Uncertainty", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Punctuality", "width": 165, "wrapText": True, "autoHeight": True},
                        {"field": "Total", "width": 165, "wrapText": True, "autoHeight": True},
                    ]
                }
            ]

            AgGrid(
                mm_display_df,
                gridOptions=mm_grid_options,
                allow_unsafe_jscode=True,
                custom_css={
                    ".ag-row-odd, .ag-row-even": {
                        "background-color": "#FFFFFF !important"
                    },
                    ".ag-header-group-cell-label": {
                        "font-weight": "700 !important",
                        "justify-content": "flex-start !important",
                        "padding-left": "12px !important"
                    },
                    ".ag-body-viewport": {
                        "padding-bottom": "35px !important"
                    }
                },
                height=800,
                theme="alpine",
                key=f"mail_merge_grid_{selected_group}_{reporting_period}"
            )

            # Build formatted Excel (.xlsx) preserving exact Mail Merge template headers
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                mm_export_df.to_excel(writer, index=False, sheet_name="Mail Merge")
                ws = writer.sheets["Mail Merge"]

                from openpyxl.styles import Font, Alignment, Border, Side

                header_font = Font(bold=True)
                header_align = Alignment(wrap_text=True, vertical="bottom", horizontal="left")
                thin_side = Side(style="thin", color="D9D9D9")
                cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                ws.row_dimensions[1].height = 42
                for col_idx, col_name in enumerate(mm_export_df.columns, start=1):
                    cell = ws.cell(row=1, column=col_idx)
                    cell.font = header_font
                    cell.alignment = header_align
                    cell.border = cell_border

                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = 16 if col_idx <= 3 else 30

                for r_idx in range(2, len(mm_export_df) + 2):
                    for c_idx in range(1, len(mm_export_df.columns) + 1):
                        ws.cell(row=r_idx, column=c_idx).border = cell_border

            safe_period = reporting_period.replace(" ", "_")
            safe_group = str(selected_group).replace(" ", "_")
            file_name = f"Mail_Merge_{safe_period}_{safe_group}.xlsx"

            st.download_button(
                label="📥 Download Mail Merge Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=False
            )

    elif active_tab == "Validation":

        st.subheader("Assessment Data Validation & Matrix Audit")

        # -------------------------
        # Filters & Quick Toggles
        # -------------------------
        col_search, col_sort, col_view, col_flt = st.columns([1.8, 1.4, 1.4, 1.6])

        current_stu_search = st.session_state.saved_val_student or ""
        if current_stu_search == "All Students":
            current_stu_search = ""

        with col_search:
            val_student_search = st.text_input(
                "Find Student",
                value=current_stu_search,
                placeholder="Type a student's name...",
                key="val_student_widget",
                on_change=save_val_student,
                args=("val_student_widget",),
            )

        with col_sort:
            sort_options = ["Last Name (A → Z)", "First Name (A → Z)", "Most Issues First"]
            if st.session_state.get("saved_val_sort") not in sort_options:
                st.session_state.saved_val_sort = "Last Name (A → Z)"
            val_sort_by = st.selectbox(
                "Sort Students By",
                sort_options,
                index=sort_options.index(st.session_state.saved_val_sort),
                key="val_matrix_sort_by",
            )
            st.session_state.saved_val_sort = val_sort_by

        with col_view:
            val_view_options = ["Full Detail (All Weeks)", "Summary Rows Only"]
            if st.session_state.get("saved_val_view_mode") not in val_view_options:
                st.session_state.saved_val_view_mode = "Full Detail (All Weeks)"
            val_view_mode = st.selectbox(
                "Row Detail",
                val_view_options,
                index=val_view_options.index(st.session_state.saved_val_view_mode),
                key="val_matrix_view_mode",
            )
            st.session_state.saved_val_view_mode = val_view_mode

        with col_flt:
            val_filter_options = [
                "All Issues (Missing + Flagged)",
                "All Students",
                "🚩 Flagged Only (Scores / Comments)",
                "⚠️ Missing / Duplicate Weeks Only",
                "🧮 Math Mismatches Only",
                "✅ Verified Only",
            ]
            current_flt = st.session_state.get("saved_val_filter", "All Issues (Missing + Flagged)")
            if current_flt not in val_filter_options:
                current_flt = "All Issues (Missing + Flagged)"
            val_filter = st.selectbox(
                "Validation Status",
                val_filter_options,
                index=val_filter_options.index(current_flt),
                key="val_filter_widget",
                on_change=save_val_filter,
                args=("val_filter_widget",),
            )

        # Ensure columns exist safely
        if "Week" not in assessment_df.columns:
            assessment_df["Week"] = pd.Series(dtype="float")
        if "Name" not in assessment_df.columns:
            assessment_df["Name"] = pd.Series(dtype="str")

        assessment_df["Week"] = pd.to_numeric(assessment_df["Week"], errors="coerce")

        # Identify non-teaching weeks & max logged week so future weeks aren't falsely flagged
        no_class_weeks = set()
        if "calendar_data" in st.session_state and st.session_state.calendar_data:
            for cal_row in st.session_state.calendar_data:
                if pd.notna(cal_row.get("Week")) and pd.isna(cal_row.get("date")):
                    try:
                        no_class_weeks.add(int(float(cal_row["Week"])))
                    except (ValueError, TypeError):
                        pass

        logged_weeks = assessment_df["Week"].dropna()
        max_logged_week = int(logged_weeks.max()) if not logged_weeks.empty else 0

        cohort_df = st.session_state.student_data.copy()
        if "preferred name" not in cohort_df.columns:
            cohort_df["preferred name"] = pd.Series(dtype="str")
        if "family name" not in cohort_df.columns:
            cohort_df["family name"] = pd.Series(dtype="str")

        cohort_df["Full Name"] = (
            cohort_df["preferred name"].fillna("")
            + " "
            + cohort_df["family name"].fillna("")
        ).str.strip()

        # Check for orphan assessment rows whose IDs are not in the roster
        all_cohort_ids = set(cohort_df["ID"].astype(str).str.strip().unique()) if "ID" in cohort_df.columns else set()
        if "ID" in assessment_df.columns and all_cohort_ids:
            scoped_assess = filter_by_period(assessment_df.copy(), reporting_period, week_col="Week")
            orphan_rows = scoped_assess[~scoped_assess["ID"].astype(str).str.strip().isin(all_cohort_ids)]
        else:
            orphan_rows = pd.DataFrame()

        if selected_group != "All Groups":
            cohort_df = cohort_df[cohort_df["Group"] == selected_group]

        student_meta = {}
        for _, s_row in cohort_df.iterrows():
            f_name = str(s_row.get("Full Name", "")).strip()
            if f_name:
                student_meta[f_name] = {
                    "id": str(s_row.get("ID", "")).strip(),
                    "group": str(s_row.get("Group", "")).strip(),
                }

        unique_students = cohort_df["Full Name"].dropna().unique().tolist()
        if val_student_search.strip():
            q = val_student_search.strip().lower()
            unique_students = [
                s for s in unique_students
                if q in s.lower() or q in student_meta.get(s, {}).get("id", "").lower()
            ]

        sem1_start, sem1_end = min(SEMESTER_1_WEEKS), max(SEMESTER_1_WEEKS)
        sem2_start, sem2_end = min(SEMESTER_2_WEEKS), max(SEMESTER_2_WEEKS)

        def _audit_week_row(week_num, week_record):
            cell_flags = {}
            row_issues = []

            if week_num in no_class_weeks or week_num > max_logged_week:
                return cell_flags, row_issues

            if week_record.empty:
                row_issues.append(f"Week {week_num}: No assessment logged")
                for d in SCORE_COLUMNS:
                    cell_flags[d] = {"level": "warn", "msg": f"Week {week_num}: Missing record"}
                return cell_flags, row_issues

            if len(week_record) > 1:
                row_issues.append(f"Week {week_num}: Duplicate entries ({len(week_record)})")

            r = week_record.iloc[0]
            att = str(r.get("Attendance", "Present")).strip()
            if att in ("", "nan", "None"):
                att = "Present"
            is_absent = att.lower().startswith("absent")

            note = str(r.get("Comment", "")).strip() if pd.notna(r.get("Comment")) else ""
            if note.lower() in ("nan", "none"):
                note = ""
            has_comment = len(note) >= 5

            for d in SCORE_COLUMNS:
                raw = str(r.get(d, "")).strip()
                is_blank = raw in ("", "nan", "None", "---", "N/A")

                if is_absent:
                    if d == "Accountability":
                        if is_blank:
                            msg = f"Missing Accountability score ({att})"
                            cell_flags[d] = {"level": "error", "msg": msg}
                            row_issues.append(msg)
                        elif att == "Absent - unnotified" and not has_comment:
                            msg = "Absent - unnotified requires facilitator comment"
                            cell_flags[d] = {"level": "error", "msg": msg}
                            row_issues.append(msg)
                    else:
                        if not is_blank:
                            msg = f"{d} has score '{raw}' while marked {att}"
                            cell_flags[d] = {"level": "error", "msg": msg}
                            row_issues.append(msg)
                else:
                    if raw.upper() in ("N/A", "I/O"):
                        continue
                    if is_blank:
                        msg = f"Missing {d} score (Present)"
                        cell_flags[d] = {"level": "error", "msg": msg}
                        row_issues.append(msg)
                    else:
                        try:
                            val = float(raw)
                            if val < 0 or val > 5 or not val.is_integer():
                                msg = f"Invalid {d} score ({raw})"
                                cell_flags[d] = {"level": "error", "msg": msg}
                                row_issues.append(msg)
                            elif int(val) in (0, 5) and not has_comment:
                                msg = f"{d} = {int(val)} requires comment"
                                cell_flags[d] = {"level": "error", "msg": msg}
                                row_issues.append(msg)
                        except ValueError:
                            msg = f"Non-numeric {d} score ({raw})"
                            cell_flags[d] = {"level": "error", "msg": msg}
                            row_issues.append(msg)

            return cell_flags, row_issues

        def _verify_domain_period_math(s_df, wk_list, reported_act, reported_poss, label):
            math_flags = {}
            sub_df = s_df[s_df["Week"].isin(wk_list)]
            for d in SCORE_COLUMNS:
                exp_act = 0
                exp_poss = 0
                for _, wr in sub_df.iterrows():
                    raw = str(wr.get(d, "")).strip()
                    if raw not in ("", "nan", "None", "---", "N/A"):
                        try:
                            v = float(raw)
                            if 0 <= v <= 5:
                                exp_act += int(v)
                                exp_poss += 5
                        except ValueError:
                            pass
                if exp_act != reported_act[d]:
                    math_flags[f"act_{d}"] = {
                        "level": "math",
                        "msg": f"{label} {d} Subtotal mismatch: matrix shows {reported_act[d]}, weekly sum is {exp_act}",
                    }
                if exp_poss != reported_poss[d]:
                    math_flags[f"poss_{d}"] = {
                        "level": "math",
                        "msg": f"{label} {d} Max Possible mismatch: matrix shows {reported_poss[d]}, expected {exp_poss}",
                    }
            return math_flags

        # Build student matrix blocks first so we can sort by issue count if requested
        student_blocks = []
        total_flagged_students = 0
        total_cell_issues = 0
        export_issue_rows = []

        for student in unique_students:
            student_df = (
                assessment_df[assessment_df["Name"] == student]
                .sort_values("Week")
            )
            stu_block_rows = []
            stu_issues = []

            semester_totals = {}
            semester_possible = {}
            for domain in SCORE_COLUMNS:
                actual, possible = calculate_domain_totals(student_df, domain, sem1_start, sem1_end)
                semester_totals[domain] = actual
                semester_possible[domain] = possible

            semester2_totals = {}
            semester2_possible = {}
            for domain in SCORE_COLUMNS:
                actual, possible = calculate_domain_totals(student_df, domain, sem2_start, sem2_end)
                semester2_totals[domain] = actual
                semester2_possible[domain] = possible

            sem1_math = _verify_domain_period_math(student_df, SEMESTER_1_WEEKS, semester_totals, semester_possible, "SEM 1")
            sem2_math = _verify_domain_period_math(student_df, SEMESTER_2_WEEKS, semester2_totals, semester2_possible, "SEM 2")

            # SEMESTER 1 ROWS
            if reporting_period in ["Semester 1", "Full Year"]:
                for week in SEMESTER_1_WEEKS:
                    week_record = student_df[student_df["Week"] == week]
                    c_flags, r_issues = _audit_week_row(week, week_record)
                    stu_issues.extend(r_issues)
                    if r_issues:
                        export_issue_rows.append({
                            "Student": student,
                            "ID": student_meta.get(student, {}).get("id", ""),
                            "Group": student_meta.get(student, {}).get("group", ""),
                            "Row": f"Week {week}",
                            "Issues": "; ".join(dict.fromkeys(r_issues)),
                        })
                    row = {
                        "Student": "",
                        "Row Type": f"Week {week}",
                        "_student_name": student,
                        "_week_num": week,
                        "_cell_flags": json.dumps(c_flags),
                        "_row_summary": "; ".join(dict.fromkeys(r_issues)),
                    }
                    for domain in SCORE_COLUMNS:
                        row[domain] = week_record.iloc[0][domain] if not week_record.empty else ""
                    row["Total Score"] = format_weekly_total_cell(week_record)
                    stu_block_rows.append(row)

                subtotal_flags = {d: sem1_math[f"act_{d}"] for d in SCORE_COLUMNS if f"act_{d}" in sem1_math}
                max_flags = {d: sem1_math[f"poss_{d}"] for d in SCORE_COLUMNS if f"poss_{d}" in sem1_math}
                stu_issues.extend([v["msg"] for v in subtotal_flags.values()] + [v["msg"] for v in max_flags.values()])

                subtotal_row = {
                    "Student": "",
                    "Row Type": "SEM 1 - Student Subtotal",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(subtotal_flags),
                    "_row_summary": "; ".join(v["msg"] for v in subtotal_flags.values()),
                }
                percentage_row = {"Student": "", "Row Type": "SEM 1 - Student Percentage", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}
                grade_row = {"Student": "", "Row Type": "SEM 1 - Categorical Grade", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}

                for domain in SCORE_COLUMNS:
                    subtotal_row[domain] = semester_totals[domain]
                    possible = semester_possible[domain]
                    if possible > 0:
                        pct = (semester_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                s1_tot_act = sum(semester_totals.values())
                s1_tot_poss = sum(semester_possible.values())
                subtotal_row["Total Score"] = s1_tot_act
                if s1_tot_poss > 0:
                    s1_tot_pct = (s1_tot_act / s1_tot_poss) * 100
                    percentage_row["Total Score"] = f"{s1_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(s1_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                stu_block_rows.extend([subtotal_row, percentage_row, grade_row])

                if reporting_period == "Semester 1":
                    sem1_max_row = {
                        "Student": "",
                        "Row Type": "SEM 1 - Maximum Possible",
                        "_student_name": student,
                        "_week_num": "",
                        "_cell_flags": json.dumps(max_flags),
                        "_row_summary": "; ".join(v["msg"] for v in max_flags.values()),
                    }
                    for domain in SCORE_COLUMNS:
                        sem1_max_row[domain] = semester_possible[domain]
                    sem1_max_row["Total Score"] = s1_tot_poss
                    stu_block_rows.append(sem1_max_row)

            # SEMESTER 2 ROWS
            if reporting_period in ["Semester 2", "Full Year"]:
                for week in SEMESTER_2_WEEKS:
                    week_record = student_df[student_df["Week"] == week]
                    c_flags, r_issues = _audit_week_row(week, week_record)
                    stu_issues.extend(r_issues)
                    if r_issues:
                        export_issue_rows.append({
                            "Student": student,
                            "ID": student_meta.get(student, {}).get("id", ""),
                            "Group": student_meta.get(student, {}).get("group", ""),
                            "Row": f"Week {week}",
                            "Issues": "; ".join(dict.fromkeys(r_issues)),
                        })
                    row = {
                        "Student": "",
                        "Row Type": f"Week {week}",
                        "_student_name": student,
                        "_week_num": week,
                        "_cell_flags": json.dumps(c_flags),
                        "_row_summary": "; ".join(dict.fromkeys(r_issues)),
                    }
                    for domain in SCORE_COLUMNS:
                        row[domain] = week_record.iloc[0][domain] if not week_record.empty else ""
                    row["Total Score"] = format_weekly_total_cell(week_record)
                    stu_block_rows.append(row)

                subtotal_flags = {d: sem2_math[f"act_{d}"] for d in SCORE_COLUMNS if f"act_{d}" in sem2_math}
                max_flags = {d: sem2_math[f"poss_{d}"] for d in SCORE_COLUMNS if f"poss_{d}" in sem2_math}
                stu_issues.extend([v["msg"] for v in subtotal_flags.values()] + [v["msg"] for v in max_flags.values()])

                subtotal_row = {
                    "Student": "",
                    "Row Type": "SEM 2 - Student Subtotal",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(subtotal_flags),
                    "_row_summary": "; ".join(v["msg"] for v in subtotal_flags.values()),
                }
                percentage_row = {"Student": "", "Row Type": "SEM 2 - Student Percentage", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}
                grade_row = {"Student": "", "Row Type": "SEM 2 - Categorical Grade", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}

                for domain in SCORE_COLUMNS:
                    subtotal_row[domain] = semester2_totals[domain]
                    possible = semester2_possible[domain]
                    if possible > 0:
                        pct = (semester2_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                s2_tot_act = sum(semester2_totals.values())
                s2_tot_poss = sum(semester2_possible.values())
                subtotal_row["Total Score"] = s2_tot_act
                if s2_tot_poss > 0:
                    s2_tot_pct = (s2_tot_act / s2_tot_poss) * 100
                    percentage_row["Total Score"] = f"{s2_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(s2_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                stu_block_rows.extend([subtotal_row, percentage_row, grade_row])

                if reporting_period == "Semester 2":
                    sem2_max_row = {
                        "Student": "",
                        "Row Type": "SEM 2 - Maximum Possible",
                        "_student_name": student,
                        "_week_num": "",
                        "_cell_flags": json.dumps(max_flags),
                        "_row_summary": "; ".join(v["msg"] for v in max_flags.values()),
                    }
                    for domain in SCORE_COLUMNS:
                        sem2_max_row[domain] = semester2_possible[domain]
                    sem2_max_row["Total Score"] = s2_tot_poss
                    stu_block_rows.append(sem2_max_row)

            # FULL YEAR SUMMARY & MAXIMUMS
            if reporting_period == "Full Year":
                full_year_totals = {}
                full_year_possible = {}

                for domain in SCORE_COLUMNS:
                    full_year_totals[domain] = semester_totals[domain] + semester2_totals[domain]
                    full_year_possible[domain] = semester_possible[domain] + semester2_possible[domain]

                fy_math = _verify_domain_period_math(
                    student_df, SEMESTER_1_WEEKS + SEMESTER_2_WEEKS, full_year_totals, full_year_possible, "Full Year"
                )
                fy_sub_flags = {d: fy_math[f"act_{d}"] for d in SCORE_COLUMNS if f"act_{d}" in fy_math}
                fy_max_flags = {d: fy_math[f"poss_{d}"] for d in SCORE_COLUMNS if f"poss_{d}" in fy_math}
                s1_max_flags = {d: sem1_math[f"poss_{d}"] for d in SCORE_COLUMNS if f"poss_{d}" in sem1_math}
                s2_max_flags = {d: sem2_math[f"poss_{d}"] for d in SCORE_COLUMNS if f"poss_{d}" in sem2_math}

                stu_issues.extend(
                    [v["msg"] for v in fy_sub_flags.values()]
                    + [v["msg"] for v in fy_max_flags.values()]
                )

                full_year_total_row = {
                    "Student": "",
                    "Row Type": "Full Year - Student Total",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(fy_sub_flags),
                    "_row_summary": "; ".join(v["msg"] for v in fy_sub_flags.values()),
                }
                percentage_row = {"Student": "", "Row Type": "Full Year - Student Percentage", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}
                grade_row = {"Student": "", "Row Type": "Full Year - Categorical Grade", "_student_name": student, "_week_num": "", "_cell_flags": "{}", "_row_summary": ""}

                for domain in SCORE_COLUMNS:
                    full_year_total_row[domain] = full_year_totals[domain]
                    possible = full_year_possible[domain]
                    if possible > 0:
                        pct = (full_year_totals[domain] / possible) * 100
                        percentage_row[domain] = f"{pct:.1f}%"
                        grade_row[domain] = get_categorical_grade_code(pct)
                    else:
                        percentage_row[domain] = ""
                        grade_row[domain] = ""

                fy_tot_act = sum(full_year_totals.values())
                fy_tot_poss = sum(full_year_possible.values())
                full_year_total_row["Total Score"] = fy_tot_act
                if fy_tot_poss > 0:
                    fy_tot_pct = (fy_tot_act / fy_tot_poss) * 100
                    percentage_row["Total Score"] = f"{fy_tot_pct:.1f}%"
                    grade_row["Total Score"] = get_categorical_grade_code(fy_tot_pct)
                else:
                    percentage_row["Total Score"] = ""
                    grade_row["Total Score"] = ""

                stu_block_rows.extend([full_year_total_row, percentage_row, grade_row])

                sem1_max_row = {
                    "Student": "",
                    "Row Type": "SEM 1 - Maximum Possible",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(s1_max_flags),
                    "_row_summary": "; ".join(v["msg"] for v in s1_max_flags.values()),
                }
                sem2_max_row = {
                    "Student": "",
                    "Row Type": "SEM 2 - Maximum Possible",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(s2_max_flags),
                    "_row_summary": "; ".join(v["msg"] for v in s2_max_flags.values()),
                }
                full_year_max_row = {
                    "Student": "",
                    "Row Type": "Full Year - Maximum Possible",
                    "_student_name": student,
                    "_week_num": "",
                    "_cell_flags": json.dumps(fy_max_flags),
                    "_row_summary": "; ".join(v["msg"] for v in fy_max_flags.values()),
                }

                for domain in SCORE_COLUMNS:
                    sem1_max_row[domain] = semester_possible[domain]
                    sem2_max_row[domain] = semester2_possible[domain]
                    full_year_max_row[domain] = full_year_possible[domain]

                sem1_max_row["Total Score"] = sum(semester_possible.values())
                sem2_max_row["Total Score"] = sum(semester2_possible.values())
                full_year_max_row["Total Score"] = fy_tot_poss

                stu_block_rows.extend([sem1_max_row, sem2_max_row, full_year_max_row])

            unique_stu_issues = list(dict.fromkeys(stu_issues))
            if unique_stu_issues:
                total_flagged_students += 1
                total_cell_issues += len(unique_stu_issues)

            student_blocks.append({
                "student": student,
                "issues": unique_stu_issues,
                "rows": stu_block_rows,
            })

        # Sort student blocks
        if val_sort_by == "First Name (A → Z)":
            student_blocks.sort(key=lambda b: b["student"].lower())
        elif val_sort_by == "Most Issues First":
            student_blocks.sort(key=lambda b: (-len(b["issues"]), b["student"].split()[-1].lower()))
        else:
            student_blocks.sort(key=lambda b: b["student"].split()[-1].lower())

        # Status bar & legend
        if not orphan_rows.empty:
            st.error(f"⚠️ {len(orphan_rows)} assessment row(s) have Student IDs not found in the active roster.")

        if total_flagged_students == 0:
            st.success(
                f"✅ Verified {len(unique_students)} student(s) in {reporting_period}: "
                "all weekly entries, domain subtotals, and maximum possible scores balance."
            )
        else:
            st.warning(
                f"🚩 **{total_flagged_students} of {len(unique_students)} student(s)** have validation flags ({total_cell_issues} total). "
                "**Red** = missing/invalid score or required comment (0, 5, or Absent - unnotified) · **Amber** = missing week record · "
                "**Purple** = subtotal/max calculation mismatch. Click any row below to fix it."
            )

        matrix_rows = []
        for blk in student_blocks:
            issues = blk["issues"]
            has_missing = any("No assessment logged" in i or "Duplicate entries" in i for i in issues)
            has_math = any("mismatch" in i.lower() for i in issues)
            has_score_flag = any(
                "No assessment logged" not in i
                and "Duplicate entries" not in i
                and "mismatch" not in i.lower()
                for i in issues
            )

            if val_filter == "All Issues (Missing + Flagged)" and not issues:
                continue
            elif val_filter == "🚩 Flagged Only (Scores / Comments)" and not has_score_flag:
                continue
            elif val_filter == "⚠️ Missing / Duplicate Weeks Only" and not has_missing:
                continue
            elif val_filter == "🧮 Math Mismatches Only" and not has_math:
                continue
            elif val_filter == "✅ Verified Only" and issues:
                continue

            stu = blk["student"]
            first_flagged_wk = next(
                (
                    str(r.get("_week_num", "")).strip()
                    for r in blk["rows"]
                    if r.get("_row_summary") and r.get("_week_num") != ""
                ),
                "",
            )

            header_label = stu.replace(" ", "\n", 1)
            if issues:
                header_label += f"\n🚩 {len(issues)} issue(s)"
            else:
                header_label += "\n✅ Verified"

            matrix_rows.append({
                "Student": header_label,
                "Row Type": "Student Header",
                "Engagement": "",
                "Teamwork": "",
                "Written Tasks": "",
                "Patience/Respect": "",
                "Accountability": "",
                "Uncertainty": "",
                "Punctuality": "",
                "Total Score": "",
                "_student_name": stu,
                "_week_num": first_flagged_wk,
                "_cell_flags": "{}",
                "_row_summary": "; ".join(issues),
            })
            matrix_rows.extend(blk["rows"])

        if not matrix_rows:
            st.info("No flagged students match the current filter. Toggle **🚩 Flagged Only** off to inspect all students.")
        else:
            matrix_rows.append({
                "Student": "END OF REPORT",
                "Row Type": "Student Header",
                "Engagement": "",
                "Teamwork": "",
                "Written Tasks": "",
                "Patience/Respect": "",
                "Accountability": "",
                "Uncertainty": "",
                "Punctuality": "",
                "Total Score": "",
                "_student_name": "",
                "_week_num": "",
                "_cell_flags": "{}",
                "_row_summary": "",
            })

            if val_view_mode == "Summary Rows Only":
                matrix_rows = [
                    r for r in matrix_rows
                    if not str(r.get("Row Type", "")).startswith("Week ")
                ]

            val_matrix_df = pd.DataFrame(matrix_rows)
            for col in SCORE_COLUMNS + ["Student", "Row Type", "_student_name", "_week_num", "_cell_flags", "_row_summary"]:
                if col in val_matrix_df.columns:
                    val_matrix_df[col] = val_matrix_df[col].fillna("").astype(str)

            val_gb = GridOptionsBuilder.from_dataframe(val_matrix_df)
            val_gb.configure_default_column(
                resizable=True,
                sortable=False,
                filter=False,
                wrapHeaderText=True,
                autoHeaderHeight=True,
            )

            val_gb.configure_column(
                "Student",
                width=185,
                pinned="left",
                wrapText=True,
                autoHeight=True,
                checkboxSelection=True,
                headerCheckboxSelection=False,
                cellStyle={
                    "white-space": "pre-line",
                    "word-break": "normal",
                    "overflow-wrap": "normal",
                    "line-height": "1.35",
                    "padding-top": "6px",
                    "padding-bottom": "6px",
                },
            )
            val_gb.configure_column(
                "Row Type",
                header_name="",
                width=235,
                pinned="left",
                wrapText=True,
                autoHeight=True,
                valueFormatter=JsCode("""
                    function(params) {
                        return params.value === 'Student Header' ? '' : params.value;
                    }
                """),
            )

            for hidden_col in ["_student_name", "_week_num", "_cell_flags", "_row_summary"]:
                val_gb.configure_column(hidden_col, hide=True)

            val_gb.configure_column("Engagement")
            val_gb.configure_column("Teamwork")
            val_gb.configure_column("Written Tasks")
            val_gb.configure_column("Patience/Respect", header_name="Patience/ Respect")
            val_gb.configure_column("Accountability")
            val_gb.configure_column("Uncertainty")
            val_gb.configure_column("Punctuality")

            val_grid_options = val_gb.build()
            val_grid_options["rowSelection"] = "single"
            val_grid_options["suppressRowClickSelection"] = False

            col_widths = {
                "Student": 150,
                "Row Type": 235,
                "Engagement": 134,
                "Teamwork": 120,
                "Written Tasks": 105,
                "Patience/Respect": 115,
                "Accountability": 148,
                "Uncertainty": 130,
                "Punctuality": 130,
                "Total Score": 95
            }

            cell_tooltip_js = JsCode("""
                function(params) {
                    try {
                        if (params.data && params.data['_cell_flags']) {
                            const flags = JSON.parse(params.data['_cell_flags']);
                            if (flags[params.colDef.field]) {
                                return '🚩 ' + flags[params.colDef.field].msg;
                            }
                        }
                    } catch (e) {}
                    const map = {
                        'C': 'Commendation',
                        'AS': 'At Standard',
                        'NY': 'Not Yet At Standard',
                        'CU': 'Concerning / Unacceptable Behaviour'
                    };
                    return map[params.value] || null;
                }
            """)

            cell_style_js = JsCode("""
                function(params) {
                    try {
                        if (params.data && params.data['_cell_flags']) {
                            const flags = JSON.parse(params.data['_cell_flags']);
                            const flag = flags[params.colDef.field];
                            if (flag) {
                                if (flag.level === 'error') {
                                    return {
                                        'background-color': '#FADBD8',
                                        'color': '#78281F',
                                        'font-weight': '700',
                                        'border': '1px solid #E74C3C'
                                    };
                                }
                                if (flag.level === 'warn') {
                                    return {
                                        'background-color': '#FCF3CF',
                                        'color': '#7D6608',
                                        'font-weight': '600',
                                        'border': '1px solid #F1C40F'
                                    };
                                }
                                if (flag.level === 'math') {
                                    return {
                                        'background-color': '#EBDEF0',
                                        'color': '#512E5F',
                                        'font-weight': '700',
                                        'border': '1px solid #8E44AD'
                                    };
                                }
                            }
                        }
                    } catch (e) {}
                    return null;
                }
            """)

            for col_def in val_grid_options["columnDefs"]:
                field = col_def.get("field")
                if field in ["Student", "Row Type"]:
                    col_def["pinned"] = "left"
                if field in col_widths:
                    col_def["width"] = col_widths[field]
                    col_def["minWidth"] = col_widths[field]
                    col_def["maxWidth"] = col_widths[field]
                    col_def["flex"] = 0
                    col_def["suppressSizeToFit"] = True
                if field in SCORE_COLUMNS + ["Total Score"]:
                    col_def["tooltipValueGetter"] = cell_tooltip_js
                    col_def["cellStyle"] = cell_style_js

            val_grid_options["tooltipShowDelay"] = 100
            val_grid_options["rowClassRules"] = {
                "student-header": JsCode("""
                    function(params) {
                        return params.data['Row Type'] === 'Student Header';
                    }
                """),
                "categorical-grade-row": JsCode("""
                    function(params) {
                        return params.data['Row Type'] && params.data['Row Type'].endsWith('Categorical Grade');
                    }
                """),
                "flagged-matrix-row": JsCode("""
                    function(params) {
                        return params.data && params.data['_row_summary'] && params.data['Row Type'] !== 'Student Header';
                    }
                """),
            }

            val_grid_response = AgGrid(
                val_matrix_df,
                gridOptions=val_grid_options,
                allow_unsafe_jscode=True,
                fit_columns_on_grid_load=False,
                columns_auto_size_mode=0,
                update_on=["selectionChanged"],
                custom_css={
                    ".ag-row-odd, .ag-row-even": {
                        "background-color": "#FFFFFF !important"
                    },
                    ".student-header": {
                        "background-color": "#D9EAF7 !important",
                        "font-weight": "bold !important"
                    },
                    ".categorical-grade-row": {
                        "background-color": "#FEF9E7 !important",
                        "font-weight": "600 !important",
                        "border-top": "1px solid #F3E8C8 !important",
                        "border-bottom": "1px solid #E6D8B5 !important"
                    },
                    ".flagged-matrix-row": {
                        "background-color": "#FFF9F8 !important"
                    },
                    ".ag-body-viewport": {
                        "padding-bottom": "35px !important"
                    },
                },
                height=760,
                theme="alpine",
                key=f"validation_matrix_grid_{selected_group}_{reporting_period}_{val_student_search}_{val_sort_by}_{val_view_mode}_{val_filter}_{pd.util.hash_pandas_object(val_matrix_df).sum()}",
            )

            # Click-to-fix drawer for any clicked row
            selected_val_rows = val_grid_response.get("selected_rows") or []
            if isinstance(selected_val_rows, pd.DataFrame):
                selected_val_rows = selected_val_rows.to_dict("records")

            if selected_val_rows:
                picked = selected_val_rows[0]
                p_stu = str(picked.get("_student_name", "")).strip()
                p_wk = str(picked.get("_week_num", "")).strip()
                p_summary = str(picked.get("_row_summary", "")).strip()

                if p_stu:
                    if p_summary:
                        st.warning(f"**{p_stu}** ({picked.get('Row Type', '')}): {p_summary}")
                    else:
                        st.info(f"**{p_stu}** ({picked.get('Row Type', '')}): No validation issues.")

                    if p_wk and p_stu in student_meta:
                        meta = student_meta[p_stu]
                        if st.button(
                            f"🔧 Open {p_stu} (Week {p_wk}) in Weekly Assessment",
                            type="primary",
                            key="val_grid_jump_btn",
                        ):
                            st.session_state["jump_to_assessment"] = {
                                "student_id": meta["id"],
                                "student_name": p_stu,
                                "group": f"Group {meta['group']}",
                                "week": p_wk,
                            }
                            st.session_state["pending_page"] = "Weekly Assessment"
                            save_ui_state()
                            st.rerun()

        # Build 2-tab Validation Excel export (Color-Coded Matrix + Actionable Checklist)
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.comments import Comment

        val_buffer = io.BytesIO()
        with pd.ExcelWriter(val_buffer, engine="openpyxl") as writer:
            # Tab 1: Validation Matrix (if rows exist)
            if matrix_rows:
                export_matrix_cols = ["Student", "Row Type"] + SCORE_COLUMNS
                export_mat_df = pd.DataFrame(matrix_rows)[export_matrix_cols].copy()
                export_mat_df["Row Type"] = export_mat_df["Row Type"].replace("Student Header", "")
                export_mat_df.to_excel(writer, index=False, sheet_name="Validation Matrix")

                ws_mat = writer.sheets["Validation Matrix"]
                thin_side = Side(style="thin", color="D9D9D9")
                cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                header_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
                grade_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")
                err_fill = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
                warn_fill = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
                math_fill = PatternFill(start_color="EBDEF0", end_color="EBDEF0", fill_type="solid")

                bold_font = Font(bold=True)
                err_font = Font(bold=True, color="78281F")
                warn_font = Font(bold=True, color="7D6608")
                math_font = Font(bold=True, color="512E5F")

                for col_idx in range(1, len(export_matrix_cols) + 1):
                    c = ws_mat.cell(row=1, column=col_idx)
                    c.font = bold_font
                    c.border = cell_border
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    ws_mat.column_dimensions[col_letter].width = 22 if col_idx <= 2 else 16

                for r_idx, raw_row in enumerate(matrix_rows, start=2):
                    r_type = str(raw_row.get("Row Type", ""))
                    is_stu_hdr = r_type == "Student Header"
                    is_cat_grade = r_type.endswith("Categorical Grade")
                    c_flags = json.loads(raw_row.get("_cell_flags", "{}") or "{}")

                    for c_idx, col_name in enumerate(export_matrix_cols, start=1):
                        cell = ws_mat.cell(row=r_idx, column=c_idx)
                        cell.border = cell_border
                        if c_idx == 1:
                            cell.alignment = Alignment(wrap_text=True, vertical="top")

                        if is_stu_hdr:
                            cell.fill = header_fill
                            cell.font = bold_font
                        elif is_cat_grade:
                            cell.fill = grade_fill
                            cell.font = bold_font

                        if col_name in c_flags:
                            flag_info = c_flags[col_name]
                            lvl = flag_info.get("level")
                            if lvl == "error":
                                cell.fill = err_fill
                                cell.font = err_font
                            elif lvl == "warn":
                                cell.fill = warn_fill
                                cell.font = warn_font
                            elif lvl == "math":
                                cell.fill = math_fill
                                cell.font = math_font
                            cell.comment = Comment(flag_info.get("msg", ""), "Validation Audit")

            # Tab 2: Flat Actionable Checklist
            export_df = (
                pd.DataFrame(export_issue_rows)
                if export_issue_rows
                else pd.DataFrame(columns=["Student", "ID", "Group", "Row", "Issues"])
            )
            export_df.to_excel(writer, index=False, sheet_name="Actionable Issues List")
            ws_list = writer.sheets["Actionable Issues List"]
            for col_idx, col_name in enumerate(export_df.columns, start=1):
                ws_list.cell(row=1, column=col_idx).font = Font(bold=True)
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws_list.column_dimensions[col_letter].width = 48 if col_name == "Issues" else 18

        period_slug = str(st.session_state.saved_period).replace(" ", "_")
        group_slug = str(st.session_state.saved_group).replace(" ", "_")

        st.download_button(
            label="📥 Download Validation Report (Excel)",
            data=val_buffer.getvalue(),
            file_name=f"Validation_Report_{group_slug}_{period_slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_validation_excel",
        )

# --- PAGE 3: WEEKLY ASSESSMENT ---
elif page == "Weekly Assessment":
    if st.session_state.student_data.empty:
        st.warning("Upload data first.")
    elif not is_logged_in:
        st.error("🔒 Please log in via the sidebar using your passcode.")
    else:

        # --- SMART WEEK SELECTION ---
        import pytz
        from datetime import datetime
        
        tz = pytz.timezone('Australia/Melbourne')
        today = datetime.now(tz).date()
        
        detected_idx = 0
        latest_valid_start = None

        # Scan the records to find the current week
        for i, row in enumerate(st.session_state.calendar_data):
            try:
                commence_date = row['Week commencing'].date()
                if commence_date <= today:
                    if latest_valid_start is None or commence_date > latest_valid_start:
                        latest_valid_start = commence_date
                        detected_idx = i
            except: continue

       # UI: The dropdown for the week
        week_options = []

        for r in st.session_state.calendar_data:

            if pd.isna(r['Week']):
                continue

            wk_int = int(float(r['Week']))

            if pd.notna(r['date']):
                week_label = (
                    f"Week {wk_int} "
                    f"({r['date'].strftime('%d %b %Y')})"
                )
            else:
                week_label = f"Week {wk_int} - No Classes this week"

            week_options.append(week_label)

        if detected_idx >= len(week_options):
            detected_idx = 0

        jump_target = st.session_state.get("jump_to_assessment")
        if jump_target:
            target_wk_int = str(int(float(jump_target['week'])))
            for w_idx, w_lbl in enumerate(week_options):
                if w_lbl.split(" ")[1] == target_wk_int:
                    detected_idx = w_idx
                    break

        selected_week = week_options[detected_idx]
        selected_week_str = str(int(float(selected_week.split(" ")[1])))

        user_groups = st.session_state.get("groups", [])

        if "ALL" in user_groups:

            group_list = ["All Groups"] + [
                f"Group {g}"
                for g in sorted(
                    st.session_state.student_data["Group"]
                    .dropna()
                    .unique()
                    .tolist()
                )
            ]

        else:

            group_list = [
                f"Group {g}"
                for g in user_groups
            ]

        # Format dates for the success box display
       # Get data for the selected week
        current_row = next(
            (
                item for item in st.session_state.calendar_data
                if pd.notna(item.get('Week'))
                and str(int(float(item['Week']))) == selected_week_str
            ),
            st.session_state.calendar_data[detected_idx]
        )

        if pd.notna(current_row['Week commencing']):
            fmt_commence = current_row['Week commencing'].strftime('%d %B %Y')
        else:
            fmt_commence = "No Classes this week"
        if pd.notna(current_row['date']):
            fmt_session = current_row['date'].strftime('%d %B %Y')
        else:
            fmt_session = "No Classes this week"

        assessment_header = st.container()

        with assessment_header:
            if fmt_session == "No Classes this week":

                st.warning(
                    f"Week {selected_week_str}: No Classes this week"
                )

            else:

                st.success(
                    f"**Week Commencing:** {fmt_commence} | "
                    f"**Week Number:** {selected_week_str} | "
                    f"**Specific Session Date:** {fmt_session}"
                )

# --- PDF RUBRIC (Forced Fit-to-Width) ---
            with st.expander("🔍 View Full Scoring Rubric", expanded=False):
                try:
                    with open("scoring_rubric.pdf", "rb") as f:
                        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
                    pdf_display = f'''
                    <div style="border: 2px solid #4A90E2; border-radius: 10px; overflow: hidden;">
                        <object data="data:application/pdf;base64,{base64_pdf}#toolbar=0&navpanes=0&scrollbar=1&page=1&view=FitH" 
                        type="application/pdf" width="100%" height="350px">
                            <embed src="data:application/pdf;base64,{base64_pdf}#view=FitH" type="application/pdf" />
                        </object>
                    </div>
                    '''
                    st.markdown(pdf_display, unsafe_allow_html=True)
                except:
                    st.error("scoring_rubric.pdf not found.")

            save_all = False
                                    
            if "saved_count" not in st.session_state:
                st.session_state.saved_count = 0

            if "modified_count" not in st.session_state:
                st.session_state.modified_count = 0

            if "remaining_count" not in st.session_state:
                st.session_state.remaining_count = 0
            
            toolbar = st.container(border=True)

            with toolbar:

                filter_col, spacer_col, status_col = st.columns([3, 1, 6])

                with filter_col:

                    selected_week = st.selectbox(
                        "Week",
                        week_options,
                        index=detected_idx,
                        label_visibility="collapsed"
                    )

                    role = st.session_state.get("role")

                    active_user = st.session_state.get("active_user")

                selected_week_str = str(int(float(selected_week.split(" ")[1])))

                if role in ["Coordinator", "AssessmentTeam"] or active_user == "Relief Tutor":

                    action1, action2 = st.columns([3,7])

                    # 1. Start from saved_group (synced across Dashboard & Assessment Analytics)
                    saved_grp = st.session_state.get("saved_group", "All Groups")
                    default_grp_label = "All Groups" if saved_grp == "All Groups" else f"Group {saved_grp}"

                    # 2. If jumping from Validation, override with the target student's group & sync saved_group
                    if jump_target and jump_target.get("group") in group_list:
                        default_grp_label = jump_target["group"]
                        raw_j = default_grp_label.replace("Group ", "", 1).strip()
                        st.session_state.saved_group = int(raw_j) if raw_j.isdigit() else raw_j
                        save_ui_state()

                    default_grp_idx = group_list.index(default_grp_label) if default_grp_label in group_list else 0

                    with action1:
                        sel_group = st.selectbox(
                            "Group",
                            group_list,
                            index=default_grp_idx,
                            label_visibility="collapsed",
                            key="weekly_group_select",
                            on_change=save_group,
                            args=("weekly_group_select",),
                        )

                    with action2:
                        save_all = st.button(
                            "💾 Save All",
                            use_container_width=True
                        )

                else:

                    sel_group = group_list[0]

                    save_all = st.button(
                        "💾 Save All",
                        use_container_width=True
                    )
      
                                
                # Build filtered view first

                if sel_group == "All Groups":

                    view_df = st.session_state.student_data.copy()

                else:

                    group_number = int(
                        sel_group.replace("Group ", "")
                    )

                    view_df = st.session_state.student_data[
                        st.session_state.student_data["Group"] == group_number
                    ]
                view_df = view_df.sort_values(
                    by=["family name", "preferred name"]
                )

                if jump_target:
                    target_id = str(jump_target.get("student_id", "")).strip()
                    target_name = jump_target.get("student_name", "")
                    target_wk = jump_target.get("week", "")

                    banner_c1, banner_c2 = st.columns([5, 1])
                    with banner_c1:
                        st.info(
                            f"🎯 **Focused from Validation:** {target_name} (ID: {target_id}) — Week {target_wk}"
                        )
                    with banner_c2:
                        if st.button("✖ Clear Focus", use_container_width=True):
                            del st.session_state["jump_to_assessment"]
                            st.rerun()

                    # Move the jumped student to the very top of the list
                    is_target = view_df["ID"].astype(str).str.strip() == target_id
                    if is_target.any():
                        view_df = pd.concat([view_df[is_target], view_df[~is_target]], ignore_index=True)  

                saved_count = 0
                modified_count = 0
                remaining_count = 0

                total_count = len(view_df)

                with spacer_col:
                    st.write("")
                                
                with status_col:

                    st.markdown(
                        f"✅ {st.session_state.saved_count} Saved &nbsp;&nbsp;|&nbsp;&nbsp; "
                        f"✏️ {st.session_state.modified_count} Unsaved Changes &nbsp;&nbsp;|&nbsp;&nbsp; "
                        f"⚪ {st.session_state.remaining_count} Remaining &nbsp;&nbsp;|&nbsp;&nbsp; "
                        f"⌨️ Type score + Enter",
                        unsafe_allow_html=True
                    )

                # SHOW SAVE-ALL RESULTS AT TOP OF PAGE

                if "save_all_success" in st.session_state:

                    st.success(
                        st.session_state.save_all_success
                    )

                    if st.session_state.get("skipped_students"):

                        st.warning(
                            f"{len(st.session_state.skipped_students)} assessment(s) were skipped. Expand below for details."
                        )

                        with st.expander("View skipped assessments"):

                            for item in st.session_state.skipped_students:
                                st.write(item)

                    del st.session_state.save_all_success
                    del st.session_state.skipped_students

        # --- THE SCORING LOOP ---
    
        if "unsaved_changes" not in st.session_state:
            st.session_state.unsaved_changes = False

        if st.session_state.unsaved_changes:
            st.warning("⚠ You have unsaved assessment changes on this page.")

        modified_count = 0
        
        all_assessments_to_save = []

        students_container = st.container(height=900)
        
        with students_container:

            for _, student in view_df.iterrows():
                s_id = str(student['ID'])
                s_email = str(student.get('Email', '')).strip().lower()
                
                full_name = (
                    f"{student.get('preferred name', '')} "
                    f"{student.get('family name', '')}"
                ).strip()

                status_icon = "⚪"

                selected_week_int = int(float(selected_week_str))
                existing_assessment = pd.DataFrame()

                if not st.session_state.assessments.empty:
                    wk_series = pd.to_numeric(
                        st.session_state.assessments["Week"],
                        errors="coerce"
                    )
                    existing_assessment = st.session_state.assessments[
                        (st.session_state.assessments["ID"].astype(str) == s_id)
                        & (wk_series == selected_week_int)
                    ]

                saved_record_exists = len(existing_assessment) > 0

                student_modified = False

                # New unsaved assessment

                if existing_assessment.empty:

                    attendance_key = f"att_{s_id}_{selected_week_str}"
                    comment_key = f"com_{s_id}_{selected_week_str}"

                    if (
                        attendance_key in st.session_state
                        and st.session_state[attendance_key] != "Present"
                    ):
                        student_modified = True

                    if (
                        comment_key in st.session_state
                        and str(st.session_state[comment_key]).strip() != ""
                    ):
                        student_modified = True

                    for score_cat in SCORE_COLUMNS:

                        possible_keys = [
                            f"s_{score_cat}_{s_id}_{selected_week_str}_Present",
                            f"s_{score_cat}_{s_id}_{selected_week_str}_Absent - notified",
                            f"s_{score_cat}_{s_id}_{selected_week_str}_Absent - unnotified"
                        ]

                        for key in possible_keys:

                            if key in st.session_state:

                                current_score = str(st.session_state[key])

                                if current_score not in ["---", "", "N/A"]:
                                    student_modified = True
                                    break

                if not existing_assessment.empty:

                    # Attendance

                    attendance_key = f"att_{s_id}_{selected_week_str}"

                    if attendance_key in st.session_state:

                        if (
                            str(st.session_state[attendance_key])
                            != str(existing_assessment.iloc[0]["Attendance"])
                        ):
                            student_modified = True
                    # Comment comparison

                    comment_key = f"com_{s_id}_{selected_week_str}"

                    if (
                        comment_key in st.session_state
                        and "Comment" in existing_assessment.columns
                    ):

                        if (
                            str(st.session_state[comment_key])
                            != str(existing_assessment.iloc[0]["Comment"])
                        ):
                            student_modified = True
                    # Score comparison

                    for score_cat in SCORE_COLUMNS:

                        if score_cat in existing_assessment.columns:

                                saved_score = str(
                                    existing_assessment.iloc[0][score_cat]
                                )

                                possible_keys = [
                                    f"s_{score_cat}_{s_id}_{selected_week_str}_Present",
                                    f"s_{score_cat}_{s_id}_{selected_week_str}_Absent - notified",
                                    f"s_{score_cat}_{s_id}_{selected_week_str}_Absent - unnotified"
                                ]

                                for key in possible_keys:

                                    if key in st.session_state:

                                        current_score = str(
                                            st.session_state[key]
                                        )

                                        if current_score != saved_score:
                                            student_modified = True
                                            break                              
                                
                if student_modified:

                    status_icon = "✏️"
                    modified_count += 1

                elif saved_record_exists:

                    status_icon = "✅"
                    saved_count += 1

                else:

                    status_icon = "⚪"
                    remaining_count += 1

                expander_label = (
                    f"{status_icon} {full_name} | Student ID: {s_id}"
                )
                                
                is_jumped_student = (
                    jump_target is not None
                    and str(jump_target.get("student_id")) == s_id
                )

                with st.expander(
                    expander_label,
                    expanded=True if (student_modified or is_jumped_student) else False
                ):

                    c1, c2 = st.columns([1, 4])

                    with c1:
                        if s_email in st.session_state.image_library:
                            st.image(st.session_state.image_library[s_email], width=150)
                    with c2:
                        # UPDATED: Changed wk_choice to selected_week_str to match the new logic
                        attendance_options = [
                            "Present",
                            "Absent - notified",
                            "Absent - unnotified"
                        ]

                        attendance_index = 0

                        if (
                            not existing_assessment.empty
                            and "Attendance" in existing_assessment.columns
                        ):

                            saved_attendance = str(
                                existing_assessment.iloc[0]["Attendance"]
                            )

                            if saved_attendance in attendance_options:
                                attendance_index = attendance_options.index(
                                    saved_attendance
                                )

                        att = st.radio(
                            "Attendance",
                            attendance_options,
                            index=attendance_index,
                            key=f"att_{s_id}_{selected_week_str}"
                        )
                        score_key = f"s_Accountability_{s_id}_{selected_week_str}"
                        if att == "Absent - notified":
                            st.session_state[score_key] = "3"

                        elif att == "Absent - unnotified":
                            st.session_state[score_key] = "0"

                        elif score_key in st.session_state:
                            del st.session_state[score_key]
                    st.divider()
                    
                    scores = {}
                    cols = st.columns(len(SCORE_COLUMNS))
                    
                    for i, cat in enumerate(SCORE_COLUMNS):
                        if "Absent" in att and cat != "Accountability":
                            scores[cat] = "N/A"
                            cols[i].selectbox(
                                cat,
                                [""],
                                disabled=True,
                                key=f"na_{cat}_{s_id}_{selected_week_str}"
                            )
                        else:
                            opts = ["---", "N/A", "I/O", "0", "1", "2", "3", "4", "5"]
                            if cat == "Accountability" and att == "Absent - notified":
                                opts = ["1", "2", "3", "4", "5"]

                                scores[cat] = cols[i].selectbox(
                                    cat,
                                    opts,
                                    index=2,
                                    key=f"s_{cat}_{s_id}_{selected_week_str}_{att}"
                                )

                            elif cat == "Accountability" and att == "Absent - unnotified":
                                opts = ["0", "1", "2", "3", "4", "5"]

                                scores[cat] = cols[i].selectbox(
                                    cat,
                                    opts,
                                    index=0,
                                    key=f"s_{cat}_{s_id}_{selected_week_str}_{att}"
                                )

                            else:
                                default_idx = 0

                                if (
                                    not existing_assessment.empty
                                    and cat in existing_assessment.columns
                                ):
                                    saved_value = str(
                                        existing_assessment.iloc[0][cat]
                                    )

                                    if saved_value in opts:
                                        default_idx = opts.index(saved_value)

                                scores[cat] = cols[i].selectbox(
                                    cat,
                                    opts,
                                    index=default_idx,
                                    key=f"s_{cat}_{s_id}_{selected_week_str}_{att}"
                                )

                    # --- CORRECTED COMMENT LOGIC & SAVE BUTTON ---
                    
                    # Denominator (Item 15)
                    valid_scores = [v for v in scores.values() if str(v).isdigit()]
                    st.caption(f"**Denominator:** {len(valid_scores)}")

                    # Comment Logic - FIXED: Changed wk_choice to selected_week_str
                    must_comment = any(v in ["0", "5"] for v in scores.values())
                    saved_comment = ""

                    if (
                        not existing_assessment.empty
                        and "Comment" in existing_assessment.columns
                    ):
                        saved_comment = str(
                            existing_assessment.iloc[0]["Comment"]
                        )

                    must_comment = any(
                        v in ["0", "5"]
                        for v in scores.values()
                    )

                    positive_comment_suggestions = []
                    concern_comment_suggestions = []

                    for category, score in scores.items():

                        if str(score) == "4":

                            positive_comment_suggestions.append(
                                category
                            )

                        elif str(score) in ["1", "2"]:

                            concern_comment_suggestions.append(
                                category
                            )

                    if concern_comment_suggestions:

                        st.info(
                            "Consider adding context for areas of concern: "
                            + ", ".join(concern_comment_suggestions)
                        )

                    if positive_comment_suggestions:

                        st.info(
                            "Positive feedback may be helpful for: "
                            + ", ".join(positive_comment_suggestions)
                        )

                    comm = st.text_area(
                        "Comment (Required for 0 or 5)" if must_comment else "Comment",
                        value=saved_comment,
                        key=f"com_{s_id}_{selected_week_str}"
                    )
                    comment_missing = (
                        must_comment
                        and len(comm.strip()) < 5
                    )
                    if comment_missing:
                        st.error(
                            "Comment required for 0 or 5."
                        )

                    if not existing_assessment.empty:

                        if str(existing_assessment.iloc[0]["Attendance"]) != str(att):
                            student_modified = True

                        if str(existing_assessment.iloc[0]["Comment"]) != str(comm):
                            student_modified = True

                        for cat in SCORE_COLUMNS:

                            if cat in existing_assessment.columns:

                                if str(existing_assessment.iloc[0][cat]) != str(scores[cat]):
                                    student_modified = True

                    if student_modified:
                        st.warning(
                            "✏️ This assessment has unsaved changes"
                        )

                    student_assessment = {
                        "ID": s_id,
                        "Name": full_name,
                        "Week": int(selected_week_str),
                        "Role": active_user,
                        "Group": student["Group"],
                        "Attendance": att,
                        "Comment": comm,
                        **scores
                    }

                    # Only queue students with actual assessment data
                    has_score = any(
                        str(v) not in ["---", "", "N/A"]
                        for v in scores.values()
                    )

                    if has_score:
                        all_assessments_to_save.append(student_assessment)

                    # Save Button - FIXED: Changed wk_choice to selected_week_str
                    if st.button(f"Save Week {selected_week_str} for {student.get('preferred name')}", key=f"btn_{s_id}_{selected_week_str}"):
                        if "---" in scores.values():
                            st.error("Select all scores.")
                        elif comment_missing:
                            pass
                        else:
                            # Logic to save the data
                            new_entry = {
                                "ID": s_id,
                                "Name": full_name,
                                "Week": int(float(selected_week_str)),
                                "Role": active_user,
                                "Group": student["Group"],
                                "Attendance": att,
                                "Comment": comm,
                                **scores
                            }

                            # Remove existing assessment
                            if not st.session_state.assessments.empty:
                                wk_series = pd.to_numeric(
                                    st.session_state.assessments["Week"],
                                    errors="coerce"
                                )
                                st.session_state.assessments = (
                                    st.session_state.assessments[
                                        ~(
                                            (
                                                st.session_state.assessments["ID"].astype(str)
                                                == s_id
                                            )
                                            &
                                            (
                                                wk_series
                                                == int(float(selected_week_str))
                                            )
                                        )
                                    ]
                                )

                            # Save latest version

                            st.session_state.assessments = pd.concat(
                                [
                                    st.session_state.assessments,
                                    pd.DataFrame([new_entry])
                                ],
                                ignore_index=True
                            )
                            
                            # Save assessments to disk

                            st.session_state.assessments.to_pickle(
                                "assessments.pkl"
                            )

                            st.rerun()
        if save_all:

            saved_counter = 0
            skipped_students = []

            for assessment in all_assessments_to_save:

                score_values = [
                    v
                    for k, v in assessment.items()
                    if k in SCORE_COLUMNS
                ]

                # Skip untouched students

                if all(v == "---" for v in score_values):
                    skipped_students.append(
                        f"{assessment['Name']} - No scores entered"
                    )
                    continue
                if "---" in score_values:
                    skipped_students.append(
                        f"{assessment['Name']} - Assessment incomplete"
                    )
                    continue

                # Remove existing assessment
                if not st.session_state.assessments.empty:
                    wk_series = pd.to_numeric(
                        st.session_state.assessments["Week"],
                        errors="coerce"
                    )
                    st.session_state.assessments = (
                        st.session_state.assessments[
                            ~(
                                (
                                    st.session_state.assessments["ID"].astype(str)
                                    == str(assessment["ID"])
                                )
                                &
                                (
                                    wk_series
                                    == int(float(assessment["Week"]))
                                )
                            )
                        ]
                    )

                st.session_state.assessments = pd.concat(
                    [
                        st.session_state.assessments,
                        pd.DataFrame([assessment])
                    ],
                    ignore_index=True
                )

                saved_counter += 1

            st.session_state.assessments.to_pickle(
                "assessments.pkl"
            )

            st.session_state.save_all_success = (
                f"Saved {saved_counter} assessments."
            )

            st.session_state.skipped_students = skipped_students

            st.rerun()
        if (
            st.session_state.saved_count != saved_count
            or st.session_state.modified_count != modified_count
            or st.session_state.remaining_count != remaining_count
        ):
            st.session_state.saved_count = saved_count
            st.session_state.modified_count = modified_count
            st.session_state.remaining_count = remaining_count
            st.rerun()

# --- PAGE 4: DASHBOARD (Item 17 Preview) ---
elif page == "Dashboard":
    st.header("Results Overview")
    if st.session_state.assessments.empty:
        st.info("No data recorded.")
# TODO:
# Replace with Monash SSO authentication check
    elif not is_logged_in:
        st.error("🔒 Please log in via the sidebar using your passcode.")
    else:

        df = st.session_state.assessments.copy()

        user_groups = st.session_state.get("groups", [])
        dash_selected_group = "All Groups"

        if "ALL" in user_groups or len(user_groups) > 1:
            if not st.session_state.student_data.empty and "Group" in st.session_state.student_data.columns:
                raw_groups = st.session_state.student_data["Group"].dropna().unique().tolist()
                if "ALL" not in user_groups:
                    raw_groups = [g for g in raw_groups if g in user_groups]
                sorted_groups = sorted(raw_groups, key=lambda x: int(x) if str(x).isdigit() else str(x))
                dash_group_options = ["All Groups"] + [f"Group {g}" for g in sorted_groups]
            else:
                dash_group_options = ["All Groups"]

            saved_grp = st.session_state.get("saved_group", "All Groups")
            default_dash_grp = "All Groups" if saved_grp == "All Groups" else f"Group {saved_grp}"
            if default_dash_grp not in dash_group_options:
                default_dash_grp = "All Groups"

            dash_selected_group = st.selectbox(
                "Filter by Group",
                dash_group_options,
                index=dash_group_options.index(default_dash_grp),
                key="dashboard_group_filter",
                on_change=save_group,
                args=("dashboard_group_filter",),
            )

            if dash_selected_group != "All Groups":
                chosen_grp = dash_selected_group.replace("Group ", "").strip()
                allowed_students = st.session_state.student_data[
                    st.session_state.student_data["Group"].astype(str).str.strip() == chosen_grp
                ]
                if "ID" in df.columns and "ID" in allowed_students.columns:
                    df = df[df["ID"].astype(str).isin(allowed_students["ID"].astype(str))]
                else:
                    allowed_names = (
                        allowed_students["preferred name"].fillna("")
                        + " "
                        + allowed_students["family name"].fillna("")
                    ).str.strip()
                    df = df[df["Name"].isin(allowed_names)]

        elif "ALL" not in user_groups:
            allowed_students = st.session_state.student_data[
                st.session_state.student_data["Group"].isin(user_groups)
            ]
            if "ID" in df.columns and "ID" in allowed_students.columns:
                df = df[df["ID"].astype(str).isin(allowed_students["ID"].astype(str))]
            else:
                allowed_names = (
                    allowed_students["preferred name"].fillna("")
                    + " "
                    + allowed_students["family name"].fillna("")
                ).str.strip()
                df = df[df["Name"].isin(allowed_names)]

        report_options = [
            "Tutor Group Summary",
            "Individual Student Summary",
            "Assessment Completion",
            "Raw Assessment Data"
        ]

        report_col, attendance_col, status_col = st.columns(
            [2, 1, 1.5]
        )

        with report_col:

            report_view = st.radio(
                "**Report**",
                report_options
            )

        with attendance_col:

            st.markdown("**Attendance**")

            st.markdown("""
            ✅ Present  
            🟣 Absent - Notified  
            🔴 Absent - Unnotified
            """)

        with status_col:

            st.markdown("**Status**")

            st.markdown("""
            ⭐ Commendation  
            🔷 At Standard  
            ⏳ Not Yet At Standard  
            🔺 Concerning / Unacceptable Behaviour  
            ⚪ No Assessments Recorded
            """)


        if report_view == "Tutor Group Summary":

            st.subheader(
                "Tutor Group Summary"
            )

            period_options = ["Semester 1", "Semester 2", "Full Year"]
            if st.session_state.get("saved_period") not in period_options:
                st.session_state.saved_period = "Semester 1"

            period = st.radio(
                "**Reporting Period**",
                period_options,
                index=period_options.index(st.session_state.saved_period),
                horizontal=True,
                key="dashboard_period_radio",
                on_change=save_period,
                args=("dashboard_period_radio",)
            )

            cohort_df = st.session_state.student_data.copy()
            df = filter_by_period(df, period)

            if "ALL" not in user_groups:

                cohort_df = cohort_df[
                    cohort_df["Group"].isin(user_groups)
                ]

            if dash_selected_group != "All Groups":
                chosen_grp = dash_selected_group.replace("Group ", "").strip()
                cohort_df = cohort_df[
                    cohort_df["Group"].astype(str).str.strip() == chosen_grp
                ]

            cohort_df["Full Name"] = (
                cohort_df["preferred name"].fillna("")
                + " "
                + cohort_df["family name"].fillna("")
            ).str.strip()

            student_list = (
                cohort_df["Full Name"]
                .tolist()
            )
            
            group_summary = []

            for student_name in student_list:

                student_df = df[
                    df["Name"] == student_name
                ]

                student_row = {
                    "Student": student_name
                }

                available_weeks = sorted(
                    df["Week"]
                    .dropna()
                    .unique()
                )

                for week in available_weeks:

                    week_num = int(float(week))

                    week_df = student_df[
                        student_df["Week"] == week
                    ]

                    if week_df.empty:

                        student_row[f"W{week_num}"] = "—"

                    else:

                        row = week_df.iloc[0]

                        weekly_total, weekly_possible = (
                            calculate_week_totals(
                                row,
                                SCORE_COLUMNS
                            )
                        )

                        attendance = str(row["Attendance"])

                        attendance_icon = "✅"

                        if attendance == "Absent - notified":
                            attendance_icon = "🟣"

                        elif attendance == "Absent - unnotified":
                            attendance_icon = "🔴"

                        student_row[f"W{week_num}"] = (
                            f"{weekly_total} {attendance_icon}"
                        )

                actual_total = 0
                possible_total = 0

                weekly_percentages = []
                
                for _, row in student_df.iterrows():

                    weekly_total, weekly_possible = (
                        calculate_week_totals(
                            row,
                            SCORE_COLUMNS
                        )
                    )

                    actual_total += weekly_total
                    possible_total += weekly_possible

                    if weekly_possible > 0:

                        weekly_percentages.append(
                            (weekly_total / weekly_possible) * 100
                        )

                if possible_total > 0:

                    percentage = (
                        actual_total
                        / possible_total
                    ) * 100

                else:

                    percentage = 0

                if len(weekly_percentages) < 2:

                    trend = "⚪"

                else:

                    first = weekly_percentages[0]
                    last = weekly_percentages[-1]

                    if last >= first + 10:

                        trend = "📈"

                    elif last <= first - 10:

                        trend = "📉"

                    else:

                        trend = "➡️"

                if possible_total == 0:

                    status = "⚪"

                else:

                    grade = get_categorical_grade(
                        percentage
                    )

                    status_map = {
                        "Commendation": "⭐",
                        "At Standard": "🔷",
                        "Not Yet At Standard": "⏳",
                        "Concerning / Unacceptable Behaviour": "🔺"
                    }

                    status = status_map[grade]

                student_row["Overall %"] = (
                    f"{percentage:.1f}%"
                    if possible_total > 0
                    else "N/A"
                )

                student_row["Trend"] = trend

                student_row["Status"] = status

                group_summary.append(
                    student_row
                )

            group_summary_df = pd.DataFrame(
                group_summary
            )

            week_cols = sorted(
                [
                    col
                    for col in group_summary_df.columns
                    if col.startswith("W")
                ],
                key=lambda x: int(x[1:])
            )

            group_summary_df = group_summary_df[
                ["Student"]
                + week_cols
                + ["Overall %", "Trend", "Status"]
            ]

            group_summary_df = group_summary_df.reset_index(
                drop=True
            )


            gb = GridOptionsBuilder.from_dataframe(
                group_summary_df
            )

            gb.configure_default_column(
                width=85,
                resizable=True
            )

            gb.configure_selection(
                selection_mode="single",
                use_checkbox=True
            )

            gb.configure_column(
                "Student",
                pinned="left",
                width=220
            )

            grid_options = gb.build()

            week_cols = [
                col
                for col in group_summary_df.columns
                if col.startswith("W")
            ]

            grid_options["columnDefs"] = [
                {
                    "field": "Student",
                    "pinned": "left",
                    "width": 220,
                    "checkboxSelection": True,
                    "headerCheckboxSelection": False
                },
                {
                    "headerName": "Weekly Results (Score + Attendance)",
                    "children": [
                        {"field": col}
                        for col in week_cols
                    ]
                },
                {
                    "field": "Overall %",
                    "pinned": "right",
                    "width": 100
                },
                {
                    "field": "Trend",
                    "pinned": "right",
                    "width": 75
                },
                {
                    "field": "Status",
                    "pinned": "right",
                    "width": 80
                }
            ]
            
            grid_response = AgGrid(
                group_summary_df,
                gridOptions=grid_options,
                height=800,
                fit_columns_on_grid_load=False,
                theme="alpine",
                custom_css={
                    ".ag-body-viewport": {
                        "padding-bottom": "35px !important"
                    }
                }
            )

            selected_rows = (
                grid_response.get("selected_rows")
                or []
            )

            if selected_rows:

                selected_student = (
                    selected_rows[0]["Student"]
                )

                st.divider()

                st.subheader(
                    f"Individual Student Summary: {selected_student}"
                )

                student_df = df[
                    df["Name"] == selected_student
                ]

                render_individual_student_summary(
                    student_df
                )
        
        if report_view == "Individual Student Summary":
        
            st.subheader("Individual Student Summary")

            period_options = ["Semester 1", "Semester 2", "Full Year"]
            if st.session_state.get("saved_period") not in period_options:
                st.session_state.saved_period = "Semester 1"

            period = st.radio(
                "**Reporting Period**",
                period_options,
                index=period_options.index(st.session_state.saved_period),
                horizontal=True,
                key="dashboard_period_radio",
                on_change=save_period,
                args=("dashboard_period_radio",)
            )

            cohort_df = st.session_state.student_data.copy()
            df = filter_by_period(df, period)

            cohort_df = st.session_state.student_data.copy()

            if "ALL" not in user_groups:

                cohort_df = cohort_df[
                    cohort_df["Group"].isin(user_groups)
                ]
            if dash_selected_group != "All Groups":
                chosen_grp = dash_selected_group.replace("Group ", "").strip()
                cohort_df = cohort_df[
                    cohort_df["Group"].astype(str).str.strip() == chosen_grp
                ]

            cohort_df["Full Name"] = (
                cohort_df["preferred name"].fillna("")
                + " "
                + cohort_df["family name"].fillna("")
            ).str.strip()

            student_list = (
                cohort_df["Full Name"]
                .tolist()
            ) 

            default_index = 0

            if (
                "dashboard_student" in st.session_state
                and st.session_state.dashboard_student in student_list
            ):
                default_index = student_list.index(
                    st.session_state.dashboard_student
                )

            selected_student = st.selectbox(
                "Select Student",
                student_list,
                index=default_index
            )

            # Filter assessment data by semester
            df = filter_by_period(df, period)

            student_df = df[
                df["Name"] == selected_student
            ]

            render_individual_student_summary(
                student_df
            )

        elif report_view == "Assessment Completion":

            st.subheader("Assessment Completion")

            all_weeks = pd.to_numeric(st.session_state.assessments["Week"], errors="coerce").dropna()
            current_week = int(all_weeks.max()) if not all_weeks.empty else 1

            st.info(
                f"Current Teaching Week: {int(current_week)}"
            )

            cohort_df = st.session_state.student_data.copy()

            if "ALL" not in user_groups:

                cohort_df = cohort_df[
                    cohort_df["Group"].isin(user_groups)
                ]

            if dash_selected_group != "All Groups":
                chosen_grp = dash_selected_group.replace("Group ", "").strip()
                cohort_df = cohort_df[
                    cohort_df["Group"].astype(str).str.strip() == chosen_grp
                ]

            missing_by_week = {}

            for week in range(1, int(current_week) + 1):

                missing_students = []

                for _, student in cohort_df.iterrows():

                    student_id = str(student["ID"])

                    assessment_exists = len(
                        df[
                            (df["ID"].astype(str) == student_id)
                            &
                            (
                                pd.to_numeric(
                                    df["Week"],
                                    errors="coerce"
                                ) == week
                            )
                        ]
                    ) > 0

                    if not assessment_exists:

                        full_name = (
                            f"{student.get('preferred name', '')} "
                            f"{student.get('family name', '')}"
                        ).strip()

                        missing_students.append(
                            full_name
                        )

                if len(missing_students) > 0:

                    missing_by_week[week] = (
                        missing_students
                    )
            total_outstanding = sum(
                len(students)
                for students in missing_by_week.values()
            )

            if total_outstanding == 0:

                st.success(
                    "✅ All assessments are complete."
                )

            else:

                st.warning(
                    "⚠ You have overdue assessments."
                )

                st.metric(
                    "Total Outstanding Assessments",
                    total_outstanding
                )

                for week, students in sorted(
                    missing_by_week.items()
                ):

                    with st.expander(
                        f"Week {week} ({len(students)} students)"
                    ):

                        for student in students:

                            st.write(student)
        
        elif report_view == "Raw Assessment Data":

            st.subheader(
                "Raw Assessment Data"
            )

            display_cols = ["Name"] + [
            col for col in df.columns
            if col != "Name"
            and col != "ID"
        ]

            display_df = df[display_cols].copy()

            display_df["Comment"] = (
                display_df["Comment"]
                .fillna("")
                .astype(str)
            )

            display_df["Comment"] = display_df["Comment"].apply(
                lambda x: (
                    x[:30] + "..."
                    if len(x) > 30
                    else x
                )
            )

            st.data_editor(
                display_df,
                use_container_width=True,
                hide_index=True,
                disabled=True,
                column_config={
                    "Comment": st.column_config.TextColumn(
                        "Comment",
                        width="medium"
                    )
                }
            )

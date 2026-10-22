"""
EduPro Instructor Performance and Course Quality Evaluation Dashboard
Run with: streamlit run app.py
"""
from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="EduPro | Instructor & Course Quality",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
# Automatically locate the workbook bundled with this app. No file uploader is shown.
WORKBOOK_CANDIDATES = [
    BASE_DIR / "data" / "EduPro Online Platform.xlsx",
    BASE_DIR / "EduPro Online Platform.xlsx",
    BASE_DIR / "EduPro instructor program dashboard.xlsx",
    BASE_DIR / "EduPro instructor program dashboard.xlsm",
    BASE_DIR / "DOC-20261021-WA0001(1).xlsx",
    BASE_DIR / "DOC-20261021-WA0001.xlsx",
]

REQUIRED_COLUMNS = {
    "Users": {"UserID", "UserName", "Age", "Gender", "Email"},
    "Teachers": {"TeacherID", "TeacherName", "Age", "Gender", "Expertise",
                 "YearsOfExperience", "TeacherRating"},
    "Courses": {"CourseID", "CourseName", "CourseCategory", "CourseType",
                "CourseLevel", "CoursePrice", "CourseDuration", "CourseRating"},
    "Transactions": {"TransactionID", "UserID", "CourseID", "TransactionDate",
                     "Amount", "PaymentMethod", "TeacherID"},
}

st.markdown("""
<style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {
        background: #f7f9fc; border: 1px solid #e6eaf0;
        padding: 14px 16px; border-radius: 12px;
    }
    [data-testid="stMetricLabel"] {font-size: 0.88rem;}
    .section-note {color: #5f6b7a; font-size: 0.92rem;}
    h1, h2, h3 {letter-spacing: -0.02em;}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_workbook(path_string: str):
    """Load the actual workbook and validate sheet/column contracts."""
    path = Path(path_string)
    if not path.exists():
        raise FileNotFoundError(f"Dataset file not found at: {path}")
    excel = pd.ExcelFile(path, engine="openpyxl")
    missing_sheets = set(REQUIRED_COLUMNS) - set(excel.sheet_names)
    if missing_sheets:
        raise ValueError(
            f"Missing required worksheet(s): {', '.join(sorted(missing_sheets))}. "
            f"Found sheets: {', '.join(excel.sheet_names)}"
        )
    tables = {sheet: pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
              for sheet in REQUIRED_COLUMNS}
    errors = []
    for sheet, required in REQUIRED_COLUMNS.items():
        missing = required - set(tables[sheet].columns)
        if missing:
            errors.append(f"{sheet}: missing columns {', '.join(sorted(missing))}")
    if errors:
        raise ValueError("Workbook column validation failed: " + "; ".join(errors))

    # Parse dates/numeric values defensively; do not redo or overwrite source cleaning.
    tables["Transactions"]["TransactionDate"] = pd.to_datetime(
        tables["Transactions"]["TransactionDate"], errors="coerce"
    )
    for sheet, columns in {
        "Teachers": ["Age", "YearsOfExperience", "TeacherRating"],
        "Courses": ["CoursePrice", "CourseDuration", "CourseRating"],
        "Transactions": ["Amount"],
    }.items():
        for col in columns:
            tables[sheet][col] = pd.to_numeric(tables[sheet][col], errors="coerce")

    # Keep dimensions separate; enrich transactions with many-to-one lookups only.
    teachers = tables["Teachers"].copy()
    courses = tables["Courses"].copy()
    transactions = tables["Transactions"].copy()
    teacher_keys = teachers["TeacherID"].dropna()
    course_keys = courses["CourseID"].dropna()
    if teacher_keys.duplicated().any():
        raise ValueError("Teachers.TeacherID contains duplicates; instructor-level metrics would be ambiguous.")
    if course_keys.duplicated().any():
        raise ValueError("Courses.CourseID contains duplicates; course-level metrics would be ambiguous.")

    transaction_enriched = transactions.merge(
        courses[["CourseID", "CourseName", "CourseCategory", "CourseType",
                 "CourseLevel", "CoursePrice", "CourseRating"]],
        on="CourseID", how="left", validate="many_to_one", indicator="_course_join"
    ).merge(
        teachers[["TeacherID", "TeacherName", "Expertise", "Gender",
                  "YearsOfExperience", "TeacherRating"]],
        on="TeacherID", how="left", validate="many_to_one", indicator="_teacher_join"
    )
    quality = {
        "unmatched_course_transactions": int((transaction_enriched["_course_join"] == "left_only").sum()),
        "unmatched_teacher_transactions": int((transaction_enriched["_teacher_join"] == "left_only").sum()),
        "transaction_rows": int(len(transactions)),
        "duplicate_transaction_ids": int(transactions["TransactionID"].duplicated().sum()),
        "missing_transaction_dates": int(transactions["TransactionDate"].isna().sum()),
    }
    transaction_enriched = transaction_enriched.drop(columns=["_course_join", "_teacher_join"])
    return tables, transaction_enriched, quality


def multi_filter(frame, column, selected):
    if column not in frame.columns or not selected:
        return frame.copy()
    return frame[frame[column].isin(selected)].copy()


def safe_mean(series):
    return float(series.mean()) if series.notna().any() else None


def fmt_number(value, decimals=0):
    if value is None or pd.isna(value):
        return "—"
    return f"{value:,.{decimals}f}"


def interpretation(text):
    st.markdown(f'<p class="section-note">{text}</p>', unsafe_allow_html=True)


def make_download(frame, label):
    st.download_button(
        label=f"Download {label} (CSV)",
        data=frame.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"edupro_{label.lower().replace(' ', '_')}.csv",
        mime="text/csv",
        key=f"download_{label.lower().replace(' ', '_')}",
    )


st.title("EduPro | Instructor Performance & Course Quality")
st.caption("Interactive management dashboard • Instructor quality, course ratings, and transaction activity")

# Load the bundled workbook automatically. There is intentionally no upload control.
with st.sidebar:
    st.header("Dashboard controls")
    workbook_path = next((p for p in WORKBOOK_CANDIDATES if p.is_file()), None)
    if workbook_path is None:
        # Fallback: discover an Excel workbook in the repository root or data folder.
        possible_files = sorted(BASE_DIR.glob("*.xlsx")) + sorted((BASE_DIR / "data").glob("*.xlsx")) if (BASE_DIR / "data").exists() else sorted(BASE_DIR.glob("*.xlsx"))
        for candidate in possible_files:
            try:
                names = set(pd.ExcelFile(candidate, engine="openpyxl").sheet_names)
                if set(REQUIRED_COLUMNS).issubset(names):
                    workbook_path = candidate
                    break
            except Exception:
                continue
    if workbook_path is None:
        st.error("EduPro dataset was not found in this GitHub repository.")
        st.info("Keep the Excel dataset in the same folder as app.py, or in a folder named data. The dashboard does not require manual upload.")
        st.stop()
    try:
        tables, tx, quality = load_workbook(str(workbook_path))
    except Exception as exc:
        st.error(f"Unable to load the bundled EduPro dataset: {exc}")
        st.stop()

    teachers_all = tables["Teachers"].copy()
    courses_all = tables["Courses"].copy()
    tx_all = tx.copy()

    expertise_options = sorted(teachers_all["Expertise"].dropna().astype(str).unique())
    category_options = sorted(courses_all["CourseCategory"].dropna().astype(str).unique())
    level_options = sorted(courses_all["CourseLevel"].dropna().astype(str).unique())
    gender_options = sorted(teachers_all["Gender"].dropna().astype(str).unique())

    selected_expertise = st.multiselect("Instructor expertise", expertise_options, default=expertise_options)
    selected_gender = st.multiselect("Instructor gender", gender_options, default=gender_options)
    selected_categories = st.multiselect("Course category", category_options, default=category_options)
    selected_levels = st.multiselect("Course level", level_options, default=level_options)

    ratings = pd.to_numeric(teachers_all["TeacherRating"], errors="coerce").dropna()
    if ratings.empty:
        rating_range = (0.0, 5.0)
        st.caption("No valid teacher ratings found; using a neutral slider range.")
    else:
        rating_min, rating_max = float(ratings.min()), float(ratings.max())
        if rating_min == rating_max:
            rating_range = (rating_min, rating_max)
            st.caption(f"Teacher rating is constant at {rating_min:.2f}.")
        else:
            rating_range = st.slider(
                "Teacher rating range", min_value=rating_min, max_value=rating_max,
                value=(rating_min, rating_max), step=0.1
            )

    valid_dates = tx_all["TransactionDate"].dropna()
    if valid_dates.empty:
        date_range = None
        st.caption("No valid transaction dates; date filter is unavailable.")
    else:
        min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
        date_range = st.date_input("Transaction date range", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date)
        if not isinstance(date_range, (tuple, list)) or len(date_range) != 2:
            date_range = (min_date, max_date)

# Dimension filters apply to their own grains. Transaction filters use the transaction's
# recorded TeacherID/CourseID; they do not imply a permanent teacher-course assignment.
teachers_f = teachers_all[
    teachers_all["Expertise"].astype(str).isin(selected_expertise)
    & teachers_all["Gender"].astype(str).isin(selected_gender)
    & pd.to_numeric(teachers_all["TeacherRating"], errors="coerce").between(
        rating_range[0], rating_range[1], inclusive="both"
    )
].copy()
courses_f = courses_all[
    courses_all["CourseCategory"].astype(str).isin(selected_categories)
    & courses_all["CourseLevel"].astype(str).isin(selected_levels)
].copy()
tx_f = tx_all[
    tx_all["CourseCategory"].astype(str).isin(selected_categories)
    & tx_all["CourseLevel"].astype(str).isin(selected_levels)
    & tx_all["Expertise"].astype(str).isin(selected_expertise)
    & tx_all["Gender"].astype(str).isin(selected_gender)
    & pd.to_numeric(tx_all["TeacherRating"], errors="coerce").between(
        rating_range[0], rating_range[1], inclusive="both"
    )
].copy()
if date_range and len(date_range) == 2:
    start_date, end_date = pd.Timestamp(date_range[0]), pd.Timestamp(date_range[1]) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
    tx_f = tx_f[tx_f["TransactionDate"].between(start_date, end_date, inclusive="both")].copy()

# Header and data-quality notes
with st.expander("Data validation & metric definitions", expanded=False):
    st.write(f"**Workbook sheets:** {', '.join(tables.keys())}")
    st.write(f"**Transaction rows loaded:** {quality['transaction_rows']:,}")
    st.write(f"**Unmatched course references:** {quality['unmatched_course_transactions']:,}")
    st.write(f"**Unmatched instructor references:** {quality['unmatched_teacher_transactions']:,}")
    st.write(f"**Duplicate transaction IDs:** {quality['duplicate_transaction_ids']:,}")
    st.write(f"**Transactions with invalid/missing dates:** {quality['missing_transaction_dates']:,}")
    st.caption("Instructor and course ratings are dimension-level averages; transaction amounts are summed from Transactions.Amount. The recorded TeacherID on a transaction is used for transaction attribution only. It does not establish a permanent catalog-level teacher-course assignment.")
    st.caption("The 'free' versus 'paid' label comes from Courses.CourseType. Revenue is reported as recorded transaction amount, not assumed course list price.")

if quality["unmatched_course_transactions"] or quality["unmatched_teacher_transactions"]:
    st.warning("Some transactions do not match a course or instructor record. Those rows are retained for transaction counts but may be excluded from category/instructor-specific charts.")
if quality["duplicate_transaction_ids"]:
    st.warning("Duplicate TransactionID values were detected. Counts use transaction rows; inspect duplicates before treating each row as a unique transaction.")

tab_overview, tab_instructors, tab_courses, tab_expertise, tab_transactions = st.tabs([
    "Executive Overview", "Instructor Performance", "Course Quality",
    "Expertise Analysis", "Enrollment & Transactions"
])

with tab_overview:
    st.subheader("Executive Overview")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Instructors", f"{len(teachers_f):,}")
    c2.metric("Courses", f"{len(courses_f):,}")
    c3.metric("Transactions", f"{len(tx_f):,}")
    c4.metric("Avg. teacher rating", fmt_number(safe_mean(teachers_f["TeacherRating"]), 2))
    c5.metric("Avg. course rating", fmt_number(safe_mean(courses_f["CourseRating"]), 2))
    c6.metric("Avg. experience (years)", fmt_number(safe_mean(teachers_f["YearsOfExperience"]), 1))

    left, right = st.columns(2)
    with left:
        if not tx_f.empty:
            free_paid = tx_f.groupby("CourseType", dropna=False).size().reset_index(name="Transactions")
            fig = px.pie(free_paid, names="CourseType", values="Transactions", hole=0.55,
                         title="Free vs. Paid Transaction Distribution")
            fig.update_traces(textposition="inside", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
            interpretation("Shows the share of recorded transaction rows by the course's Free/Paid classification.")
        else:
            st.info("No transactions match the current filters.")
    with right:
        if not tx_f.empty:
            monthly = tx_f.assign(Month=tx_f["TransactionDate"].dt.to_period("M").dt.to_timestamp())
            monthly = monthly.dropna(subset=["Month"]).groupby("Month").size().reset_index(name="Transactions")
            fig = px.line(monthly, x="Month", y="Transactions", markers=True,
                          title="Monthly Transaction Trend")
            st.plotly_chart(fig, use_container_width=True)
            interpretation("Monthly volume is calculated from filtered transaction rows, using the transaction date.")
        else:
            st.info("No transactions match the current filters.")
    st.subheader("Filtered data export")
    export_source = tx_f.drop(columns=[], errors="ignore")
    make_download(export_source, "filtered_transactions")

with tab_instructors:
    st.subheader("Instructor Performance")
    if teachers_f.empty:
        st.info("No instructors match the current expertise, gender, and rating filters.")
    else:
        ranking = teachers_f.sort_values("TeacherRating", ascending=False).copy()
        ranking["TransactionCount"] = ranking["TeacherID"].map(tx_f.groupby("TeacherID").size()).fillna(0).astype(int)
        col_a, col_b = st.columns(2)
        with col_a:
            top = ranking.head(10).sort_values("TeacherRating")
            fig = px.bar(top, x="TeacherRating", y="TeacherName", orientation="h",
                         color="TeacherRating", color_continuous_scale="Blues",
                         hover_data=["Expertise", "YearsOfExperience", "TransactionCount"],
                         title="Top 10 Instructors by Teacher Rating")
            fig.update_layout(coloraxis_showscale=False, yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            bottom = ranking.sort_values("TeacherRating", ascending=True).head(10).sort_values("TeacherRating")
            fig = px.bar(bottom, x="TeacherRating", y="TeacherName", orientation="h",
                         color="TeacherRating", color_continuous_scale="Reds_r",
                         hover_data=["Expertise", "YearsOfExperience", "TransactionCount"],
                         title="Bottom 10 Instructors by Teacher Rating")
            fig.update_layout(coloraxis_showscale=False, yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        interpretation("Leaderboards use TeacherRating from the Teachers sheet. Transaction volume is a separate activity measure and should not be interpreted as teaching quality.")
        col_c, col_d = st.columns(2)
        with col_c:
            fig = px.histogram(teachers_f, x="TeacherRating", nbins=10,
                               title="Teacher Rating Distribution", labels={"TeacherRating": "Teacher rating"})
            st.plotly_chart(fig, use_container_width=True)
        with col_d:
            tx_counts = tx_f.groupby("TeacherID").size().rename("TransactionCount").reset_index()
            volume = teachers_f.merge(tx_counts, on="TeacherID", how="left", validate="one_to_one")
            volume["TransactionCount"] = volume["TransactionCount"].fillna(0)
            volume = volume.sort_values("TransactionCount", ascending=False).head(15)
            fig = px.bar(volume, x="TeacherName", y="TransactionCount", color="Expertise",
                         title="Instructor Transaction Volume (Top 15)",
                         hover_data=["TeacherRating", "YearsOfExperience"])
            fig.update_layout(xaxis_tickangle=-35, xaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        fig = px.scatter(
            teachers_f, x="YearsOfExperience", y="TeacherRating", color="Expertise",
            hover_name="TeacherName", hover_data=["Gender", "Age"],
            title="Teaching Experience vs. Teacher Rating",
            labels={"YearsOfExperience": "Teaching experience (years)", "TeacherRating": "Teacher rating"}
        )
        st.plotly_chart(fig, use_container_width=True)
        interpretation("The scatter plot describes association only; it does not prove that experience causes higher ratings.")
        make_download(ranking, "instructor_performance")

with tab_courses:
    st.subheader("Course Quality")
    if courses_f.empty:
        st.info("No courses match the selected category and level filters.")
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            by_category = courses_f.groupby("CourseCategory", dropna=False)["CourseRating"].mean().reset_index()
            by_category = by_category.sort_values("CourseRating", ascending=True)
            fig = px.bar(by_category, x="CourseRating", y="CourseCategory", orientation="h",
                         title="Average Course Rating by Category",
                         labels={"CourseRating": "Average course rating", "CourseCategory": "Category"})
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            by_level = courses_f.groupby("CourseLevel", dropna=False)["CourseRating"].mean().reset_index()
            fig = px.bar(by_level, x="CourseLevel", y="CourseRating", color="CourseLevel",
                         title="Average Course Rating by Level",
                         labels={"CourseRating": "Average course rating", "CourseLevel": "Course level"})
            st.plotly_chart(fig, use_container_width=True)
        fig = px.histogram(courses_f, x="CourseRating", nbins=10, color="CourseType",
                           title="Course Rating Distribution")
        st.plotly_chart(fig, use_container_width=True)
        col_c, col_d = st.columns(2)
        with col_c:
            st.markdown("**Highest-rated courses**")
            st.dataframe(courses_f.sort_values("CourseRating", ascending=False).head(10)[
                ["CourseName", "CourseCategory", "CourseLevel", "CourseType", "CourseRating"]
            ], use_container_width=True, hide_index=True)
        with col_d:
            st.markdown("**Lowest-rated courses**")
            st.dataframe(courses_f.sort_values("CourseRating", ascending=True).head(10)[
                ["CourseName", "CourseCategory", "CourseLevel", "CourseType", "CourseRating"]
            ], use_container_width=True, hide_index=True)
        interpretation("Course ratings are taken from the Courses sheet, so each course contributes once to course-level averages regardless of transaction volume.")
        make_download(courses_f, "course_quality")

with tab_expertise:
    st.subheader("Expertise Analysis")
    if teachers_f.empty:
        st.info("No instructors match the current filters.")
    else:
        expertise_summary = teachers_f.groupby("Expertise").agg(
            Instructors=("TeacherID", "nunique"),
            AvgTeacherRating=("TeacherRating", "mean"),
            AvgExperience=("YearsOfExperience", "mean")
        ).reset_index().sort_values("Instructors", ascending=False)
        fig = px.bar(expertise_summary, x="Expertise", y="Instructors", color="AvgTeacherRating",
                     hover_data=["AvgTeacherRating", "AvgExperience"],
                     title="Instructor Expertise Distribution")
        fig.update_layout(xaxis_tickangle=-35, xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

        # Exact text match is a limited proxy, not a semantic skill-match score.
        expertise_category = tx_f.dropna(subset=["Expertise", "CourseCategory"]).copy()
        expertise_category["ExactLabelMatch"] = (
            expertise_category["Expertise"].astype(str).str.strip().str.casefold()
            == expertise_category["CourseCategory"].astype(str).str.strip().str.casefold()
        )
        if not expertise_category.empty:
            match_summary = expertise_category.groupby("CourseCategory").agg(
                TransactionRows=("TransactionID", "size"),
                ExactLabelMatches=("ExactLabelMatch", "sum")
            ).reset_index()
            match_summary["ExactLabelMatchRatePct"] = (
                match_summary["ExactLabelMatches"] / match_summary["TransactionRows"] * 100
            )
            fig = px.bar(match_summary, x="CourseCategory", y="ExactLabelMatchRatePct",
                         hover_data=["TransactionRows", "ExactLabelMatches"],
                         title="Exact Expertise Label Match by Course Category (%)",
                         labels={"ExactLabelMatchRatePct": "Exact label match (%)", "CourseCategory": "Course category"})
            fig.update_layout(xaxis_tickangle=-35, xaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
            st.caption("This is an exact text-label comparison between the transaction's recorded instructor Expertise and CourseCategory. It is not a validated subject-matter match score; related labels can differ, and matching labels do not prove instructional fit.")
        else:
            st.info("No transaction rows have both instructor expertise and course category available.")
        interpretation("The workbook records TeacherID on transactions, which supports transaction-level attribution. It does not define a permanent instructor-to-course assignment in the Courses catalog.")
        make_download(expertise_summary, "expertise_summary")

with tab_transactions:
    st.subheader("Enrollment & Transactions")
    c1, c2, c3 = st.columns(3)
    c1.metric("Filtered transaction amount", f"{fmt_number(tx_f['Amount'].sum(), 2)}")
    paid_tx = tx_f[tx_f["CourseType"].astype(str).str.casefold() == "paid"]
    free_tx = tx_f[tx_f["CourseType"].astype(str).str.casefold() == "free"]
    c2.metric("Paid transaction amount", f"{fmt_number(paid_tx['Amount'].sum(), 2)}")
    c3.metric("Free transaction rows", f"{len(free_tx):,}")

    if tx_f.empty:
        st.info("No transactions match the current filters and date range.")
    else:
        monthly = tx_f.dropna(subset=["TransactionDate"]).copy()
        monthly["Month"] = monthly["TransactionDate"].dt.to_period("M").dt.to_timestamp()
        monthly_summary = monthly.groupby("Month").agg(
            Transactions=("TransactionID", "size"),
            RecordedAmount=("Amount", "sum")
        ).reset_index()
        fig = go.Figure()
        fig.add_trace(go.Bar(x=monthly_summary["Month"], y=monthly_summary["Transactions"],
                             name="Transactions", yaxis="y"))
        fig.add_trace(go.Scatter(x=monthly_summary["Month"], y=monthly_summary["RecordedAmount"],
                                 name="Recorded amount", mode="lines+markers", yaxis="y2"))
        fig.update_layout(
            title="Monthly Transactions and Recorded Amount",
            xaxis_title="Month",
            yaxis=dict(title="Transaction rows"),
            yaxis2=dict(title="Recorded amount", overlaying="y", side="right"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            hovermode="x unified"
        )
        st.plotly_chart(fig, use_container_width=True)

        col_a, col_b = st.columns(2)
        with col_a:
            cat_tx = tx_f.groupby("CourseCategory").size().reset_index(name="Transactions").sort_values("Transactions", ascending=False)
            fig = px.bar(cat_tx, x="CourseCategory", y="Transactions", color="Transactions",
                         title="Transactions by Course Category")
            fig.update_layout(xaxis_tickangle=-35, xaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        with col_b:
            level_tx = tx_f.groupby("CourseLevel").size().reset_index(name="Transactions")
            fig = px.bar(level_tx, x="CourseLevel", y="Transactions", color="CourseLevel",
                         title="Transactions by Course Level")
            st.plotly_chart(fig, use_container_width=True)

        free_paid = tx_f.groupby("CourseType").agg(
            Transactions=("TransactionID", "size"),
            RecordedAmount=("Amount", "sum")
        ).reset_index()
        fig = px.bar(free_paid, x="CourseType", y="Transactions", color="CourseType",
                     hover_data=["RecordedAmount"], title="Free vs Paid Course Activity")
        st.plotly_chart(fig, use_container_width=True)

        teacher_share = tx_f.groupby(["TeacherID", "TeacherName"], dropna=False).agg(
            Transactions=("TransactionID", "size"),
            RecordedAmount=("Amount", "sum")
        ).reset_index().sort_values("Transactions", ascending=False)
        teacher_share["TransactionSharePct"] = (
            teacher_share["Transactions"] / teacher_share["Transactions"].sum() * 100
        )
        fig = px.bar(teacher_share.head(20), x="TeacherName", y="TransactionSharePct",
                     color="RecordedAmount", hover_data=["Transactions", "RecordedAmount"],
                     title="Instructor Share of Filtered Transactions (Top 20)",
                     labels={"TransactionSharePct": "Share of transactions (%)"})
        fig.update_layout(xaxis_tickangle=-35, xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
        interpretation("Recorded transaction amount is summed from Transactions.Amount. For free courses, a zero amount is expected; for paid courses, it is not automatically equivalent to recognized revenue without payment/refund/accounting definitions.")
        make_download(tx_f, "filtered_transactions_detailed")
        make_download(monthly_summary, "monthly_transactions")

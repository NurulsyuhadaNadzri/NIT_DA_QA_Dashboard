from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="NIT / DA / Conventional QA Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("NIT / DA / Conventional Quality Monitoring")
st.caption(
    "Moisture • Protein • Ash | Daily, Weekly & Monthly Review"
)


# ============================================================
# FILE
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

EXCEL_FILE = (
    BASE_DIR
    / "NIT,DA & Conventional Report- Recovery.xlsx"
)


# ============================================================
# COLORS
# ============================================================

# Status
GREEN = "#2ECC71"
GREEN_DARK = "#196F3D"
GREEN_LIGHT = "#D5F5E3"

RED = "#E74C3C"
RED_DARK = "#922B21"
RED_LIGHT = "#FADBD8"

GREY = "#7B7D7D"
GREY_LIGHT = "#F2F3F4"

# Trend
BLUE = "#1F618D"

# Method
CONV_COLOR = "#7B2CBF"   # Purple
LAB_COLOR = "#F59E0B"    # Orange
DA_COLOR = "#0077B6"     # Blue


# ============================================================
# HELPERS
# ============================================================

def normalize_column_name(column):

    return " ".join(
        str(column)
        .strip()
        .upper()
        .split()
    )


def to_numeric(series):

    return pd.to_numeric(
        series.replace(
            {
                "": np.nan,
                "-": np.nan,
                "–": np.nan,
                "—": np.nan,
            }
        ),
        errors="coerce",
    )


def clean_status(series):

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
    )


def first_existing_column(
    df,
    candidates,
):

    for column in candidates:

        if column in df.columns:
            return column

    return None


def create_period(
    date_series,
    trend_view,
):

    if trend_view == "Daily":

        return (
            date_series
            .dt
            .normalize()
        )

    elif trend_view == "Weekly":

        return (
            date_series
            .dt
            .to_period("W-SUN")
            .dt
            .start_time
        )

    else:

        return (
            date_series
            .dt
            .to_period("M")
            .dt
            .start_time
        )


def make_period_label(
    period,
    view,
):

    period = pd.Timestamp(period)

    if view == "Weekly":

        week_end = (
            period
            + pd.Timedelta(days=6)
        )

        return (
            f"{period.strftime('%d %b')} - "
            f"{week_end.strftime('%d %b')}"
        )

    elif view == "Monthly":

        return (
            period.strftime(
                "%b %Y"
            )
        )

    else:

        return (
            period.strftime(
                "%d %b"
            )
        )


def format_time_value(value):

    if pd.isna(value):
        return ""

    if hasattr(value, "hour"):

        try:

            return (
                f"{int(value.hour):02d}:"
                f"{int(value.minute):02d}"
            )

        except Exception:
            pass

    if isinstance(
        value,
        (
            int,
            float,
            np.integer,
            np.floating,
        ),
    ):

        try:

            fraction = float(value) % 1

            total_seconds = round(
                fraction
                * 24
                * 60
                * 60
            )

            hours = (
                total_seconds // 3600
            ) % 24

            minutes = (
                total_seconds % 3600
            ) // 60

            return (
                f"{hours:02d}:"
                f"{minutes:02d}"
            )

        except Exception:
            pass

    try:

        parsed = pd.to_datetime(
            value,
            errors="coerce",
        )

        if not pd.isna(parsed):

            return (
                parsed.strftime(
                    "%H:%M"
                )
            )

    except Exception:
        pass

    return ""


def time_to_delta(value):

    if pd.isna(value):
        return pd.Timedelta(0)

    if hasattr(value, "hour"):

        try:

            return pd.Timedelta(
                hours=int(value.hour),
                minutes=int(value.minute),
                seconds=int(
                    getattr(
                        value,
                        "second",
                        0,
                    )
                ),
            )

        except Exception:
            pass

    if isinstance(
        value,
        (
            int,
            float,
            np.integer,
            np.floating,
        ),
    ):

        try:

            fraction = float(value) % 1

            return pd.Timedelta(
                days=fraction
            )

        except Exception:
            pass

    try:

        parsed = pd.to_datetime(
            value,
            errors="coerce",
        )

        if not pd.isna(parsed):

            return pd.Timedelta(
                hours=parsed.hour,
                minutes=parsed.minute,
                seconds=parsed.second,
            )

    except Exception:
        pass

    return pd.Timedelta(0)


def highlight_status(value):

    value = (
        str(value)
        .strip()
        .lower()
    )

    if value == "within tolerance":

        return (
            f"background-color: {GREEN_LIGHT}; "
            f"color: {GREEN_DARK}; "
            "font-weight: bold;"
        )

    if value == "out of tolerance":

        return (
            f"background-color: {RED_LIGHT}; "
            f"color: {RED_DARK}; "
            "font-weight: bold;"
        )

    return (
        f"background-color: {GREY_LIGHT}; "
        f"color: {GREY};"
    )


# ============================================================
# READ WORKBOOK
# ============================================================

def read_workbook(file_path):

    xls = pd.ExcelFile(
        file_path,
        engine="openpyxl",
    )


    sheet_map = {

        str(sheet)
        .strip()
        .lower(): sheet

        for sheet
        in xls.sheet_names
    }


    required_sheets = {

        "moisture":
            "Moisture",

        "protein":
            "Protein",

        "ash":
            "Ash",
    }


    frames = []


    for (
        sheet_key,
        parameter,
    ) in required_sheets.items():


        if sheet_key not in sheet_map:

            raise ValueError(
                f"Sheet '{parameter}' not found."
            )


        actual_sheet = (
            sheet_map[
                sheet_key
            ]
        )


        df = pd.read_excel(
            file_path,
            sheet_name=actual_sheet,
            engine="openpyxl",
        )


        # ====================================================
        # SOURCE ROW
        # ====================================================

        df["SOURCE_ROW"] = (
            df.index + 2
        )


        # ====================================================
        # DROP EMPTY COLUMNS
        # ====================================================

        df = (
            df
            .dropna(
                axis=1,
                how="all",
            )
            .copy()
        )


        # ====================================================
        # NORMALIZE HEADERS
        # ====================================================

        df.columns = [

            normalize_column_name(
                column
            )

            for column
            in df.columns
        ]


        # ====================================================
        # DATE
        # ====================================================

        if "DATE" not in df.columns:

            raise ValueError(
                f"{parameter}: DATE column not found."
            )


        df["DATE"] = pd.to_datetime(
            df["DATE"],
            errors="coerce",
            dayfirst=True,
        )


        df = (
            df[
                df[
                    "DATE"
                ]
                .notna()
            ]
            .copy()
        )


        # ====================================================
        # CONV
        # ====================================================

        if "CONV" in df.columns:

            df["CONV"] = (
                to_numeric(
                    df["CONV"]
                )
            )

        else:

            df["CONV"] = np.nan


        # ====================================================
        # DA
        # ====================================================

        if "DA" in df.columns:

            df["DA"] = (
                to_numeric(
                    df["DA"]
                )
            )

        else:

            df["DA"] = np.nan


        # ====================================================
        # LAB = NOVA / NIT
        # ====================================================

        if parameter == "Ash":

            if "NIT" in df.columns:

                df["LAB"] = (
                    to_numeric(
                        df["NIT"]
                    )
                )

            else:

                df["LAB"] = np.nan

        else:

            if "NOVA" in df.columns:

                df["LAB"] = (
                    to_numeric(
                        df["NOVA"]
                    )
                )

            else:

                df["LAB"] = np.nan


        # ====================================================
        # PROCESS
        # ====================================================

        if "PROCESS" not in df.columns:

            process_column = (
                first_existing_column(
                    df,
                    [
                        "MILL/ MIX/ PCK",
                        "MILL/MIX/PCK",
                        "MILL / MIX / PCK",
                    ],
                )
            )

            if process_column:

                df["PROCESS"] = (
                    df[
                        process_column
                    ]
                )


        # ====================================================
        # TIME
        # ====================================================

        if "TIME" in df.columns:

            df[
                "TIME_DISPLAY"
            ] = (
                df["TIME"]
                .apply(
                    format_time_value
                )
            )


            time_delta = (
                df["TIME"]
                .apply(
                    time_to_delta
                )
            )

        else:

            df[
                "TIME_DISPLAY"
            ] = ""


            time_delta = (
                pd.Series(
                    pd.Timedelta(0),
                    index=df.index,
                )
            )


        # ====================================================
        # DATETIME
        # ====================================================

        df["DATETIME"] = (
            df["DATE"]
            .dt
            .normalize()
            + time_delta
        )


        # ====================================================
        # PARAMETER
        # ====================================================

        df["PARAMETER"] = (
            parameter
        )


        tolerance = (

            0.02
            if parameter == "Ash"
            else 0.20

        )


        df["TOLERANCE"] = (
            tolerance
        )


        # ====================================================
        # MOISTURE / PROTEIN
        # ====================================================

        if parameter in [
            "Moisture",
            "Protein",
        ]:


            # ================================================
            # NOVA VS CONV
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "NOVA - CONV",
                        "NOVA-CONV",
                    ],
                )
            )


            if difference_column:

                df[
                    "LAB vs CONV"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "LAB vs CONV"
                ] = (
                    df["LAB"]
                    - df["CONV"]
                )


            # ================================================
            # DA VS CONV
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "DA - CONV",
                        "DA-CONV",
                    ],
                )
            )


            if difference_column:

                df[
                    "DA vs CONV"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "DA vs CONV"
                ] = (
                    df["DA"]
                    - df["CONV"]
                )


            # ================================================
            # DA VS NOVA
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "DA - NOVA",
                        "DA-NOVA",
                    ],
                )
            )


            if difference_column:

                df[
                    "DA vs LAB"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "DA vs LAB"
                ] = (
                    df["DA"]
                    - df["LAB"]
                )


            # ================================================
            # NOVA VS CONV STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "NOVA-CONV STATUS",
                        "NOVA - CONV STATUS",
                    ],
                )
            )


            if status_column:

                df[
                    "LAB vs CONV STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "LAB vs CONV STATUS"
                ] = ""


            # ================================================
            # DA VS CONV STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "DA - CONV STATUS",
                        "DA-CONV STATUS",
                    ],
                )
            )


            if status_column:

                df[
                    "DA vs CONV STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "DA vs CONV STATUS"
                ] = ""


            # ================================================
            # DA VS NOVA STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "DA - NOVA STATUS",
                        "DA-NOVA STATUS",
                    ],
                )
            )


            if status_column:

                df[
                    "DA vs LAB STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "DA vs LAB STATUS"
                ] = ""


        # ====================================================
        # ASH
        # ====================================================

        else:


            # ================================================
            # NIT VS CONV
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "NIT - CONV",
                        "NIT-CONV",
                    ],
                )
            )


            if difference_column:

                df[
                    "LAB vs CONV"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "LAB vs CONV"
                ] = (
                    df["LAB"]
                    - df["CONV"]
                )


            # ================================================
            # DA VS CONV
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "DA - CONV",
                        "DA-CONV",
                    ],
                )
            )


            if difference_column:

                df[
                    "DA vs CONV"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "DA vs CONV"
                ] = (
                    df["DA"]
                    - df["CONV"]
                )


            # ================================================
            # DA VS NIT
            # ================================================

            difference_column = (
                first_existing_column(
                    df,
                    [
                        "DA-NIT",
                        "DA - NIT",
                    ],
                )
            )


            if difference_column:

                df[
                    "DA vs LAB"
                ] = to_numeric(
                    df[
                        difference_column
                    ]
                )

            else:

                df[
                    "DA vs LAB"
                ] = (
                    df["DA"]
                    - df["LAB"]
                )


            # ================================================
            # NIT VS CONV STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "NIT - CONV STATUS",
                        "NIT-CONV STATUS",
                    ],
                )
            )


            if status_column:

                df[
                    "LAB vs CONV STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "LAB vs CONV STATUS"
                ] = ""


            # ================================================
            # DA VS CONV STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "DA - CONV STATUS",
                        "DA - CONV STAUS",
                        "DA-CONV STATUS",
                        "DA-CONV STAUS",
                    ],
                )
            )


            if status_column:

                df[
                    "DA vs CONV STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "DA vs CONV STATUS"
                ] = ""


            # ================================================
            # DA VS NIT STATUS
            # ================================================

            status_column = (
                first_existing_column(
                    df,
                    [
                        "DA-NIT STATUS",
                        "DA - NIT STATUS",
                    ],
                )
            )


            if status_column:

                df[
                    "DA vs LAB STATUS"
                ] = clean_status(
                    df[
                        status_column
                    ]
                )

            else:

                df[
                    "DA vs LAB STATUS"
                ] = ""


        frames.append(
            df
        )


    return pd.concat(
        frames,
        ignore_index=True,
        sort=False,
    )


# ============================================================
# CACHE
# ============================================================

@st.cache_data(
    show_spinner=False
)
def load_data(
    file_path,
    modified_time,
):

    return read_workbook(
        file_path
    )


# ============================================================
# FILE CHECK
# ============================================================

if not EXCEL_FILE.exists():

    st.error(
        "Excel file not found."
    )

    st.code(
        str(EXCEL_FILE)
    )

    st.stop()


# ============================================================
# AUTO REFRESH
# ============================================================

@st.fragment(
    run_every="20s"
)
def auto_refresh():

    current_modified = (
        EXCEL_FILE
        .stat()
        .st_mtime
    )


    if (
        "last_excel_modified"
        not in st.session_state
    ):

        st.session_state[
            "last_excel_modified"
        ] = current_modified


    elif (
        current_modified
        !=
        st.session_state[
            "last_excel_modified"
        ]
    ):

        st.session_state[
            "last_excel_modified"
        ] = current_modified

        st.cache_data.clear()

        st.rerun()


auto_refresh()


# ============================================================
# LOAD DATA
# ============================================================

try:

    file_modified = (
        EXCEL_FILE
        .stat()
        .st_mtime
    )


    df = load_data(
        str(EXCEL_FILE),
        file_modified,
    )


except Exception as error:

    st.error(
        "Unable to read workbook."
    )

    st.exception(
        error
    )

    st.stop()


# ============================================================
# HEADER
# ============================================================

file_modified_display = (
    pd.Timestamp(
        EXCEL_FILE
        .stat()
        .st_mtime,
        unit="s",
    )
)


header_left, header_right = (
    st.columns(
        [
            7,
            1.5,
        ],
        vertical_alignment="center",
    )
)


with header_left:

    st.caption(
        "Source file last updated: "
        f"{file_modified_display.strftime('%d %b %Y %I:%M:%S %p')}"
    )


with header_right:

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:8px 12px;
            border-radius:20px;
            background-color:#D5F5E3;
            color:#196F3D;
            font-size:14px;
            font-weight:600;
            white-space:nowrap;
        ">
            🟢 Live Monitoring
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header(
    "Filters"
)


parameter = (
    st.sidebar
    .selectbox(
        "Parameter",
        [
            "Moisture",
            "Protein",
            "Ash",
        ],
    )
)


filtered = (
    df[
        df[
            "PARAMETER"
        ]
        == parameter
    ]
    .copy()
)


# ============================================================
# DATE RANGE
# ============================================================

valid_dates = (
    filtered[
        "DATE"
    ]
    .dropna()
)


if valid_dates.empty:

    st.warning(
        "No valid dates found."
    )

    st.stop()


min_date = (
    valid_dates
    .min()
    .date()
)


max_date = (
    valid_dates
    .max()
    .date()
)


selected_dates = (
    st.sidebar
    .date_input(
        "Date Range",
        value=(
            min_date,
            max_date,
        ),
        min_value=min_date,
        max_value=max_date,
    )
)


if (
    isinstance(
        selected_dates,
        (tuple, list),
    )
    and
    len(selected_dates) == 2
):

    start_date = (
        pd.Timestamp(
            selected_dates[0]
        )
    )


    end_date = (
        pd.Timestamp(
            selected_dates[1]
        )
        + pd.Timedelta(
            days=1
        )
        - pd.Timedelta(
            microseconds=1
        )
    )


    filtered = (
        filtered[
            (
                filtered[
                    "DATE"
                ]
                >= start_date
            )
            &
            (
                filtered[
                    "DATE"
                ]
                <= end_date
            )
        ]
        .copy()
    )


# ============================================================
# BRAND FILTER
# ============================================================

if "BRAND" in filtered.columns:

    brands = sorted(
        filtered[
            "BRAND"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )


    selected_brands = (
        st.sidebar
        .multiselect(
            "Brand",
            brands,
        )
    )


    if selected_brands:

        filtered = (
            filtered[
                filtered[
                    "BRAND"
                ]
                .astype(str)
                .isin(
                    selected_brands
                )
            ]
            .copy()
        )


# ============================================================
# PROCESS FILTER
# ============================================================

if "PROCESS" in filtered.columns:

    processes = sorted(
        filtered[
            "PROCESS"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )


    selected_process = (
        st.sidebar
        .multiselect(
            "Process",
            processes,
        )
    )


    if selected_process:

        filtered = (
            filtered[
                filtered[
                    "PROCESS"
                ]
                .astype(str)
                .isin(
                    selected_process
                )
            ]
            .copy()
        )


# ============================================================
# CURVE FILTER
# ============================================================

if (
    "CURVE" in filtered.columns
    and
    parameter in [
        "Moisture",
        "Protein",
    ]
):

    curves = sorted(
        filtered[
            "CURVE"
        ]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )


    selected_curve = (
        st.sidebar
        .multiselect(
            "Curve",
            curves,
        )
    )


    if selected_curve:

        filtered = (
            filtered[
                filtered[
                    "CURVE"
                ]
                .astype(str)
                .isin(
                    selected_curve
                )
            ]
            .copy()
        )


# ============================================================
# NO DATA
# ============================================================

if filtered.empty:

    st.warning(
        "No records found for the selected filters."
    )

    st.stop()


# ============================================================
# DATA DISPLAYED UP TO
# ============================================================

latest_data = (
    filtered[
        "DATETIME"
    ]
    .dropna()
)


if not latest_data.empty:

    latest_datetime = (
        latest_data.max()
    )


    latest_rows = (
        filtered[
            filtered[
                "DATETIME"
            ]
            == latest_datetime
        ]
    )


    has_real_time = False


    if (
        "TIME_DISPLAY"
        in latest_rows.columns
    ):

        has_real_time = (
            latest_rows[
                "TIME_DISPLAY"
            ]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
            .any()
        )


    if has_real_time:

        latest_display = (
            latest_datetime
            .strftime(
                "%d %b %Y, %I:%M %p"
            )
        )

    else:

        latest_display = (
            latest_datetime
            .strftime(
                "%d %b %Y"
            )
        )


    st.markdown(
        f"""
        <div style="
            margin-top:4px;
            margin-bottom:14px;
            color:#555555;
            font-size:14px;
        ">
            📅 Data displayed up to:
            <strong>{latest_display}</strong>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# TREND VIEW + COMPARISON
# ============================================================

control1, control2 = (
    st.columns(2)
)


with control1:

    view = st.radio(
        "Trend View",
        [
            "Daily",
            "Weekly",
            "Monthly",
        ],
        horizontal=True,
    )


if parameter == "Ash":

    lab_name = "NIT"
    tolerance = 0.02

else:

    lab_name = "NOVA"
    tolerance = 0.20


comparison_options = {

    f"{lab_name} vs CONV":
        "LAB vs CONV",

    "DA vs CONV":
        "DA vs CONV",

    f"DA vs {lab_name}":
        "DA vs LAB",
}


with control2:

    selected_label = (
        st.selectbox(
            "Comparison",
            list(
                comparison_options.keys()
            ),
        )
    )


comparison = (
    comparison_options[
        selected_label
    ]
)


status_column_map = {

    "LAB vs CONV":
        "LAB vs CONV STATUS",

    "DA vs CONV":
        "DA vs CONV STATUS",

    "DA vs LAB":
        "DA vs LAB STATUS",
}


status_column = (
    status_column_map[
        comparison
    ]
)


# ============================================================
# KPI
# ============================================================

total_records = (
    len(filtered)
)


status_series = (
    clean_status(
        filtered[
            status_column
        ]
    )
)


status_lower = (
    status_series
    .str.lower()
)


within_tolerance = int(
    (
        status_lower
        == "within tolerance"
    )
    .sum()
)


out_tolerance = int(
    (
        status_lower
        == "out of tolerance"
    )
    .sum()
)


blank_count = (
    total_records
    - within_tolerance
    - out_tolerance
)


available_comparisons = (
    within_tolerance
    + out_tolerance
)


compliance = (

    (
        within_tolerance
        / available_comparisons
    )
    * 100

    if available_comparisons > 0

    else 0
)


difference_series = (
    pd.to_numeric(
        filtered[
            comparison
        ],
        errors="coerce",
    )
)


avg_abs_difference = (
    difference_series
    .abs()
    .mean()
)


# ============================================================
# KPI CARDS
# ============================================================

k1, k2, k3, k4, k5, k6 = (
    st.columns(6)
)


k1.metric(
    "Total Records",
    f"{total_records:,}",
)


k2.metric(
    "🟢 Within Tolerance",
    f"{within_tolerance:,}",
)


k3.metric(
    "🔴 Out of Tolerance",
    f"{out_tolerance:,}",
)


k4.metric(
    "⚪ Blank",
    f"{blank_count:,}",
)


k5.metric(
    "Compliance",
    f"{compliance:.1f}%",
)


if pd.isna(
    avg_abs_difference
):

    avg_display = "-"


elif parameter == "Ash":

    avg_display = (
        f"{avg_abs_difference:.3f}"
    )


else:

    avg_display = (
        f"{avg_abs_difference:.2f}"
    )


k6.metric(
    "Avg |Difference|",
    avg_display,
)


st.caption(
    f"{selected_label} | "
    f"Tolerance ±{tolerance:.2f} | "
    f"Available comparisons: "
    f"{available_comparisons:,}"
)


st.divider()


# ============================================================
# PERFORMANCE DATA
# ============================================================

comparison_data = (
    filtered[
        filtered[
            comparison
        ]
        .notna()
    ]
    .copy()
)


# ============================================================
# DAILY PERFORMANCE
# ============================================================

if view == "Daily":

    st.subheader(
        f"Daily {selected_label} Trend"
    )


    if comparison_data.empty:

        st.info(
            "No daily comparison data available."
        )

    else:

        comparison_data[
            "SOURCE_STATUS"
        ] = clean_status(
            comparison_data[
                status_column
            ]
        )


        comparison_data[
            "STATUS_LOWER"
        ] = (
            comparison_data[
                "SOURCE_STATUS"
            ]
            .str.lower()
        )


        # DAILY AVERAGE
        daily_avg = (
            comparison_data
            .assign(
                DAY=(
                    comparison_data[
                        "DATE"
                    ]
                    .dt
                    .normalize()
                )
            )
            .groupby(
                "DAY",
                as_index=False,
            )
            .agg(
                Average_Difference=(
                    comparison,
                    "mean",
                ),
                Reading_Count=(
                    comparison,
                    "count",
                ),
            )
            .sort_values(
                "DAY"
            )
        )


        fig = go.Figure()


        fig.add_hrect(
            y0=-tolerance,
            y1=tolerance,
            fillcolor=(
                "rgba(46, 204, 113, 0.13)"
            ),
            line_width=0,
            annotation_text=(
                "Within Tolerance"
            ),
            annotation_position=(
                "top left"
            ),
        )


        # WITHIN
        within_points = (
            comparison_data[
                comparison_data[
                    "STATUS_LOWER"
                ]
                == "within tolerance"
            ]
            .copy()
        )


        if not within_points.empty:

            fig.add_trace(
                go.Scatter(
                    x=within_points[
                        "DATETIME"
                    ],
                    y=within_points[
                        comparison
                    ],
                    mode="markers",
                    name="Within Tolerance",
                    marker=dict(
                        size=8,
                        color=GREEN,
                        opacity=0.70,
                    ),
                    hovertemplate=(
                        "<b>%{x|%d %b %Y %H:%M}</b><br>"
                        f"{selected_label}: "
                        "%{y:.3f}<br>"
                        "<b>Within Tolerance</b>"
                        "<extra></extra>"
                    ),
                )
            )


        # OOT
        oot_points = (
            comparison_data[
                comparison_data[
                    "STATUS_LOWER"
                ]
                == "out of tolerance"
            ]
            .copy()
        )


        if not oot_points.empty:

            fig.add_trace(
                go.Scatter(
                    x=oot_points[
                        "DATETIME"
                    ],
                    y=oot_points[
                        comparison
                    ],
                    mode="markers",
                    name="Out of Tolerance",
                    marker=dict(
                        size=12,
                        symbol="diamond",
                        color=RED,
                        line=dict(
                            width=1,
                            color=RED_DARK,
                        ),
                    ),
                    hovertemplate=(
                        "<b>%{x|%d %b %Y %H:%M}</b><br>"
                        f"{selected_label}: "
                        "%{y:.3f}<br>"
                        "<b>OUT OF TOLERANCE</b>"
                        "<extra></extra>"
                    ),
                )
            )


        # DAILY AVERAGE
        fig.add_trace(
            go.Scatter(
                x=daily_avg[
                    "DAY"
                ],
                y=daily_avg[
                    "Average_Difference"
                ],
                mode="lines+markers",
                name="Daily Average",
                line=dict(
                    width=3,
                    color=BLUE,
                ),
                marker=dict(
                    size=8,
                    color=BLUE,
                ),
                customdata=(
                    daily_avg[
                        ["Reading_Count"]
                    ]
                ),
                hovertemplate=(
                    "<b>%{x|%d %b %Y}</b><br>"
                    "Daily Average: %{y:.3f}<br>"
                    "Readings: %{customdata[0]}"
                    "<extra></extra>"
                ),
            )
        )


        fig.add_hline(
            y=0,
            line_dash="dot",
            line_color="#555555",
            annotation_text="Target = 0",
        )


        fig.add_hline(
            y=tolerance,
            line_dash="dash",
            line_color=RED,
            annotation_text=(
                f"+{tolerance:.2f}"
            ),
        )


        fig.add_hline(
            y=-tolerance,
            line_dash="dash",
            line_color=RED,
            annotation_text=(
                f"-{tolerance:.2f}"
            ),
        )


        # EVERY DAY
        fig.update_xaxes(
            dtick=(
                24
                * 60
                * 60
                * 1000
            ),
            tickformat="%d %b",
            tickangle=-45,
        )


        fig.update_layout(
            height=520,
            xaxis_title="Date",
            yaxis_title="Difference",
            hovermode="closest",
            legend=dict(
                orientation="h",
                y=1.08,
            ),
            margin=dict(
                b=90
            ),
        )


        st.plotly_chart(
            fig,
            use_container_width=True,
        )


# ============================================================
# WEEKLY / MONTHLY PERFORMANCE
# ============================================================

else:

    st.subheader(
        f"{view} {selected_label} Performance"
    )


    if comparison_data.empty:

        st.info(
            f"No {view.lower()} comparison data available."
        )

    else:

        aggregate_data = (
            comparison_data.copy()
        )


        aggregate_data[
            "PERIOD"
        ] = create_period(
            aggregate_data[
                "DATE"
            ],
            view,
        )


        aggregate_summary = (
            aggregate_data
            .groupby(
                "PERIOD",
                as_index=False,
            )
            .agg(
                Average_Difference=(
                    comparison,
                    "mean",
                ),
                Readings=(
                    comparison,
                    "count",
                ),
            )
        )


        aggregate_summary[
            "PERIOD_LABEL"
        ] = (
            aggregate_summary[
                "PERIOD"
            ]
            .apply(
                lambda x:
                make_period_label(
                    x,
                    view,
                )
            )
        )


        aggregate_summary[
            "Status"
        ] = np.where(
            aggregate_summary[
                "Average_Difference"
            ]
            .abs()
            <= tolerance,

            "Within Tolerance",

            "Out of Tolerance",
        )


        fig = px.bar(
            aggregate_summary,
            x="PERIOD_LABEL",
            y="Average_Difference",
            color="Status",
            text="Average_Difference",
            custom_data=[
                "Readings"
            ],
            color_discrete_map={
                "Within Tolerance":
                    GREEN,

                "Out of Tolerance":
                    RED,
            },
        )


        fig.update_traces(
            texttemplate=(
                "%{y:.3f}"
            ),
            textposition="outside",
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Average Difference: %{y:.3f}<br>"
                "Readings: %{customdata[0]}"
                "<extra></extra>"
            ),
        )


        fig.add_hrect(
            y0=-tolerance,
            y1=tolerance,
            fillcolor=(
                "rgba(46, 204, 113, 0.10)"
            ),
            line_width=0,
        )


        fig.add_hline(
            y=0,
            line_dash="dot",
            line_color="#555555",
        )


        fig.add_hline(
            y=tolerance,
            line_dash="dash",
            line_color=RED,
        )


        fig.add_hline(
            y=-tolerance,
            line_dash="dash",
            line_color=RED,
        )


        fig.update_xaxes(
            type="category",
            categoryorder="array",
            categoryarray=(
                aggregate_summary[
                    "PERIOD_LABEL"
                ]
                .tolist()
            ),
        )


        fig.update_layout(
            height=450,
            xaxis_title=(
                "Week"
                if view == "Weekly"
                else "Month"
            ),
            yaxis_title=(
                "Average Difference"
            ),
            legend_title_text=(
                "Status"
            ),
        )


        st.plotly_chart(
            fig,
            use_container_width=True,
        )


# ============================================================
# COMPLIANCE + TOP OOT
# ============================================================

g1, g2 = (
    st.columns(2)
)


# ============================================================
# COMPLIANCE TREND
# ============================================================

with g1:

    st.subheader(
        f"{view} Compliance Trend"
    )


    compliance_df = (
        filtered.copy()
    )


    compliance_df[
        "STATUS_CLEAN"
    ] = clean_status(
        compliance_df[
            status_column
        ]
    )


    compliance_df[
        "STATUS_LOWER"
    ] = (
        compliance_df[
            "STATUS_CLEAN"
        ]
        .str.lower()
    )


    compliance_df = (
        compliance_df[
            compliance_df[
                "STATUS_LOWER"
            ]
            .isin(
                [
                    "within tolerance",
                    "out of tolerance",
                ]
            )
        ]
        .copy()
    )


    if compliance_df.empty:

        st.info(
            "No compliance data available."
        )

    else:

        compliance_df[
            "PERIOD"
        ] = create_period(
            compliance_df[
                "DATE"
            ],
            view,
        )


        compliance_df[
            "IS_WITHIN"
        ] = (
            compliance_df[
                "STATUS_LOWER"
            ]
            == "within tolerance"
        )


        compliance_summary = (
            compliance_df
            .groupby(
                "PERIOD",
                as_index=False,
            )
            .agg(
                Compliance=(
                    "IS_WITHIN",
                    "mean",
                ),
                Total=(
                    "IS_WITHIN",
                    "size",
                ),
            )
        )


        compliance_summary[
            "Compliance"
        ] *= 100


        compliance_summary[
            "PERIOD_LABEL"
        ] = (
            compliance_summary[
                "PERIOD"
            ]
            .apply(
                lambda x:
                make_period_label(
                    x,
                    view,
                )
            )
        )


        fig2 = go.Figure()


        fig2.add_trace(
            go.Scatter(
                x=compliance_summary[
                    "PERIOD_LABEL"
                ],
                y=compliance_summary[
                    "Compliance"
                ],
                mode="lines+markers",
                line=dict(
                    color=BLUE,
                    width=3,
                ),
                marker=dict(
                    size=9,
                    color=BLUE,
                ),
                customdata=(
                    compliance_summary[
                        ["Total"]
                    ]
                ),
                hovertemplate=(
                    "<b>%{x}</b><br>"
                    "Compliance: %{y:.1f}%<br>"
                    "Comparisons: %{customdata[0]}"
                    "<extra></extra>"
                ),
            )
        )


        fig2.add_hline(
            y=95,
            line_dash="dash",
            line_color=RED,
            annotation_text=(
                "95% Target"
            ),
        )


        fig2.update_yaxes(
            range=[
                0,
                100,
            ]
        )


        fig2.update_xaxes(
            type="category",
            categoryorder="array",
            categoryarray=(
                compliance_summary[
                    "PERIOD_LABEL"
                ]
                .tolist()
            ),
        )


        fig2.update_layout(
            height=350,
            xaxis_title=(
                "Date"
                if view == "Daily"
                else (
                    "Week"
                    if view == "Weekly"
                    else "Month"
                )
            ),
            yaxis_title=(
                "Compliance %"
            ),
            showlegend=False,
        )


        st.plotly_chart(
            fig2,
            use_container_width=True,
        )


# ============================================================
# TOP OOT BRANDS
# ============================================================

with g2:

    st.subheader(
        "Top Out of Tolerance Brands"
    )


    if "BRAND" not in filtered.columns:

        st.info(
            "Brand column unavailable."
        )

    else:

        oot_status = (
            clean_status(
                filtered[
                    status_column
                ]
            )
            .str.lower()
        )


        oot_brand_data = (
            filtered[
                oot_status
                == "out of tolerance"
            ]
            .copy()
        )


        if oot_brand_data.empty:

            st.success(
                "No Out of Tolerance readings."
            )

        else:

            oot_summary = (
                oot_brand_data
                .groupby(
                    "BRAND"
                )
                .size()
                .reset_index(
                    name="OOT Count"
                )
                .sort_values(
                    "OOT Count",
                    ascending=False,
                )
                .head(10)
            )


            fig3 = px.bar(
                oot_summary,
                x="OOT Count",
                y="BRAND",
                orientation="h",
                text="OOT Count",
                color_discrete_sequence=[
                    RED
                ],
            )


            fig3.update_yaxes(
                categoryorder=(
                    "total ascending"
                )
            )


            fig3.update_layout(
                height=350,
                showlegend=False,
                xaxis_title=(
                    "OOT Count"
                ),
                yaxis_title="Brand",
            )


            st.plotly_chart(
                fig3,
                use_container_width=True,
            )


# ============================================================
# METHOD RESULT TREND
# ============================================================

st.subheader(
    f"{view} Method Result Trend"
)


# ============================================================
# DAILY METHOD
# ============================================================

if view == "Daily":

    daily_method = (
        filtered[
            [
                "DATETIME",
                "CONV",
                "LAB",
                "DA",
            ]
        ]
        .copy()
    )


    daily_long = (
        daily_method
        .melt(
            id_vars=[
                "DATETIME"
            ],
            value_vars=[
                "CONV",
                "LAB",
                "DA",
            ],
            var_name="Method",
            value_name="Result",
        )
    )


    daily_long = (
        daily_long
        .dropna(
            subset=[
                "DATETIME",
                "Result",
            ]
        )
    )


    daily_long[
        "Method"
    ] = (
        daily_long[
            "Method"
        ]
        .replace(
            {
                "LAB":
                    lab_name
            }
        )
    )


    if daily_long.empty:

        st.info(
            "No daily method result data available."
        )

    else:

        method_colors = {
            "CONV":
                CONV_COLOR,

            "NOVA":
                LAB_COLOR,

            "NIT":
                LAB_COLOR,

            "DA":
                DA_COLOR,
        }


        method_symbols = {
            "CONV":
                "diamond",

            "NOVA":
                "circle",

            "NIT":
                "circle",

            "DA":
                "square",
        }


        fig4 = px.scatter(
            daily_long,
            x="DATETIME",
            y="Result",
            color="Method",
            symbol="Method",
            color_discrete_map=(
                method_colors
            ),
            symbol_map=(
                method_symbols
            ),
        )


        fig4.update_traces(
            marker=dict(
                size=8,
                opacity=0.82,
                line=dict(
                    width=0.6,
                    color="white",
                ),
            )
        )


        fig4.update_xaxes(
            dtick=(
                24
                * 60
                * 60
                * 1000
            ),
            tickformat="%d %b",
            tickangle=-45,
        )


        fig4.update_layout(
            height=440,
            xaxis_title="Date",
            yaxis_title=(
                f"{parameter} Result"
            ),
            legend_title_text=(
                "Method"
            ),
            hovermode="closest",
            margin=dict(
                b=90
            ),
        )


        st.plotly_chart(
            fig4,
            use_container_width=True,
        )


# ============================================================
# WEEKLY / MONTHLY METHOD
# ============================================================

else:

    method_trend = (
        filtered.copy()
    )


    method_trend[
        "PERIOD"
    ] = create_period(
        method_trend[
            "DATE"
        ],
        view,
    )


    method_summary = (
        method_trend
        .groupby(
            "PERIOD",
            as_index=False,
        )
        .agg(
            CONV=(
                "CONV",
                "mean",
            ),
            LAB=(
                "LAB",
                "mean",
            ),
            DA=(
                "DA",
                "mean",
            ),
        )
    )


    method_summary[
        "PERIOD_LABEL"
    ] = (
        method_summary[
            "PERIOD"
        ]
        .apply(
            lambda x:
            make_period_label(
                x,
                view,
            )
        )
    )


    method_long = (
        method_summary
        .melt(
            id_vars=[
                "PERIOD_LABEL"
            ],
            value_vars=[
                "CONV",
                "LAB",
                "DA",
            ],
            var_name="Method",
            value_name=(
                "Average Result"
            ),
        )
    )


    method_long = (
        method_long
        .dropna(
            subset=[
                "Average Result"
            ]
        )
    )


    method_long[
        "Method"
    ] = (
        method_long[
            "Method"
        ]
        .replace(
            {
                "LAB":
                    lab_name
            }
        )
    )


    if method_long.empty:

        st.info(
            f"No {view.lower()} method result data available."
        )

    else:

        method_colors = {
            "CONV":
                CONV_COLOR,

            "NOVA":
                LAB_COLOR,

            "NIT":
                LAB_COLOR,

            "DA":
                DA_COLOR,
        }


        method_symbols = {
            "CONV":
                "diamond",

            "NOVA":
                "circle",

            "NIT":
                "circle",

            "DA":
                "square",
        }


        fig4 = px.line(
            method_long,
            x="PERIOD_LABEL",
            y="Average Result",
            color="Method",
            symbol="Method",
            markers=True,
            color_discrete_map=(
                method_colors
            ),
            symbol_map=(
                method_symbols
            ),
        )


        fig4.update_traces(
            line=dict(
                width=3
            ),
            marker=dict(
                size=9
            ),
        )


        fig4.update_xaxes(
            type="category",
        )


        fig4.update_layout(
            height=420,
            xaxis_title=(
                "Week"
                if view == "Weekly"
                else "Month"
            ),
            yaxis_title=(
                f"Average {parameter} Result"
            ),
            hovermode="x unified",
            legend_title_text=(
                "Method"
            ),
        )


        st.plotly_chart(
            fig4,
            use_container_width=True,
        )


# ============================================================
# DETAILED READINGS
# ============================================================

st.divider()


st.subheader(
    "Detailed Readings"
)


detail = (
    filtered.copy()
)


detail[
    "STATUS_DISPLAY"
] = clean_status(
    detail[
        status_column
    ]
)


detail_columns = [
    "SOURCE_ROW",
    "DATE",
    "TIME_DISPLAY",
    "BRAND",
    "CURVE",
    "PROCESS",
    "CONV",
    "LAB",
    "DA",
    comparison,
    "STATUS_DISPLAY",
]


detail_columns = [

    column

    for column
    in detail_columns

    if column
    in detail.columns
]


detail = (
    detail[
        detail_columns
    ]
    .copy()
)


detail = (
    detail.rename(
        columns={
            "SOURCE_ROW":
                "Excel Row",

            "DATE":
                "Date",

            "TIME_DISPLAY":
                "Time",

            "BRAND":
                "Brand",

            "CURVE":
                "Curve",

            "PROCESS":
                "Process",

            "LAB":
                lab_name,

            comparison:
                selected_label,

            "STATUS_DISPLAY":
                "Status",
        }
    )
)


if "Date" in detail.columns:

    detail[
        "Date"
    ] = (
        pd.to_datetime(
            detail[
                "Date"
            ]
        )
        .dt
        .strftime(
            "%d/%m/%Y"
        )
    )


styled_detail = (
    detail.style
    .map(
        highlight_status,
        subset=[
            "Status"
        ],
    )
)


st.dataframe(
    styled_detail,
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# OUT OF TOLERANCE DETAILS
# ============================================================

st.subheader(
    "Out of Tolerance Details"
)


oot_mask = (
    clean_status(
        filtered[
            status_column
        ]
    )
    .str.lower()
    == "out of tolerance"
)


oot_detail = (
    filtered[
        oot_mask
    ]
    .copy()
)


if oot_detail.empty:

    st.success(
        "No Out of Tolerance readings."
    )

else:

    oot_detail[
        "STATUS_DISPLAY"
    ] = clean_status(
        oot_detail[
            status_column
        ]
    )


    oot_columns = [
        "SOURCE_ROW",
        "DATE",
        "TIME_DISPLAY",
        "BRAND",
        "CURVE",
        "PROCESS",
        "CONV",
        "LAB",
        "DA",
        comparison,
        "STATUS_DISPLAY",
    ]


    oot_columns = [

        column

        for column
        in oot_columns

        if column
        in oot_detail.columns
    ]


    oot_show = (
        oot_detail[
            oot_columns
        ]
        .copy()
    )


    oot_show = (
        oot_show
        .rename(
            columns={
                "SOURCE_ROW":
                    "Excel Row",

                "DATE":
                    "Date",

                "TIME_DISPLAY":
                    "Time",

                "BRAND":
                    "Brand",

                "CURVE":
                    "Curve",

                "PROCESS":
                    "Process",

                "LAB":
                    lab_name,

                comparison:
                    selected_label,

                "STATUS_DISPLAY":
                    "Status",
            }
        )
    )


    if "Date" in oot_show.columns:

        oot_show[
            "Date"
        ] = (
            pd.to_datetime(
                oot_show[
                    "Date"
                ]
            )
            .dt
            .strftime(
                "%d/%m/%Y"
            )
        )


    styled_oot = (
        oot_show.style
        .map(
            highlight_status,
            subset=[
                "Status"
            ],
        )
    )


    st.dataframe(
        styled_oot,
        use_container_width=True,
        hide_index=True,
    )


    csv_data = (
        oot_show
        .to_csv(
            index=False
        )
        .encode(
            "utf-8-sig"
        )
    )


    st.download_button(
        "⬇️ Download OOT CSV",
        data=csv_data,
        file_name=(
            f"{parameter}_"
            f"{selected_label}_OOT.csv"
        ),
        mime="text/csv",
    )


# ============================================================
# DATA QUALITY & AUDIT
# ============================================================

with st.expander(
    "Data Quality & Audit"
):

    parameter_data = (
        df[
            df[
                "PARAMETER"
            ]
            == parameter
        ]
        .copy()
    )


    st.write(
        f"Source records for "
        f"{parameter}: "
        f"**{len(parameter_data):,}**"
    )


    st.write(
        "Records after current filters: "
        f"**{len(filtered):,}**"
    )


    st.write(
        f"🟢 {selected_label} — "
        "Within Tolerance: "
        f"**{within_tolerance:,}**"
    )


    st.write(
        f"🔴 {selected_label} — "
        "Out of Tolerance: "
        f"**{out_tolerance:,}**"
    )


    st.write(
        f"⚪ {selected_label} — "
        "Blank: "
        f"**{blank_count:,}**"
    )


    check_total = (
        within_tolerance
        + out_tolerance
        + blank_count
    )


    st.write(
        "Check: "
        f"**{within_tolerance:,} + "
        f"{out_tolerance:,} + "
        f"{blank_count:,} = "
        f"{check_total:,}**"
    )


    if (
        check_total
        == total_records
    ):

        st.success(
            "✅ Status count matches Total Records."
        )

    else:

        st.error(
            "❌ Status count does not match Total Records."
        )


    if not latest_data.empty:

        st.write(
            "Latest displayed record: "
            f"**{latest_display}**"
        )


    if not parameter_data.empty:

        st.write(
            "Excel source row range: "
            f"**"
            f"{int(parameter_data['SOURCE_ROW'].min())}"
            f" to "
            f"{int(parameter_data['SOURCE_ROW'].max())}"
            f"**"
        )


    st.caption(
        "Status colours: Green = Within Tolerance | "
        "Red = Out of Tolerance | "
        "Grey = Blank. "
        "Method colours: Purple = CONV | "
        "Orange = NOVA/NIT | "
        "Blue = DA."
    )
import re
import pathlib
import streamlit as st
import pandas as pd
import io
import plotly.express as px
from datetime import datetime, date, time, timedelta
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.chart import BarChart, Reference

st.set_page_config(
    page_title="Agent AUX Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────────────
# Styling
# ──────────────────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
        .block-container { padding-top: 1.5rem; }
        .main-title {
            font-size: 2rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }
        .sub-title {
            font-size: 0.95rem;
            opacity: 0.65;
            margin-bottom: 1.5rem;
        }
        .section-header {
            font-size: 1.2rem;
            font-weight: 600;
            border-bottom: 2px solid rgba(128,128,128,0.35);
            padding-bottom: 4px;
            margin-top: 1rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="main-title">📊 CS Agent AUX Code Analyzer</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Dynamically explore which AUX codes any agent was on during any time slot, '
    "with exact overlap durations and visual breakdowns. "
    "<strong>All dates &amp; times are displayed in IST (UTC+5:30).</strong></div>",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def fmt_duration(seconds: float) -> str:
    """Format seconds to HH:MM:SS string."""
    seconds = int(max(0, seconds))
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Case-insensitive column lookup from a list of candidate names."""
    lookup = {c.lower().strip().replace(" ", ""): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().replace(" ", "")
        if key in lookup:
            return lookup[key]
    return None


def parse_time_slot(slot_str: str):
    """Parse slot strings like '6 PM-7PM', '10 PM to 11 PM', '8PM to 9 PM'.
    Returns (start_time, end_time) or None if unparseable."""
    s = str(slot_str).strip()
    parts = re.split(r'\s*(?:\bto\b|-)\s*', s, maxsplit=1, flags=re.IGNORECASE)
    if len(parts) != 2:
        return None

    def _t(ts: str):
        ts = ts.strip()
        m = re.match(r'(\d{1,2})(?::(\d{2}))?\s*(AM|PM)', ts, re.IGNORECASE)
        if not m:
            return None
        h, mi, period = int(m.group(1)), int(m.group(2) or 0), m.group(3).upper()
        if period == "PM" and h != 12:
            h += 12
        if period == "AM" and h == 12:
            h = 0
        return time(h, mi)

    s_time, e_time = _t(parts[0]), _t(parts[1])
    return (s_time, e_time) if s_time and e_time else None


# ── Disk-based file cache (survives new browser sessions) ───────────────────
_CACHE_DIR = pathlib.Path(".file_cache")
_CACHE_DIR.mkdir(exist_ok=True)


def _save_to_disk(key: str, file_bytes: bytes, file_name: str) -> None:
    (_CACHE_DIR / f"{key}.bin").write_bytes(file_bytes)
    (_CACHE_DIR / f"{key}.name").write_text(file_name, encoding="utf-8")


def _load_from_disk(key: str):
    b_path = _CACHE_DIR / f"{key}.bin"
    n_path = _CACHE_DIR / f"{key}.name"
    if b_path.exists() and n_path.exists():
        return b_path.read_bytes(), n_path.read_text(encoding="utf-8")
    return None, None


def _clear_from_disk(key: str) -> None:
    for suffix in (".bin", ".name"):
        p = _CACHE_DIR / f"{key}{suffix}"
        if p.exists():
            p.unlink()


# ── Excel formatting helpers ──────────────────────────────────────────────────
def _style_ws(ws, pct_col="Compliance %", has_total_row=False):
    HDR_FILL   = PatternFill("solid", fgColor="2E75B6")
    ALT_FILL   = PatternFill("solid", fgColor="DCE6F1")
    HDR_FONT   = Font(bold=True, color="FFFFFF", size=10)
    GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
    RED_FILL   = PatternFill("solid", fgColor="FFC7CE")
    GREEN_FONT = Font(color="276221", bold=True)
    RED_FONT   = Font(color="9C0006", bold=True)
    max_row  = ws.max_row
    data_end = max_row - 1 if has_total_row else max_row
    for cell in ws[1]:
        cell.font      = HDR_FONT
        cell.fill      = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    for r in range(2, data_end + 1):
        for cell in ws[r]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if r % 2 == 0:
                cell.fill = ALT_FILL
    for col in ws.iter_cols(min_row=1, max_row=max_row):
        cl = get_column_letter(col[0].column)
        w  = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[cl].width = min(max(w + 3, 10), 42)
    ws.freeze_panes = "A2"
    pct_cl = None
    if pct_col:
        for cell in ws[1]:
            if cell.value == pct_col:
                pct_cl = get_column_letter(cell.column)
                break
        if pct_cl:
            rng = f"{pct_cl}2:{pct_cl}{data_end}"
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="greaterThanOrEqual", formula=["0.7"],
                                fill=GREEN_FILL, font=GREEN_FONT))
            ws.conditional_formatting.add(
                rng, CellIsRule(operator="lessThan", formula=["0.7"],
                                fill=RED_FILL, font=RED_FONT))
            for r in range(2, max_row + 1):
                ws[f"{pct_cl}{r}"].number_format = "0.0%"
    if has_total_row and max_row >= 2:
        for cell in ws[max_row]:
            cell.font      = HDR_FONT
            cell.fill      = HDR_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[max_row].height = 22


def _add_bar_chart(ws, name_col, title, exclude_last=False):
    hdr = [cell.value for cell in ws[1]]
    if "Compliance %" not in hdr or name_col not in hdr:
        return
    pct_idx  = hdr.index("Compliance %") + 1
    name_idx = hdr.index(name_col) + 1
    max_row  = ws.max_row - 1 if exclude_last else ws.max_row
    chart = BarChart()
    chart.type      = "bar"
    chart.grouping  = "clustered"
    chart.title     = title
    chart.style     = 2
    chart.height    = max(8, min(20, (max_row - 1) * 0.6 + 5))
    chart.width     = 18
    chart.y_axis.title = name_col
    chart.x_axis.title = "Compliance %"
    data_ref = Reference(ws, min_col=pct_idx, min_row=1, max_row=max_row)
    cats_ref = Reference(ws, min_col=name_idx, min_row=2, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    anchor = get_column_letter(ws.max_column + 2) + "2"
    ws.add_chart(chart, anchor)


WORKING_AUX: set[str] = {
    "outboundpending",
    "callbackpending",
    "consultpending",
    "inboundpending",
    "outbond",
    "outbound",
    "acw",
    "in call - working",
    "in call- working",
    "e-mails",
    "emails",
    "follow-up case work",
    "followup case work",
    "followup",
    "follow-up",
}


def is_working_aux(label: str) -> bool:
    """Return True if the AUX label represents active/working activity."""
    return label.lower().strip() in WORKING_AUX


@st.cache_data(show_spinner="Loading file…")
def load_file(file_bytes: bytes, file_name: str) -> pd.DataFrame:
    if file_name.lower().endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))
    return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")


def parse_datetimes(df: pd.DataFrame, login_col: str, logout_col: str) -> pd.DataFrame:
    df = df.copy()
    df[login_col]  = pd.to_datetime(df[login_col],  errors="coerce", dayfirst=False)
    df[logout_col] = pd.to_datetime(df[logout_col], errors="coerce", dayfirst=False)
    bad = df[login_col].isna() | df[logout_col].isna() | (df[login_col] > df[logout_col])
    if bad.any():
        st.sidebar.warning(f"⚠️ Skipping {bad.sum()} rows with invalid/missing datetimes.")
    df = df[~bad].reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────
# Sidebar – File upload & column mapping
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("📁 Data Source")
    uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "xls", "csv"])

    if uploaded is not None:
        _b = uploaded.read()
        _save_to_disk("login", _b, uploaded.name)
        st.session_state["login_bytes"] = _b
        st.session_state["login_name"]  = uploaded.name

    if "login_bytes" not in st.session_state:
        _b, _n = _load_from_disk("login")
        if _b is not None:
            st.session_state["login_bytes"] = _b
            st.session_state["login_name"]  = _n

    if "login_bytes" not in st.session_state:
        st.info("Upload a file to begin.")
        st.stop()

    st.caption(f"📄 **{st.session_state['login_name']}**")
    if st.button("🗑️ Clear login file", key="clear_login"):
        _clear_from_disk("login")
        st.session_state.pop("login_bytes", None)
        st.session_state.pop("login_name",  None)
        st.rerun()

    raw_df = load_file(st.session_state["login_bytes"], st.session_state["login_name"])
    st.success(f"✅ {len(raw_df):,} rows loaded")
    st.info("🕐 UTC → IST (+5:30) conversion applied")

    st.divider()
    st.subheader("🗂️ Column Mapping")
    st.caption("Auto-detected from common names – adjust if needed.")

    cols = list(raw_df.columns)

    def _idx(col): return cols.index(col) if col else 0

    agent_default  = find_col(raw_df, ["agent name", "agentname", "agent"])
    login_default  = find_col(raw_df, ["login date", "logindate", "start time", "starttime", "start date", "startdate"])
    logout_default = find_col(raw_df, ["logout date", "logoutdate", "end time", "endtime", "end date", "enddate"])
    aux_default    = find_col(raw_df, ["unavailable code", "unavailablecode", "aux code", "auxcode", "status code"])
    dur_default    = find_col(raw_df, ["duration"])

    agent_col  = st.selectbox("Agent Name",          cols, index=_idx(agent_default))
    login_col  = st.selectbox("Login / Start Time",  cols, index=_idx(login_default))
    logout_col = st.selectbox("Logout / End Time",   cols, index=_idx(logout_default))
    aux_col    = st.selectbox("AUX / Unavailable Code (optional)", ["— skip —"] + cols,
                               index=(_idx(aux_default) + 1) if aux_default else 0)
    if aux_col == "— skip —":
        aux_col = None

    skill_default = find_col(raw_df, ["skill name", "skillname", "skill", "queue name", "queue"])
    skill_col  = st.selectbox("Skill Name Column (optional)", ["— skip —"] + cols,
                               index=(_idx(skill_default) + 1) if skill_default else 0)
    if skill_col == "— skip —":
        skill_col = None

    avail_default = find_col(raw_df, ["available time", "availabletime", "available", "avail time", "avail"])
    avail_col = st.selectbox("Available Time Column (optional)", ["— skip —"] + cols,
                              index=(_idx(avail_default) + 1) if avail_default else 0)
    if avail_col == "— skip —":
        avail_col = None

# ──────────────────────────────────────────────────────────────────────────────
# Parse datetimes
# ──────────────────────────────────────────────────────────────────────────────
try:
    df = parse_datetimes(raw_df, login_col, logout_col)
except Exception as exc:
    st.error(f"Could not parse datetime columns: {exc}")
    st.stop()

if df.empty:
    st.error("No valid rows remain after parsing datetimes. Check your column mapping.")
    st.stop()

# ── UTC → IST conversion (UTC+5:30) ──────────────────────────────────────────
IST_OFFSET = timedelta(hours=5, minutes=30)
df[login_col]  = df[login_col]  + IST_OFFSET
df[logout_col] = df[logout_col] + IST_OFFSET

# ── Derive AUX label ─────────────────────────────────────────────────────────
def _is_nonzero(val):
    if pd.isna(val):
        return False
    s = str(val).strip()
    if not s or s.lower() in ("0", "nan", "-", "00:00:00", "00:00", "0:00:00"):
        return False
    try:
        return float(s) > 0
    except ValueError:
        pass
    try:
        return any(int(p) > 0 for p in s.split(":"))
    except Exception:
        return bool(s)


def derive_aux_label(row):
    aux_val   = str(row[aux_col]).strip()   if aux_col   and pd.notna(row[aux_col])   and str(row[aux_col]).strip()   != "" else ""
    skill_val = str(row[skill_col]).strip() if skill_col and pd.notna(row[skill_col]) and str(row[skill_col]).strip() != "" else ""
    if aux_val:
        return aux_val
    elif skill_val:
        return "In Call - Working"
    elif avail_col and _is_nonzero(row.get(avail_col)):
        return "Available"
    else:
        return "Unavailable - Not Working"

df["_aux_label"] = df.apply(derive_aux_label, axis=1)

# ──────────────────────────────────────────────────────────────────────────────
# Query controls
# ──────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">🔍 Query Parameters</div>', unsafe_allow_html=True)
st.write("")

qc1, qc2, qc3, qc4 = st.columns([3, 2, 1.5, 1.5])

with qc1:
    agents = sorted(df[agent_col].dropna().unique().tolist())
    sel_agent = st.selectbox("👤 Agent", agents)

agent_df = df[df[agent_col] == sel_agent].copy()

with qc2:
    avail_dates = sorted(agent_df[login_col].dt.date.unique())
    sel_date = st.selectbox(
        "📅 Date (IST)",
        avail_dates,
        format_func=lambda d: d.strftime("%a, %b %d %Y"),
    )

with qc3:
    slot_start = st.time_input("⏰ From (IST)", value=time(13, 0))

with qc4:
    slot_end = st.time_input("⏰ To (IST)", value=time(14, 0))

if slot_start >= slot_end:
    st.error("'From' time must be earlier than 'To' time.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Overlap filtering logic
# ──────────────────────────────────────────────────────────────────────────────
q_start = datetime.combine(sel_date, slot_start)
q_end   = datetime.combine(sel_date, slot_end)

# All rows for that agent on that date
day_df = agent_df[agent_df[login_col].dt.date == sel_date].copy()

# Overlap condition: row interval intersects [q_start, q_end]
overlap_mask = (day_df[login_col] < q_end) & (day_df[logout_col] > q_start)
result_df = day_df[overlap_mask].copy()

if not result_df.empty:
    # Effective (clipped) times within the queried slot
    result_df["_eff_start"] = result_df[login_col].clip(lower=q_start, upper=q_end)
    result_df["_eff_end"]   = result_df[logout_col].clip(lower=q_start, upper=q_end)
    result_df["_overlap_sec"] = (
        result_df["_eff_end"] - result_df["_eff_start"]
    ).dt.total_seconds()
    result_df["_overlap_min"] = result_df["_overlap_sec"] / 60

# ──────────────────────────────────────────────────────────────────────────────
# Display results
# ──────────────────────────────────────────────────────────────────────────────
st.divider()
st.markdown(
    f'<div class="section-header">📋 Results — {sel_agent}</div>',
    unsafe_allow_html=True,
)
st.caption(
    f"Time slot (IST): **{slot_start.strftime('%H:%M')}** → **{slot_end.strftime('%H:%M')}** "
    f"on **{sel_date.strftime('%A, %B %d, %Y')}**"
)

if result_df.empty:
    st.warning("No AUX records overlap the selected time slot for this agent on this date.")
else:
    slot_total_sec = (q_end - q_start).total_seconds()
    total_overlap_sec = result_df["_overlap_sec"].sum()

    # ── Metrics ──────────────────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Matching Records", len(result_df))
    m2.metric("Total AUX Time in Slot", fmt_duration(total_overlap_sec))
    m3.metric("Slot Duration", fmt_duration(slot_total_sec))
    m4.metric("Slot Coverage", f"{total_overlap_sec / slot_total_sec * 100:.1f}%")

    st.write("")

    # ── Detailed table ────────────────────────────────────────────────────────
    st.markdown("#### 📄 Detailed Records")

    show_df = result_df.drop(columns=["_eff_start", "_eff_end", "_overlap_sec", "_overlap_min"]).copy()
    show_df = show_df.rename(columns={"_aux_label": "AUX Label"})

    # Add human-readable overlap columns
    show_df.insert(
        show_df.columns.get_loc(login_col) + 1,
        "Slot Effective Start (IST)",
        result_df["_eff_start"].dt.strftime("%H:%M:%S"),
    )
    show_df.insert(
        show_df.columns.get_loc(logout_col) + 1,
        "Slot Effective End (IST)",
        result_df["_eff_end"].dt.strftime("%H:%M:%S"),
    )
    show_df["Time in Slot (HH:MM:SS)"] = result_df["_overlap_sec"].apply(fmt_duration)

    # Format datetime cols for readability and rename to show IST
    for dc in [login_col, logout_col]:
        show_df[dc] = show_df[dc].dt.strftime("%m/%d/%Y %H:%M:%S")
    show_df = show_df.rename(columns={
        login_col:  f"{login_col} (IST)",
        logout_col: f"{logout_col} (IST)",
    })

    st.dataframe(show_df, use_container_width=True, hide_index=True)

    # ── AUX Breakdown ─────────────────────────────────────────────────────────
    if aux_col or skill_col:
        st.write("")
        st.markdown("#### 📊 AUX Breakdown in Slot")

        summary = (
            result_df.groupby("_aux_label")["_overlap_sec"]
            .sum()
            .reset_index()
            .rename(columns={"_aux_label": "AUX Code", "_overlap_sec": "Seconds"})
            .sort_values("Seconds", ascending=False)
        )
        summary["Duration (HH:MM:SS)"] = summary["Seconds"].apply(fmt_duration)
        summary["Minutes"] = (summary["Seconds"] / 60).round(2)
        summary["% of Slot"] = (summary["Seconds"] / slot_total_sec * 100).round(1)

        sc1, sc2 = st.columns([1, 1.5])

        with sc1:
            st.dataframe(
                summary[["AUX Code", "Duration (HH:MM:SS)", "% of Slot"]],
                use_container_width=True,
                hide_index=True,
            )

        with sc2:
            pie_data = summary[summary["Seconds"] > 0]
            if not pie_data.empty:
                fig_pie = px.pie(
                    pie_data,
                    values="Seconds",
                    names="AUX Code",
                    title=f"AUX Distribution  {slot_start.strftime('%H:%M')}–{slot_end.strftime('%H:%M')} IST",
                    hole=0.42,
                    color_discrete_sequence=px.colors.qualitative.Set2,
                )
                fig_pie.update_traces(textinfo="label+percent", pull=[0.03] * len(pie_data))
                fig_pie.update_layout(
                    height=320,
                    margin=dict(t=40, b=10, l=10, r=10),
                    showlegend=False,
                )
                st.plotly_chart(fig_pie, use_container_width=True)

    # ── Timeline / Gantt ──────────────────────────────────────────────────────
    st.write("")
    st.markdown("#### ⏱️ Activity Timeline (within slot)")

    timeline_rows = []
    for _, row in result_df.iterrows():
        label = row["_aux_label"]
        timeline_rows.append(
            {
                "AUX": label,
                "Start": row["_eff_start"],
                "Finish": row["_eff_end"],
            }
        )

    tl_df = pd.DataFrame(timeline_rows)

    if not tl_df.empty:
        # Sort so stacking looks sensible
        tl_df = tl_df.sort_values("Start")
        fig_gantt = px.timeline(
            tl_df,
            x_start="Start",
            x_end="Finish",
            y="AUX",
            color="AUX",
            color_discrete_sequence=px.colors.qualitative.Set2,
            title=f"Agent Timeline: {sel_agent}  (IST)",
        )
        fig_gantt.update_xaxes(
            range=[q_start, q_end],
            tickformat="%H:%M",
            title="Time (IST)",
        )
        fig_gantt.update_yaxes(title="")
        fig_gantt.update_layout(
            height=max(280, 60 + 45 * tl_df["AUX"].nunique()),
            showlegend=False,
            margin=dict(t=40, b=30, l=10, r=10),
        )
        st.plotly_chart(fig_gantt, use_container_width=True)

    # ── Export ────────────────────────────────────────────────────────────────
    st.write("")
    st.markdown("#### 💾 Export")

    out_buf = io.BytesIO()
    _show_clean = show_df.loc[:, ~show_df.columns.str.startswith("Unnamed:")].copy()
    with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
        _show_clean.to_excel(writer, sheet_name="Filtered Records", index=False)
        _style_ws(writer.sheets["Filtered Records"], pct_col=None)
        if aux_col or skill_col:
            summary[["AUX Code", "Duration (HH:MM:SS)", "Minutes", "% of Slot"]].to_excel(
                writer, sheet_name="AUX Summary", index=False
            )
            _style_ws(writer.sheets["AUX Summary"], pct_col=None)

    file_label = (
        f"AUX_{sel_agent.replace(' ', '_').replace(',', '')}_{sel_date}"
        f"_{slot_start.strftime('%H%M')}-{slot_end.strftime('%H%M')}.xlsx"
    )
    st.download_button(
        "📥 Download Results (Excel)",
        data=out_buf.getvalue(),
        file_name=file_label,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ──────────────────────────────────────────────────────────────────────────────
# Full-day context (collapsible)
# ──────────────────────────────────────────────────────────────────────────────
with st.expander(f"🗓️ Full Day View — {sel_agent}  ({sel_date.strftime('%b %d, %Y')})", expanded=False):
    if day_df.empty:
        st.write("No records for this agent on the selected date.")
    else:
        day_show = day_df.copy()
        for dc in [login_col, logout_col]:
            day_show[dc] = day_show[dc].dt.strftime("%m/%d/%Y %H:%M:%S")
        day_show = day_show.rename(columns={
            login_col:    f"{login_col} (IST)",
            logout_col:   f"{logout_col} (IST)",
            "_aux_label": "AUX Label",
        })
        st.dataframe(day_show, use_container_width=True, hide_index=True)

        # Full-day timeline
        if aux_col or skill_col:
            fd_rows = []
            for _, row in day_df.iterrows():
                label = row["_aux_label"]
                fd_rows.append({"AUX": label, "Start": row[login_col], "Finish": row[logout_col]})

            fd_tl = pd.DataFrame(fd_rows).sort_values("Start")
            if not fd_tl.empty:
                fig_fd = px.timeline(
                    fd_tl,
                    x_start="Start",
                    x_end="Finish",
                    y="AUX",
                    color="AUX",
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    title="Full Day Activity (IST)",
                )
                fig_fd.update_xaxes(tickformat="%H:%M", title="Time (IST)")
                fig_fd.update_yaxes(title="")
                fig_fd.update_layout(
                    height=max(300, 60 + 45 * fd_tl["AUX"].nunique()),
                    showlegend=False,
                    margin=dict(t=40, b=30, l=10, r=10),
                )
                # Highlight the queried slot with a shaded rectangle
                fig_fd.add_vrect(
                    x0=q_start, x1=q_end,
                    fillcolor="rgba(21, 101, 192, 0.12)",
                    line_width=1,
                    line_color="rgba(21, 101, 192, 0.6)",
                    annotation_text=f"Queried slot",
                    annotation_position="top left",
                )
                st.plotly_chart(fig_fd, use_container_width=True)

# ──────────────────────────────────────────────────────────────────────────────
# Schedule Compliance Report
# ──────────────────────────────────────────────────────────────────────────────
st.divider()
st.markdown('<div class="section-header">📋 Schedule Compliance Report</div>', unsafe_allow_html=True)
st.write("")
st.caption(
    "Upload the agent schedule file (Excel/CSV). "
    "Agent names in the schedule must **exactly match** the Agent Name values in the login data. "
    "Expected day columns: Sun, Mon, Tue, Wed, Thu, Fri, Sat — values like '6 PM-7PM', "
    "'10 PM to 11 PM', or 'WO' (week off)."
)

sched_file = st.file_uploader(
    "Upload Schedule File (Excel / CSV)",
    type=["xlsx", "xls", "csv"],
    key="sched_upload",
)

if sched_file is not None:
    _sb = sched_file.read()
    _save_to_disk("sched", _sb, sched_file.name)
    st.session_state["sched_bytes"] = _sb
    st.session_state["sched_name"]  = sched_file.name

if "sched_bytes" not in st.session_state:
    _sb, _sn = _load_from_disk("sched")
    if _sb is not None:
        st.session_state["sched_bytes"] = _sb
        st.session_state["sched_name"]  = _sn

if "sched_bytes" in st.session_state:
    sc1, sc2 = st.columns([1, 3])
    if sc1.button("🗑️ Clear schedule file", key="clear_sched"):
        _clear_from_disk("sched")
        st.session_state.pop("sched_bytes", None)
        st.session_state.pop("sched_name",  None)
        st.rerun()
    sc2.caption(f"📄 Using: **{st.session_state['sched_name']}**")

    sched_raw = load_file(st.session_state["sched_bytes"], st.session_state["sched_name"])
    st.success(f"Schedule loaded — {len(sched_raw)} agent rows")

    with st.expander("Preview Schedule Data", expanded=False):
        st.dataframe(sched_raw, use_container_width=True)

    st.markdown("##### Configure Schedule Columns")
    sched_cols = list(sched_raw.columns)
    cfg1, cfg2, cfg3 = st.columns([2, 2, 3])

    with cfg1:
        sched_agent_default = find_col(sched_raw, ["agent name", "agentname", "agent"])
        sched_agent_col = st.selectbox(
            "Agent Name column (in schedule)",
            sched_cols,
            index=sched_cols.index(sched_agent_default) if sched_agent_default else 0,
            key="sched_agent_col",
        )

    with cfg2:
        sched_sup_default = find_col(sched_raw, [
            "supervisor name", "supervisorname", "supervisor",
            "team lead", "teamlead", "tl", "manager", "reporting to", "reports to",
        ])
        _sup_opts = ["— skip —"] + sched_cols
        sched_sup_col = st.selectbox(
            "Supervisor column (optional)",
            _sup_opts,
            index=(_sup_opts.index(sched_sup_default)
                   if sched_sup_default and sched_sup_default in _sup_opts else 0),
            key="sched_sup_col",
        )
        if sched_sup_col == "— skip —":
            sched_sup_col = None

    with cfg3:
        auto_days = [
            c for c in sched_cols
            if c.strip()[:3].lower() in ("sun", "mon", "tue", "wed", "thu", "fri", "sat")
        ]
        day_cols_sel = st.multiselect(
            "Day-of-week columns (select all that apply)",
            sched_cols,
            default=auto_days,
            key="day_cols_sel",
        )

    # Date range to analyse
    all_dates = sorted(df[login_col].dt.date.unique())
    dr1, dr2 = st.columns(2)
    with dr1:
        date_from = st.date_input("From Date (IST)", value=min(all_dates), key="sched_from")
    with dr2:
        date_to = st.date_input("To Date (IST)",   value=max(all_dates), key="sched_to")

    check_dates = [d for d in all_dates if date_from <= d <= date_to]

    if not check_dates:
        st.warning("No dates in the login data fall within the selected range.")
    else:
        st.info(
            f"Checking **{len(check_dates)} date(s)** × "
            f"**{len(sched_raw)} agent row(s)** = "
            f"**{len(check_dates) * len(sched_raw)} slot checks**."
        )

        if st.button("📊 Generate Compliance Report", type="primary"):

            # Build day-of-week abbreviation → schedule column name
            _dow_norm = {
                "sun": "Sun", "mon": "Mon", "tue": "Tue",
                "wed": "Wed", "thu": "Thu", "fri": "Fri", "sat": "Sat",
            }
            dow_col_map = {}
            for col in day_cols_sel:
                prefix = col.strip()[:3].lower()
                if prefix in _dow_norm:
                    dow_col_map[_dow_norm[prefix]] = col

            # weekday() → English day abbreviation (locale-independent)
            _wd_to_dow = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

            records = []
            out_records = []
            total_ops = max(len(sched_raw) * len(check_dates), 1)
            prog = st.progress(0.0, text="Generating report…")
            done = 0

            for _, sched_row in sched_raw.iterrows():
                sched_agent = str(sched_row[sched_agent_col]).strip()
                if not sched_agent or sched_agent.lower() == "nan":
                    done += len(check_dates)
                    prog.progress(min(done / total_ops, 1.0))
                    continue

                sched_supervisor = (
                    str(sched_row[sched_sup_col]).strip()
                    if sched_sup_col and pd.notna(sched_row.get(sched_sup_col))
                    else ""
                )
                agent_login = df[df[agent_col] == sched_agent]

                for check_date in check_dates:
                    dow_short = _wd_to_dow[check_date.weekday()]

                    slot_str = "WO"
                    if dow_short in dow_col_map:
                        raw = sched_row.get(dow_col_map[dow_short], "WO")
                        slot_str = str(raw).strip() if pd.notna(raw) else "WO"

                    base = {
                        "Supervisor":     sched_supervisor,
                        "Agent Name":     sched_agent,
                        "Date (IST)":     check_date.strftime("%d-%b-%Y"),
                        "Day":            dow_short,
                        "Scheduled Slot": slot_str,
                    }

                    if slot_str.upper() in ("WO", "", "NAN", "NONE", "-"):
                        records.append({
                            **base,
                            "Scheduled Slot":       "Week Off",
                            "Status":               "📅 Week Off",
                            "In Call Duration":     "-",
                            "Total Active in Slot": "-",
                            "AUX Breakdown":        "-",
                        })
                        wo_rows = agent_login[agent_login[login_col].dt.date == check_date]
                        for _, lrow in wo_rows.iterrows():
                            if not is_working_aux(lrow["_aux_label"]):
                                continue
                            _ds = max(0, (lrow[logout_col] - lrow[login_col]).total_seconds())
                            if _ds > 0:
                                out_records.append({
                                    "Supervisor":          sched_supervisor,
                                    "Agent Name":          sched_agent,
                                    "Date (IST)":          check_date.strftime("%d-%b-%Y"),
                                    "Day":                 dow_short,
                                    "Scheduled Slot":      "Week Off",
                                    "Login Time (IST)":    lrow[login_col].strftime("%H:%M"),
                                    "Logout Time (IST)":   lrow[logout_col].strftime("%H:%M"),
                                    "AUX Code":            lrow["_aux_label"],
                                    "Duration":            fmt_duration(_ds),
                                    "Type":                "Week Off - Logged In",
                                })

                    else:
                        parsed = parse_time_slot(slot_str)

                        if parsed is None:
                            records.append({
                                **base,
                                "Status":               "⚠️ Slot Parse Error",
                                "In Call Duration":     "-",
                                "Total Active in Slot": "-",
                                "AUX Breakdown":        f"Cannot parse: '{slot_str}'",
                            })

                        else:
                            s_time, e_time = parsed
                            q_s = datetime.combine(check_date, s_time)
                            q_e = datetime.combine(check_date, e_time)

                            day_rows = agent_login[agent_login[login_col].dt.date == check_date]
                            overlap  = day_rows[
                                (day_rows[login_col] < q_e) & (day_rows[logout_col] > q_s)
                            ].copy()

                            for _, lrow in day_rows.iterrows():
                                if not is_working_aux(lrow["_aux_label"]):
                                    continue
                                _lt  = lrow[login_col]
                                _lo  = lrow[logout_col]
                                _tot = max(0, (_lo - _lt).total_seconds())
                                _es  = max(_lt, q_s)
                                _ee  = min(_lo, q_e)
                                _in  = max(0, (_ee - _es).total_seconds()) if _ee > _es else 0
                                _out = _tot - _in
                                if _out > 0:
                                    if _lt < q_s and _lo <= q_s:
                                        _otype = "Before Slot"
                                    elif _lt >= q_e:
                                        _otype = "After Slot"
                                    elif _lt < q_s and _lo > q_e:
                                        _otype = "Before & After Slot"
                                    elif _lt < q_s:
                                        _otype = "Partially Before Slot"
                                    else:
                                        _otype = "Partially After Slot"
                                    out_records.append({
                                        "Supervisor":           sched_supervisor,
                                        "Agent Name":           sched_agent,
                                        "Date (IST)":           check_date.strftime("%d-%b-%Y"),
                                        "Day":                  dow_short,
                                        "Scheduled Slot":       slot_str,
                                        "Login Time (IST)":     _lt.strftime("%H:%M"),
                                        "Logout Time (IST)":    _lo.strftime("%H:%M"),
                                        "AUX Code":             lrow["_aux_label"],
                                        "Out-of-Slot Duration": fmt_duration(_out),
                                        "Type":                 _otype,
                                    })

                            if overlap.empty:
                                records.append({
                                    **base,
                                    "Status":               "❌ Not Logged In",
                                    "In Call Duration":     "00:00:00",
                                    "Total Active in Slot": "00:00:00",
                                    "AUX Breakdown":        "No activity",
                                })
                            else:
                                overlap["_es"] = overlap[login_col].clip(lower=q_s, upper=q_e)
                                overlap["_ee"] = overlap[logout_col].clip(lower=q_s, upper=q_e)
                                overlap["_sc"] = (overlap["_ee"] - overlap["_es"]).dt.total_seconds()

                                total_secs   = overlap["_sc"].sum()
                                in_call_secs = overlap[
                                    overlap["_aux_label"] == "In Call - Working"
                                ]["_sc"].sum()

                                aux_grp = (
                                    overlap.groupby("_aux_label")["_sc"]
                                    .sum().sort_values(ascending=False)
                                )
                                aux_str = "; ".join(
                                    f"{k}: {fmt_duration(v)}" for k, v in aux_grp.items()
                                )

                                status = (
                                    "✅ Logged In - On Calls"
                                    if in_call_secs > 0
                                    else "⚠️ Logged In - No Calls"
                                )
                                records.append({
                                    **base,
                                    "Status":               status,
                                    "In Call Duration":     fmt_duration(in_call_secs),
                                    "Total Active in Slot": fmt_duration(total_secs),
                                    "AUX Breakdown":        aux_str,
                                })

                    done += 1
                    prog.progress(min(done / total_ops, 1.0))

            prog.empty()

            if not records:
                st.info(
                    "No records generated. "
                    "Verify that agent names in the schedule match the login data exactly."
                )
            else:
                report_df = pd.DataFrame(records)

                # ── Summary metrics ───────────────────────────────────────────
                non_wo    = report_df[~report_df["Status"].str.startswith("📅", na=False)]
                on_calls  = non_wo[non_wo["Status"].str.startswith("✅", na=False)]
                no_calls  = non_wo[non_wo["Status"].str.startswith("⚠️", na=False)]
                not_in    = non_wo[non_wo["Status"].str.startswith("❌", na=False)]
                week_off  = report_df[report_df["Status"].str.startswith("📅", na=False)]

                sm1, sm2, sm3, sm4, sm5 = st.columns(5)
                sm1.metric("Scheduled Slots",       len(non_wo))
                sm2.metric("✅ On Calls",             len(on_calls))
                sm3.metric("⚠️ Logged - No Calls",   len(no_calls))
                sm4.metric("❌ Not Logged In",        len(not_in))
                sm5.metric("📅 Week Off",             len(week_off))

                # ── Compliance % per agent ────────────────────────────────────
                if len(non_wo) > 0:
                    _stats = []
                    for _agent, _grp in non_wo.groupby("Agent Name"):
                        _tot = len(_grp)
                        _on  = int(_grp["Status"].str.startswith("✅").sum())
                        _nc  = int(_grp["Status"].str.startswith("⚠️").sum())
                        _ni  = int(_grp["Status"].str.startswith("❌").sum())
                        _stats.append({
                            "Agent Name":      _agent,
                            "Scheduled Days":  _tot,
                            "✅ On Calls":      _on,
                            "⚠️ No Calls":      _nc,
                            "❌ Not Logged In": _ni,
                            "Compliance %":    round(_on / _tot, 4) if _tot > 0 else 0.0,
                        })
                    agent_summary = pd.DataFrame(_stats)

                    st.markdown("##### Agent Summary")
                    _ag_disp = agent_summary.copy()
                    _ag_disp["Compliance %"] = _ag_disp["Compliance %"].apply(lambda x: f"{x:.1%}")
                    st.dataframe(_ag_disp, use_container_width=True, hide_index=True)

                # ── Supervisor summary ───────────────────────────────────────────
                sup_summary = pd.DataFrame()
                if sched_sup_col and len(non_wo) > 0 and "Supervisor" in non_wo.columns:
                    _ss = []
                    for _sup, _sg in non_wo.groupby("Supervisor"):
                        _stot = len(_sg)
                        _son  = int(_sg["Status"].str.startswith("✅").sum())
                        _snc  = int(_sg["Status"].str.startswith("⚠️").sum())
                        _sni  = int(_sg["Status"].str.startswith("❌").sum())
                        _ss.append({
                            "Supervisor":      _sup,
                            "Agents":          _sg["Agent Name"].nunique(),
                            "Scheduled Slots": _stot,
                            "✅ On Calls":      _son,
                            "⚠️ No Calls":      _snc,
                            "❌ Not Logged In": _sni,
                            "Compliance %":    round(_son / _stot, 4) if _stot > 0 else 0.0,
                        })
                    sup_summary = pd.DataFrame(_ss)
                    st.markdown("##### 👔 Supervisor Summary")
                    _sup_disp = sup_summary.copy()
                    _sup_disp["Compliance %"] = _sup_disp["Compliance %"].apply(lambda x: f"{x:.1%}")
                    st.dataframe(_sup_disp, use_container_width=True, hide_index=True)

                # ── Overall Summary Pivot ─────────────────────────────────────────
                st.markdown("##### 📊 Overall Summary Pivot")
                st.caption(
                    "Supervisor → Agent Name × Date — each date has two columns: "
                    "**In Call Duration** (within scheduled slot) and "
                    "**Out-of-Slot Duration** (outside scheduled slot / WO day)."
                )

                def _parse_hms(s):
                    try:
                        if str(s).strip() in ("-", "", "nan"):
                            return 0
                        _h, _m, _s = str(s).strip().split(":")
                        return int(_h) * 3600 + int(_m) * 60 + int(_s)
                    except Exception:
                        return 0

                _pdata2 = report_df[
                    ~report_df["Status"].str.startswith("📅", na=False)
                ].copy()
                _pdata2["_in_secs"] = _pdata2["In Call Duration"].apply(_parse_hms)

                _idx_cols_piv = (
                    ["Supervisor", "Agent Name"]
                    if sched_sup_col and "Supervisor" in _pdata2.columns
                    else ["Agent Name"]
                )

                _in_agg = (
                    _pdata2
                    .groupby(_idx_cols_piv + ["Date (IST)"])["_in_secs"]
                    .sum()
                    .reset_index()
                )

                if out_records:
                    _oos_df2 = pd.DataFrame(out_records).copy()
                    _oos_dur = _oos_df2.get(
                        "Out-of-Slot Duration",
                        pd.Series("00:00:00", index=_oos_df2.index),
                    ).fillna(
                        _oos_df2.get("Duration",
                                     pd.Series("00:00:00", index=_oos_df2.index))
                    ).fillna("00:00:00")
                    _oos_df2["_oos_secs"] = _oos_dur.apply(_parse_hms)
                    _oos_cols2 = [c for c in _idx_cols_piv if c in _oos_df2.columns]
                    _oos_agg = (
                        _oos_df2
                        .groupby(_oos_cols2 + ["Date (IST)"])["_oos_secs"]
                        .sum()
                        .reset_index()
                    )
                    for _c2 in _idx_cols_piv:
                        if _c2 not in _oos_agg.columns:
                            _oos_agg[_c2] = ""
                else:
                    _oos_agg = pd.DataFrame(
                        columns=_idx_cols_piv + ["Date (IST)", "_oos_secs"]
                    )

                _clong = (
                    _in_agg
                    .merge(_oos_agg, on=_idx_cols_piv + ["Date (IST)"], how="outer")
                    .fillna(0)
                )
                if "_oos_secs" not in _clong.columns:
                    _clong["_oos_secs"] = 0

                _piv_dates2 = sorted(
                    _clong["Date (IST)"].unique(),
                    key=lambda _d: datetime.strptime(_d, "%d-%b-%Y"),
                )

                _piv_rows = []
                for _key, _grp in _clong.groupby(_idx_cols_piv, sort=True):
                    if not isinstance(_key, tuple):
                        _key = (_key,)
                    _row = dict(zip(_idx_cols_piv, _key))
                    _tot_in = _tot_oos = 0
                    for _date in _piv_dates2:
                        _dg    = _grp[_grp["Date (IST)"] == _date]
                        _in_s  = int(_dg["_in_secs"].sum())
                        _oos_s = int(_dg["_oos_secs"].sum())
                        _row[f"{_date} | In Call"]     = fmt_duration(_in_s)
                        _row[f"{_date} | Out-of-Slot"] = fmt_duration(_oos_s)
                        _tot_in  += _in_s
                        _tot_oos += _oos_s
                    _row["TOTAL | In Call"]     = fmt_duration(_tot_in)
                    _row["TOTAL | Out-of-Slot"] = fmt_duration(_tot_oos)
                    _piv_rows.append(_row)

                _piv_combined = (
                    pd.DataFrame(_piv_rows) if _piv_rows else pd.DataFrame()
                )
                st.dataframe(_piv_combined, use_container_width=True, hide_index=True)

                # ── Detailed table ──────────────────────────────────────────────────
                st.markdown("##### Detailed Compliance Records")
                st.dataframe(report_df, use_container_width=True, hide_index=True)

                # ── Out-of-Slot report ─────────────────────────────────────────
                st.markdown("##### ⏰ Out-of-Slot Login Report")
                if out_records:
                    out_df = pd.DataFrame(out_records)
                    st.caption(
                        f"Logins that occurred **outside** the agent’s scheduled time slot "
                        f"or on Week Off days. "
                        f"({out_df['Agent Name'].nunique()} agent(s), {len(out_df)} record(s))"
                    )
                    oos1, oos2, oos3 = st.columns(3)
                    oos1.metric("Agents with OOS Activity", out_df["Agent Name"].nunique())
                    oos2.metric("Total OOS Records",        len(out_df))
                    oos3.metric("Week Off Logins",
                                int((out_df["Type"] == "Week Off - Logged In").sum()))
                    st.dataframe(out_df, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ No out-of-slot login activity detected.")

                # ── Export ──────────────────────────────────────────────────
                out2 = io.BytesIO()
                with pd.ExcelWriter(out2, engine="openpyxl") as writer:
                    report_df.to_excel(writer, sheet_name="Compliance Report", index=False)
                    _style_ws(writer.sheets["Compliance Report"], pct_col=None)

                    if not sup_summary.empty:
                        _sup_on  = int(sup_summary["✅ On Calls"].sum())
                        _sup_tot = int(sup_summary["Scheduled Slots"].sum())
                        _sup_xl  = pd.concat([
                            sup_summary,
                            pd.DataFrame([{
                                "Supervisor":      "Total",
                                "Agents":          int(sup_summary["Agents"].sum()),
                                "Scheduled Slots": _sup_tot,
                                "✅ On Calls":      _sup_on,
                                "⚠️ No Calls":      int(sup_summary["⚠️ No Calls"].sum()),
                                "❌ Not Logged In": int(sup_summary["❌ Not Logged In"].sum()),
                                "Compliance %":    _sup_on / _sup_tot if _sup_tot > 0 else 0.0,
                            }])
                        ], ignore_index=True)
                        _sup_xl.to_excel(writer, sheet_name="Supervisor Summary", index=False)
                        _style_ws(writer.sheets["Supervisor Summary"], has_total_row=True)
                        _add_bar_chart(writer.sheets["Supervisor Summary"],
                                       "Supervisor", "Supervisor Compliance %",
                                       exclude_last=True)

                    if len(non_wo) > 0:
                        _ag_on  = int(agent_summary["✅ On Calls"].sum())
                        _ag_tot = int(agent_summary["Scheduled Days"].sum())
                        _ag_xl  = pd.concat([
                            agent_summary,
                            pd.DataFrame([{
                                "Agent Name":      "Total",
                                "Scheduled Days":  _ag_tot,
                                "✅ On Calls":      _ag_on,
                                "⚠️ No Calls":      int(agent_summary["⚠️ No Calls"].sum()),
                                "❌ Not Logged In": int(agent_summary["❌ Not Logged In"].sum()),
                                "Compliance %":    _ag_on / _ag_tot if _ag_tot > 0 else 0.0,
                            }])
                        ], ignore_index=True)
                        _ag_xl.to_excel(writer, sheet_name="Agent Summary", index=False)
                        _style_ws(writer.sheets["Agent Summary"], has_total_row=True)
                        _add_bar_chart(writer.sheets["Agent Summary"],
                                       "Agent Name", "Agent Compliance %",
                                       exclude_last=True)

                    if not _piv_combined.empty:
                        _piv_combined.to_excel(writer, sheet_name="Summary Pivot", index=False)
                        _style_ws(writer.sheets["Summary Pivot"], pct_col=None)

                    if out_records:
                        pd.DataFrame(out_records).to_excel(
                            writer, sheet_name="Out-of-Slot Logins", index=False
                        )
                        _style_ws(writer.sheets["Out-of-Slot Logins"], pct_col=None)

                st.download_button(
                    "📥 Download Compliance Report (Excel)",
                    data=out2.getvalue(),
                    file_name=f"ScheduleCompliance_{date_from}_{date_to}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

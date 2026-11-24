"""
Timesheet App (Streamlit)
-------------------------
Log the dates you worked, hours and a description, manage employees,
export an Excel file in the same layout as the monthly timesheet, and see
a quick analysis of the hours.

Run:  streamlit run timesheet_app.py
Data is saved locally in timesheet_data.json (next to this file).
"""

import calendar
import json
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# ----------------------------------------------------------------------------
# Config
# ----------------------------------------------------------------------------
DATA_FILE = Path(__file__).parent / "timesheet_data.json"
DEFAULT_PROJECT = "Kontoor Brands Project Tasks"
DEFAULT_EMPLOYEES = ["Anjana Fredricks"]
COLUMNS = ["Date", "Project", "Hours", "User", "Comment"]  # same as the Excel file

st.set_page_config(page_title="Timesheet", page_icon="🕒", layout="wide")


# ----------------------------------------------------------------------------
# Storage
# ----------------------------------------------------------------------------
def load_data() -> dict:
    if DATA_FILE.exists():
        try:
            data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            data.setdefault("employees", [])
            data.setdefault("entries", [])
            projects = data.setdefault("projects", [])
            # keep any project already used in entries available in the list
            for e in data["entries"]:
                if e.get("Project") and e["Project"] not in projects:
                    projects.append(e["Project"])
            if not projects:
                projects.append(DEFAULT_PROJECT)
            return data
        except Exception:
            pass
    return {"employees": list(DEFAULT_EMPLOYEES), "projects": [DEFAULT_PROJECT], "entries": []}


def commit() -> None:
    """Persist data to disk and invalidate the editor widgets."""
    DATA_FILE.write_text(
        json.dumps(st.session_state.data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    st.session_state.ver += 1


if "data" not in st.session_state:
    st.session_state.data = load_data()
    st.session_state.ver = 0


# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def all_entries_df() -> pd.DataFrame:
    df = pd.DataFrame(st.session_state.data["entries"], columns=COLUMNS)
    if df.empty:
        return df
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0.0)
    df["Comment"] = df["Comment"].fillna("")
    return df


def month_df(df: pd.DataFrame, year: int, month: int) -> pd.DataFrame:
    if df.empty:
        return df
    return df[df["Date"].str.startswith(f"{year}-{month:02d}")]


def days_in_month(year: int, month: int) -> list:
    n = calendar.monthrange(year, month)[1]
    return [date(year, month, d) for d in range(1, n + 1)]


def upsert_entries(employee: str, days: list, project: str, hours: float, comment: str):
    """Add entries; an existing entry for the same employee+date is replaced."""
    keys = {d.isoformat() for d in days}
    kept = [
        e
        for e in st.session_state.data["entries"]
        if not (e["User"] == employee and e["Date"] in keys)
    ]
    for d in sorted(days):
        kept.append(
            {
                "Date": d.isoformat(),
                "Project": project,
                "Hours": hours,
                "User": employee,
                "Comment": comment,
            }
        )
    st.session_state.data["entries"] = kept


def fmt_hours(x) -> float | int:
    x = float(x)
    return int(x) if x.is_integer() else x


def build_workbook(frames: dict, add_total: bool) -> bytes:
    """frames: {sheet_name: DataFrame[COLUMNS]} -> xlsx bytes (same layout as the original)."""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet, df in frames.items():
        ws = wb.create_sheet(sheet)
        ws.append(COLUMNS)
        for c in ws[1]:
            c.font = Font(name="Calibri", size=11, bold=True)
            c.alignment = Alignment(horizontal="center")

        df = df.sort_values(["Date", "User"])
        for r in df.itertuples(index=False):
            ws.append([r.Date, r.Project, fmt_hours(r.Hours), r.User, r.Comment or None])
        last = ws.max_row

        for row in ws.iter_rows(min_row=2, max_row=last):
            for c in row:
                c.font = Font(name="Aptos Narrow", size=11)
            row[0].number_format = "@"  # keep date as text, like the original

        if add_total and last >= 2:
            tr = last + 1
            ws.cell(tr, 2, "Total").font = Font(name="Aptos Narrow", size=11, bold=True)
            tc = ws.cell(tr, 3, f"=SUM(C2:C{last})")
            tc.font = Font(name="Aptos Narrow", size=11, bold=True)

        for col, w in zip("ABCDE", (12, 30, 10, 22, 36)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def on_add(employee: str, pick_key: str, project_key: str, project_default: str) -> None:
    """Button callback: runs before widgets are redrawn, so inputs can be reset here."""
    days = st.session_state.get(pick_key, [])
    if not days:
        return
    upsert_entries(
        employee,
        days,
        st.session_state[project_key].strip() or project_default,
        float(st.session_state["log_hours"]),
        st.session_state["log_comment"].strip(),
    )
    commit()
    st.session_state[pick_key] = []
    st.session_state["log_comment"] = ""
    st.session_state["flash"] = f"Added {len(days)} day(s) for {employee}."


# ----------------------------------------------------------------------------
# Sidebar: employees + settings
# ----------------------------------------------------------------------------
with st.sidebar:
    st.header("👥 Employees")
    new_emp = st.text_input("Add employee", placeholder="Full name", key=f"new_emp_{st.session_state.ver}")
    if st.button("Add employee"):
        name = new_emp.strip()
        if not name:
            st.warning("Enter a name first.")
        elif name in st.session_state.data["employees"]:
            st.warning("That employee already exists.")
        else:
            st.session_state.data["employees"].append(name)
            commit()
            st.rerun()

    with st.expander("Remove employee"):
        if st.session_state.data["employees"]:
            rem = st.selectbox("Employee", st.session_state.data["employees"], key="rem_emp")
            ok = st.checkbox("Also delete all of their entries")
            if st.button("Remove", disabled=not ok):
                st.session_state.data["employees"].remove(rem)
                st.session_state.data["entries"] = [
                    e for e in st.session_state.data["entries"] if e["User"] != rem
                ]
                commit()
                st.rerun()

    st.divider()
    st.header("📁 Projects")
    new_proj = st.text_input("Add project", placeholder="Project name", key=f"new_proj_{st.session_state.ver}")
    if st.button("Add project"):
        name = new_proj.strip()
        if not name:
            st.warning("Enter a project name first.")
        elif name.lower() in (p.lower() for p in st.session_state.data["projects"]):
            st.warning("That project already exists.")
        else:
            st.session_state.data["projects"].append(name)
            commit()
            st.rerun()

    with st.expander("Remove project"):
        if len(st.session_state.data["projects"]) <= 1:
            st.caption("At least one project must remain.")
        else:
            rem_p = st.selectbox("Project", st.session_state.data["projects"], key="rem_proj")
            used = sum(1 for e in st.session_state.data["entries"] if e["Project"] == rem_p)
            st.caption(f"{used} existing entr{'y' if used == 1 else 'ies'} use this project.")
            del_entries = st.checkbox("Also delete those entries", key="rem_proj_entries", disabled=used == 0)
            if st.button("Remove project"):
                st.session_state.data["projects"].remove(rem_p)
                if del_entries:
                    st.session_state.data["entries"] = [
                        e for e in st.session_state.data["entries"] if e["Project"] != rem_p
                    ]
                commit()
                st.rerun()

    st.divider()
    st.header("⚙️ Settings")
    projects = st.session_state.data["projects"]
    project_default = st.selectbox("Default project", projects)
    std_hours = st.number_input("Standard hours per working day", 1.0, 24.0, 8.0, 0.5)

employees = st.session_state.data["employees"]

# ----------------------------------------------------------------------------
# Header: shared selectors
# ----------------------------------------------------------------------------
st.title("🕒 Timesheet")

if not employees:
    st.info("Add an employee in the sidebar to get started.")
    st.stop()

today = date.today()
c1, c2, c3 = st.columns([2, 1, 1])
employee = c1.selectbox("Employee", employees)
month_names = list(calendar.month_name)[1:]
month = month_names.index(c2.selectbox("Month", month_names, index=today.month - 1)) + 1
year = int(c3.number_input("Year", 2020, 2100, today.year, 1))

all_df = all_entries_df()
mdf = month_df(all_df, year, month)
emp_mdf = mdf[mdf["User"] == employee] if not mdf.empty else mdf
logged_days = set(emp_mdf["Date"]) if not emp_mdf.empty else set()

m1, m2, m3 = st.columns(3)
m1.metric(f"{employee} – hours in {month_names[month-1]} {year}", f"{emp_mdf['Hours'].sum():g}" if not emp_mdf.empty else "0")
m2.metric("Days logged", len(logged_days))
m3.metric("Entries across all employees (this month)", 0 if mdf.empty else len(mdf))

tab_log, tab_review, tab_export, tab_analysis, tab_import = st.tabs(
    ["➕ Log hours", "📋 Review & edit", "📤 Export Excel", "📊 Analysis", "📥 Import"]
)

# ----------------------------------------------------------------------------
# Tab 1: log hours
# ----------------------------------------------------------------------------
with tab_log:
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    st.subheader("Pick the dates you worked")
    month_days = days_in_month(year, month)
    pick_key = f"days_{year}_{month}"

    def fmt_day(d: date) -> str:
        mark = "  ✓ logged" if d.isoformat() in logged_days else ""
        return d.strftime("%a %d %b") + mark

    b1, b2, _ = st.columns([1, 1, 4])
    b1.button(
        "Select all weekdays",
        on_click=lambda: st.session_state.__setitem__(pick_key, [d for d in month_days if d.weekday() < 5]),
    )
    b2.button("Clear", on_click=lambda: st.session_state.__setitem__(pick_key, []))

    picked = st.multiselect(
        "Dates", month_days, key=pick_key, format_func=fmt_day,
        placeholder="Click to choose dates (Sat/Sun are included too)",
    )

    h1, h2 = st.columns([1, 3])
    hours = h1.number_input("Hours per day", 0.5, 24.0, 8.0, 0.5, key="log_hours")
    project_key = "log_project_" + "|".join(projects)  # new key when the list changes
    h2.selectbox("Project", projects, index=projects.index(project_default), key=project_key)
    comment = st.text_area(
        "Description – what did you do?",
        placeholder="e.g. Cutover weekend support",
        key="log_comment",
    )
    st.caption(
        "Tip: different hours or description on some days (e.g. a weekend)? "
        "Add those dates separately with their own hours and description. "
        "Adding a date that is already logged will replace it."
    )

    st.button(
        "Add to timesheet",
        type="primary",
        disabled=not picked,
        on_click=on_add,
        args=(employee, pick_key, project_key, project_default),
    )

# ----------------------------------------------------------------------------
# Tab 2: review & edit
# ----------------------------------------------------------------------------
with tab_review:
    st.subheader(f"{employee} – {month_names[month-1]} {year}")
    st.caption("Edit cells, add rows at the bottom, or select a row and press Delete. Then click Save.")

    cur = emp_mdf.sort_values("Date") if not emp_mdf.empty else emp_mdf
    editor_df = pd.DataFrame(
        {
            "Date": pd.to_datetime(cur["Date"]) if not cur.empty else pd.Series([], dtype="datetime64[ns]"),
            "Project": cur["Project"].astype(str) if not cur.empty else pd.Series([], dtype="object"),
            "Hours": cur["Hours"].astype(float) if not cur.empty else pd.Series([], dtype="float64"),
            "Comment": cur["Comment"].astype(str) if not cur.empty else pd.Series([], dtype="object"),
        }
    ).reset_index(drop=True)
    edited = st.data_editor(
        editor_df,
        num_rows="dynamic",
        hide_index=True,
        key=f"editor_{employee}_{year}_{month}_{st.session_state.ver}",
        column_config={
            "Date": st.column_config.DateColumn("Date", format="YYYY-MM-DD", required=True),
            "Project": st.column_config.SelectboxColumn(
                "Project",
                options=list(dict.fromkeys(projects + editor_df["Project"].dropna().tolist())),
                default=project_default,
                required=True,
            ),
            "Hours": st.column_config.NumberColumn("Hours", min_value=0.0, max_value=24.0, step=0.5, default=float(std_hours)),
            "Comment": st.column_config.TextColumn("Comment", width="large"),
        },
    )

    if st.button("💾 Save changes", type="primary"):
        new_rows, errors, seen = [], [], set()
        for _, r in edited.iterrows():
            if pd.isna(r["Date"]):
                continue
            d = pd.to_datetime(r["Date"]).date()
            if (d.year, d.month) != (year, month):
                errors.append(f"{d} is not in {month_names[month-1]} {year}.")
            elif d in seen:
                errors.append(f"{d} appears more than once.")
            elif pd.isna(r["Hours"]) or r["Hours"] <= 0:
                errors.append(f"{d} needs hours greater than 0.")
            else:
                seen.add(d)
                new_rows.append(
                    {
                        "Date": d.isoformat(),
                        "Project": (r["Project"] or project_default) if not pd.isna(r["Project"]) else project_default,
                        "Hours": float(r["Hours"]),
                        "User": employee,
                        "Comment": "" if pd.isna(r["Comment"]) else str(r["Comment"]),
                    }
                )
        if errors:
            for e in errors:
                st.error(e)
        else:
            prefix = f"{year}-{month:02d}"
            st.session_state.data["entries"] = [
                e
                for e in st.session_state.data["entries"]
                if not (e["User"] == employee and e["Date"].startswith(prefix))
            ] + new_rows
            commit()
            st.success("Saved.")
            st.rerun()

# ----------------------------------------------------------------------------
# Tab 3: export
# ----------------------------------------------------------------------------
with tab_export:
    st.subheader("Download timesheet as Excel")
    e1, e2 = st.columns(2)
    who = e1.radio("Employees", [f"Only {employee}", "All employees"], horizontal=True)
    span = e2.radio("Period", [f"{month_names[month-1]} {year}", "All months"], horizontal=True)
    add_total = st.checkbox("Add a Total row at the bottom", value=False)

    src = all_df
    if who == "All employees":
        label = "All"
    else:
        src = src[src["User"] == employee] if not src.empty else src
        label = employee.split()[0]

    if span == "All months":
        frames = {m: g for m, g in src.groupby(src["Date"].str[:7])} if not src.empty else {}
        fname_month = "All_Months"
    else:
        sel = month_df(src, year, month)
        frames = {f"{year}-{month:02d}": sel} if not sel.empty else {}
        fname_month = month_names[month - 1]

    if not frames:
        st.warning("No entries for this selection yet.")
    else:
        preview = pd.concat(frames.values()).sort_values(["Date", "User"])
        st.dataframe(preview, hide_index=True)
        st.caption(f"{len(preview)} rows · {preview['Hours'].sum():g} hours")
        st.download_button(
            "⬇️ Download Excel",
            data=build_workbook(frames, add_total),
            file_name=f"Kontoor_Timesheet_-_{label}_-_{fname_month}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

# ----------------------------------------------------------------------------
# Tab 4: analysis
# ----------------------------------------------------------------------------
with tab_analysis:
    st.subheader(f"Analysis – {month_names[month-1]} {year}")
    if mdf.empty:
        st.info("No hours logged for this month yet.")
    else:
        weekdays = [d for d in days_in_month(year, month) if d.weekday() < 5]
        expected = len(weekdays) * std_hours

        rows = []
        for emp in employees:
            e = mdf[mdf["User"] == emp]
            if e.empty:
                rows.append({"Employee": emp, "Days worked": 0, "Total hours": 0.0, "Avg hours/day": 0.0,
                             "Weekend hours": 0.0, "Expected hours": expected, "Difference": -expected})
                continue
            wk = pd.to_datetime(e["Date"]).dt.weekday >= 5
            tot = float(e["Hours"].sum())
            rows.append({
                "Employee": emp,
                "Days worked": int(e["Date"].nunique()),
                "Total hours": tot,
                "Avg hours/day": round(tot / e["Date"].nunique(), 2),
                "Weekend hours": float(e.loc[wk, "Hours"].sum()),
                "Expected hours": expected,
                "Difference": tot - expected,
            })
        summary = pd.DataFrame(rows)

        st.markdown(f"**Team summary** – {len(weekdays)} weekdays × {std_hours:g} h = {expected:g} expected hours per person")
        st.dataframe(summary, hide_index=True)
        st.bar_chart(summary.set_index("Employee")["Total hours"])

        st.divider()
        st.markdown(f"**{employee}**")
        if emp_mdf.empty:
            st.info("No entries for this employee in this month.")
        else:
            s = summary[summary["Employee"] == employee].iloc[0]
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Total hours", f"{s['Total hours']:g}")
            a2.metric("Days worked", int(s["Days worked"]))
            a3.metric("Avg hours / day", f"{s['Avg hours/day']:g}")
            a4.metric("vs expected", f"{s['Difference']:+g} h")

            daily = (
                emp_mdf.groupby("Date")["Hours"].sum()
                .reindex([d.isoformat() for d in days_in_month(year, month)], fill_value=0)
            )
            st.caption("Hours per day")
            st.bar_chart(daily)

            worked = {d for d in emp_mdf["Date"]}
            missing = [d for d in weekdays if d.isoformat() not in worked]
            weekend_worked = [d for d in sorted(worked) if date.fromisoformat(d).weekday() >= 5]
            if missing:
                st.warning(
                    "Weekdays with no hours (leave / holiday?): "
                    + ", ".join(d.strftime("%a %d %b") for d in missing)
                )
            else:
                st.success("Every weekday this month has hours logged.")
            if weekend_worked:
                st.info(
                    "Weekend days worked: "
                    + ", ".join(date.fromisoformat(d).strftime("%a %d %b") for d in weekend_worked)
                )

            if emp_mdf["Project"].nunique() > 1:
                st.caption("Hours by project")
                st.bar_chart(emp_mdf.groupby("Project")["Hours"].sum())

    # Monthly trend for selected employee (all data)
    emp_all = all_df[all_df["User"] == employee] if not all_df.empty else all_df
    if not emp_all.empty and emp_all["Date"].str[:7].nunique() > 1:
        st.divider()
        st.markdown(f"**{employee} – monthly trend**")
        st.bar_chart(emp_all.groupby(emp_all["Date"].str[:7])["Hours"].sum())

# ----------------------------------------------------------------------------
# Tab 5: import
# ----------------------------------------------------------------------------
with tab_import:
    st.subheader("Import existing timesheet(s)")
    st.caption("Upload a file in the same format (Date, Project, Hours, User, Comment). "
               "Existing entries for the same user and date are replaced; new users are added as employees.")
    up = st.file_uploader("Excel file", type=["xlsx"])
    if up is not None:
        try:
            sheets = pd.read_excel(up, sheet_name=None, dtype={"Date": str})
            frames_in = [s for s in sheets.values() if set(COLUMNS[:4]).issubset(s.columns)]
            if not frames_in:
                st.error("No sheet with the columns Date, Project, Hours, User found.")
            else:
                imp = pd.concat(frames_in, ignore_index=True).dropna(subset=["Date", "Hours", "User"])
                imp["Date"] = pd.to_datetime(imp["Date"]).dt.strftime("%Y-%m-%d")
                if "Comment" not in imp:
                    imp["Comment"] = ""
                imp["Comment"] = imp["Comment"].fillna("").astype(str)
                st.dataframe(imp[COLUMNS], hide_index=True)
                if st.button(f"Import {len(imp)} rows", type="primary"):
                    keys = set(zip(imp["User"], imp["Date"]))
                    st.session_state.data["entries"] = [
                        e for e in st.session_state.data["entries"] if (e["User"], e["Date"]) not in keys
                    ] + imp[COLUMNS].to_dict("records")
                    for pr in imp["Project"].dropna().unique():
                        if pr not in st.session_state.data["projects"]:
                            st.session_state.data["projects"].append(pr)
                    for u in imp["User"].unique():
                        if u not in st.session_state.data["employees"]:
                            st.session_state.data["employees"].append(u)
                    commit()
                    st.success("Imported.")
                    st.rerun()
        except Exception as ex:
            st.error(f"Could not read the file: {ex}")

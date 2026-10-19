"""Data + report logic for the MetLife Global Operations check-wise TAT report."""
import io
import os
import re
from pathlib import Path
from datetime import date, datetime, time

import numpy as np
import pandas as pd
import pymysql
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

DEFAULT_CLIENT_CODE = "5800"

# Check names are trimmed from the first "(*" / "( *" onwards.
CHECK_NAME_CUTOFF = re.compile(r"\s*\(\s*\*.*$", re.DOTALL)

CHECK_QUERY = """
SELECT ecc.case_id,
       ecc.case_check_id,
       process_name,
       emc.company_name,
       office_name location,
       case_ars_no, CONCAT(first_name,' ',Middle_name,' ',last_name) AS Candidate_name,
       received_date case_received_date,
       ecm.created_date AS case_created_date,
       case_expected_closure_date,
       checkpoint_live.fn_case_status(case_status) AS 'Case_status',
       ecc.check_name,
       check_created_on,
       go_ahead_date,
       check_expected_closure,
       Insuff_date,
       ecc.insuff_remarks,
       @insufffulfilldate:=IF(Insuff_fulfill_date<check_closure_date,Insuff_fulfill_date,'') 'Insuff_fulfill_date',
       checkpoint_live.fn_check_status(check_status) AS 'check_status',
       check_severity,
       DATE(check_closure_date) AS 'Check Closure Date',
       CONCAT(eud.user_first_name,' ',eud.User_last_name) Created_by,
       eud1.user_id,
       CONCAT(eud1.user_first_name,' ',eud1.User_last_name) Verifier,
       (CASE ecc.VER_OVERSEAS
        WHEN '0' THEN 'No'
        WHEN '1' THEN 'Yes' END) AS Ver_Overseas,
       fn_ver_summary(ecc.VER_SUMMARY) AS Ver_Summary,
       (CASE ecc.VER_TYPE
            WHEN '1' THEN 'Verbal'
            WHEN '2' THEN 'Written'
            WHEN '3' THEN 'VRWP'
       END) AS Ver_Type,
       fn_ver_procedure(ecc.VERIFICATION_PROCEDURE) AS VERIFICATION_PROCEDURE,
       IF(ecc.family_id=4,emei.institute_name,IF(ecc.family_id=5,emc1.company_name,emcy.city_name)) AS verification_source,
       DATE(dqc_released_date) AS dqc_released_date,
       @checktat:=IF(ectc1.actual_tat IS NULL,epc.tat,ectc1.actual_tat) AS 'Check_TAT',
       check_disposition_id,
       disposition_name,
       ecc.closure_comments,
       DATE(ecc.reopen_date) AS reopen_date,
       @calreopendate:=(SELECT action_taken_on FROM ec_case_history ech WHERE ecc.case_check_id=ech.check_id AND
          ech.action_taken LIKE '%%reopen%%' AND (action_comments LIKE '%%Customer Approval%%' OR action_comments LIKE '%%Customer Error%%'
          OR action_comments LIKE '%%Customer Complaint%%' OR action_comments LIKE '%%Customer Request%%')
          ORDER BY action_id DESC LIMIT 1) AS 'Calculated Reopen_date',
       @reopendate:=IF(@calreopendate < check_closure_date,@calreopendate,'') AS 'Reopen Date (With Condition)',
       @goaheaddate:=IF(ecc.go_ahead_date < check_closure_date,ecc.go_ahead_date,'') AS 'Goahead Date (with Condition)',
       @max_date:=(GREATEST(received_date,@reopendate,@goaheaddate,@insufffulfilldate)) AS 'Max_date',
       ec1.check_name AS 'Unique Check Name',
       @processtat:=IF(ectc.actual_tat IS NULL,ecp.TAT,ectc.actual_tat) AS 'Process Wise TAT',
       ec1.CHECK_OPS_NAME,
       @checkduedate:=workday(@max_date,@checktat) AS 'Checks Due Date',
       @processduedate:=workday(@max_date,@processtat) AS 'Process Wise Due Date',
       IF(ecc.check_status IN (2,3,10),'IT',IF(@checkduedate<ifnull(DATE(check_closure_date),CURRENT_DATE()),'OT','IT')) AS 'TAT_Check TAT',
       IF(ecc.check_status IN (2,3,10),'IT',IF(@processduedate<IFNULL(DATE(check_closure_date),CURRENT_DATE()),'OT','IT')) AS 'TAT _Process TAT',
       ecc.check_status AS check_status_code

FROM ec_case_checks ecc
LEFT JOIN ec_case_master ecm ON ecc.case_id=ecm.case_id
LEFT JOIN ec_master_company emc ON ecm.client_id=emc.company_id
LEFT JOIN ec_user_details eud ON ecc.documented_by=eud.user_id
LEFT JOIN ec_user_details eud1 ON ecc.check_verifier=eud1.user_id
LEFT JOIN ec_case_tat_config ectc ON (ecm.client_id=ectc.client_id AND ecm.process_id=ectc.process_id)
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
INNER JOIN ec_process_checks epc ON (ecm.process_id=epc.process_id AND epc.check_id=ecc.check_id)
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_city emcy ON emcy.city_id=eccvs.org_id
LEFT JOIN ec_master_disposition ecd ON ecc.check_disposition_id=ecd.disposition_id
LEFT JOIN ec_client ec ON ecm.client_id=ec.client_id
LEFT JOIN ec_checks ec1 ON ecc.check_id=ec1.check_id
LEFT JOIN ec_check_tat_config ectc1 ON (ectc.id=ectc1.case_tat_id AND ecc.check_id=ectc1.check_id AND ectc1.extra=CONCAT('tr',emcy.tier))

WHERE ecc.check_status<>9 AND case_status <>8
AND client_external_id IN (%s)
AND ecm.received_date>=%s
AND ecm.received_date<=%s
"""

TAT_CONFIG_QUERY = """
SELECT client_external_id 'Client Code',
       ectc.id AS 'case_tat_id',
       ectc.client_id,
       company_name AS client_name,
       ectc.process_id,
       process_name,
       tat_tp,
       tat_catg,
       tat_catg_val,
       work_hr_catg,
       ectc.actual_tat AS 'process actual tat',
       ectc.operational_tat AS 'process operational tat',
       ectc.reopen_criteria AS 'process reopen criteria',
       ectc.is_fast_track AS 'process fast track',
       IF(ectc.is_tier_tat_addr=1,'Yes','No') AS 'process tier addr tat applicable',
       ectc1.id AS 'check tat id',
       ectc1.check_id,
       check_name,
       ectc1.actual_tat AS 'check actual tat',
       ectc1.operational_tat AS 'check operational tat',
       ectc1.reopen_tat AS 'check reopen tat',
       ectc1.is_fast_track  AS 'check fast track',
       insuff_criteria AS 'client insuff criteria', ecic.fixed 'client insuff criteria value',
       ecict.tat AS 'check insuff tat',
       ectc1.extra AS 'check tr',
       (CASE ecp.process_status
        WHEN 0 THEN 'Inactive'
        WHEN 1 THEN 'Active' END) AS 'Process Status',
       IF(ectc1.check_id=epc.check_id,(CASE is_active_check
        WHEN 0 THEN 'Inactive check'
        WHEN 1 THEN 'Active' END),'Check not existing on Process') AS 'check_status'

FROM ec_case_tat_config ectc
LEFT JOIN ec_check_tat_config ectc1 ON ectc.id=ectc1.case_tat_id
LEFT JOIN ec_client_process ecp ON ectc.process_id=ecp.process_id
LEFT JOIN ec_master_company emc ON ectc.client_id=emc.company_id
LEFT JOIN ec_checks ec ON ectc1.check_id=ec.check_id
LEFT JOIN ec_client_insuff_criteria ecic ON ectc.client_id=ecic.client_id
LEFT JOIN ec_client_insuff_check_tat ecict ON (ecic.id=ecict.client_insuff_criteria_id AND ectc.client_id=ecict.client_id AND ectc1.check_id=ecict.check_id)
INNER JOIN ec_client ec1 ON ectc.client_id=ec1.client_id
LEFT JOIN ec_process_checks epc ON (ectc.process_id=epc.process_id AND ectc1.check_id=epc.check_id)
WHERE client_external_id IN (%s)
"""

DATE_COLS = [
    "case_received_date", "case_created_date", "case_expected_closure_date", "check_created_on",
    "go_ahead_date", "check_expected_closure", "Insuff_date", "Insuff_fulfill_date",
    "Check Closure Date", "dqc_released_date", "reopen_date", "Calculated Reopen_date",
    "Reopen Date (With Condition)", "Goahead Date (with Condition)", "Max_date",
    "Checks Due Date", "Process Wise Due Date",
]


def _connect():
    return pymysql.connect(
        host=os.getenv("DB_HOST"),
        port=int(os.getenv("DB_PORT", "3306")),
        user=os.getenv("DB_USER"),
        password=os.getenv("DB_PASSWORD"),
        database=os.getenv("DB_NAME"),
        charset="utf8mb4",
        connect_timeout=20,
    )


def _query(sql, params) -> pd.DataFrame:
    conn = _connect()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            cols = [c[0] for c in cur.description]
            rows = cur.fetchall()
    finally:
        conn.close()
    return pd.DataFrame(rows, columns=cols)


def clean_check_name(name) -> str:
    if name is None or (isinstance(name, float) and np.isnan(name)):
        return ""
    return CHECK_NAME_CUTOFF.sub("", str(name)).strip()


def _key(name) -> str:
    return re.sub(r"\s+", " ", str(name)).strip().lower()


def fetch_tat_config(client_code: str) -> pd.DataFrame:
    return _query(TAT_CONFIG_QUERY, (client_code,))


def fetch_checks(client_code: str, start: date, end: date) -> pd.DataFrame:
    # Mirrors "<= today date 23:59:50" from the original query.
    start_ts = datetime.combine(start, time.min)
    end_ts = datetime.combine(end, time(23, 59, 50))
    return _query(CHECK_QUERY, (client_code, start_ts, end_ts))


def fetch_holidays() -> np.ndarray:
    """Master holidays plus global (CLIENT_ID is NULL) client holidays, as datetime64[D]."""
    df = _query(
        "SELECT HOLIDAY_DATE FROM ec_master_holidays "
        "UNION SELECT HOLIDAY_DATE FROM ec_client_holidays WHERE CLIENT_ID IS NULL",
        None,
    )
    return pd.to_datetime(df["HOLIDAY_DATE"]).values.astype("datetime64[D]")


def _check_tat_lookup(config: pd.DataFrame):
    """check_name -> 'check actual tat'. Active, process-active rows win; ties take the most common value."""
    cfg = config.dropna(subset=["check_name", "check actual tat"]).copy()
    cfg["_key"] = cfg["check_name"].map(_key)
    cfg["_rank"] = (cfg["check_status"] == "Active").astype(int) * 2 + (cfg["Process Status"] == "Active").astype(int)
    lookup = {}
    for k, grp in cfg.groupby("_key"):
        best = grp[grp["_rank"] == grp["_rank"].max()]["check actual tat"]
        lookup[k] = float(best.mode().min())
    return lookup


def _workday(start: pd.Series, days: pd.Series, holidays=None) -> pd.Series:
    """Excel-style WORKDAY (Mon-Fri, skipping holidays); NaT/NaN stay NaT."""
    ok = start.notna() & days.notna()
    out = pd.Series(pd.NaT, index=start.index, dtype="datetime64[ns]")
    if ok.any():
        s = start[ok].values.astype("datetime64[D]")
        out[ok] = np.busday_offset(s, days[ok].astype(int).values, roll="backward",
                                 holidays=holidays if holidays is not None else []).astype("datetime64[ns]")
    return out


def build_report(raw: pd.DataFrame, config: pd.DataFrame, holidays=None) -> pd.DataFrame:
    df = raw.copy()
    for c in DATE_COLS:
        df[c] = pd.to_datetime(df[c].replace("", pd.NA), errors="coerce")

    df["check_name2"] = df["check_name"].map(clean_check_name)

    lookup = _check_tat_lookup(config)
    mapped = df["check_name2"].map(_key).map(lookup)
    df["Check_TAT_source"] = np.where(mapped.notna(), "TAT configuration", "Original query (no config match)")
    df["Check_TAT_query"] = df["Check_TAT"]
    df["Check_TAT"] = mapped.fillna(pd.to_numeric(df["Check_TAT"], errors="coerce"))

    df["Month"] = df["case_received_date"].dt.strftime("%b-%y")
    # Max TAT = latest of the query's Max_date and the check's creation date (date part only).
    df["Max TAT"] = pd.concat([df["Max_date"], df["check_created_on"].dt.normalize()], axis=1).max(axis=1)
    df["Checks Due Date1"] = _workday(df["Max TAT"], df["Check_TAT"], holidays)

    closed = df["Check Closure Date"].fillna(pd.Timestamp(date.today()))
    df["TAT"] = np.where(df["Checks Due Date1"] < closed, "OT", "IT")
    return df.drop(columns=["check_status_code", "Check_TAT_query"])


# --- summaries ---------------------------------------------------------------------------

def _month_order(df):
    return df.drop_duplicates("Month").sort_values("case_received_date")["Month"].tolist()


def summary_by_month(df: pd.DataFrame) -> pd.DataFrame:
    p = df.pivot_table(index="Month", columns="TAT", values="case_check_id", aggfunc="count", fill_value=0)
    for c in ("IT", "OT"):
        if c not in p:
            p[c] = 0
    p = p[["IT", "OT"]].reindex(_month_order(df))
    p["Grand Total"] = p["IT"] + p["OT"]
    p.loc["Grand Total"] = p.sum()
    p["OT %"] = p["OT"] / p["Grand Total"]
    return p.reset_index()


def summary_by_check(df: pd.DataFrame) -> pd.DataFrame:
    months = _month_order(df)
    p = df.pivot_table(index="check_name2", columns=["TAT", "Month"], values="case_check_id",
                       aggfunc="count", fill_value=0)
    out = pd.DataFrame(index=p.index)
    for tat in ("IT", "OT"):
        for m in months:
            out[f"{tat} {m}"] = p[(tat, m)] if (tat, m) in p.columns else 0
        out[f"{tat} Total"] = out[[f"{tat} {m}" for m in months]].sum(axis=1)
    out["Grand Total"] = out["IT Total"] + out["OT Total"]
    out["IT %"] = out["IT Total"] / out["Grand Total"]
    out["OT %"] = out["OT Total"] / out["Grand Total"]
    out = out.sort_index()
    total = out.drop(columns=["IT %", "OT %"]).sum()
    total["IT %"] = total["IT Total"] / total["Grand Total"]
    total["OT %"] = total["OT Total"] / total["Grand Total"]
    out.loc["Grand Total"] = total
    out.index.name = "Check Name"
    return out.reset_index()


# --- Excel ---------------------------------------------------------------------------------

def _write_data_workbook(df: pd.DataFrame, path: str, with_static_summary: bool) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "#FFFFFF", "border": 1, "align": "center"})
        bold = wb.add_format({"bold": True})
        pct = wb.add_format({"num_format": "0.0%"})
        if with_static_summary:
            by_month, by_check = summary_by_month(df), summary_by_check(df)
            by_month.to_excel(xw, sheet_name="Summary", index=False, startrow=1)
            r2 = len(by_month) + 5
            by_check.to_excel(xw, sheet_name="Summary", index=False, startrow=r2)
            ws = xw.sheets["Summary"]
            ws.write(0, 0, "TAT by Month (count of checks)", bold)
            ws.write(r2 - 1, 0, "TAT by Check Name", bold)
            for i, c in enumerate(by_month.columns):
                ws.write(1, i, c, hdr)
            for i, c in enumerate(by_check.columns):
                ws.write(r2, i, c, hdr)
            ws.set_column(0, 0, 38)
            ws.set_column(1, len(by_check.columns), 13)
            for i, c in enumerate(by_check.columns):
                if c.endswith("%"):
                    ws.set_column(i, i, 10, pct)
            ws.set_column(4, 4, 13, pct)
        df.to_excel(xw, sheet_name="Data", index=False)
        d = xw.sheets["Data"]
        for i, c in enumerate(df.columns):
            d.write(0, i, c, hdr)
            d.set_column(i, i, min(max(len(str(c)) + 2, 12), 40))
        d.freeze_panes(1, 0)


def _add_pivots(path: str, months: list) -> None:
    """Open the workbook in Excel and build the real PivotTables on a new 'Summary' sheet."""
    import pythoncom
    import win32com.client as win32

    xlDatabase, xlRowField, xlColumnField, xlCount, xlUp = 1, 1, 2, -4112, -4162
    pythoncom.CoInitialize()
    excel = wb = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(path)
        data = wb.Sheets("Data")
        last_row = data.Cells(data.Rows.Count, 1).End(xlUp).Row
        last_col = data.Cells(1, data.Columns.Count).End(-4159).Column
        src = data.Range(data.Cells(1, 1), data.Cells(last_row, last_col))
        tbl = data.ListObjects.Add(1, src, None, 1)
        tbl.Name = "DataTab"
        tbl.TableStyle = "TableStyleMedium2"

        summ = wb.Sheets.Add(Before=wb.Sheets(1))
        summ.Name = "Summary"
        cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData="DataTab")

        def make(anchor_row, name, col_fields, row_field, row_caption):
            pt = cache.CreatePivotTable(TableDestination=summ.Cells(anchor_row, 1), TableName=name)
            pt.TableStyle2 = "PivotStyleMedium9"
            pt.RowAxisLayout(1)  # tabular: shows the "Month" / "Check Name" captions
            rf = pt.PivotFields(row_field)
            rf.Orientation = xlRowField
            rf.Caption = row_caption
            # TAT is both the counted field and a column field, so add the data field first.
            pt.AddDataField(pt.PivotFields("TAT"), "Count of TAT", xlCount)
            for i, f in enumerate(col_fields, 1):
                cf = pt.PivotFields(f)
                cf.Orientation = xlColumnField
                cf.Position = i
            return pt

        def order_months(pt):
            for pos, m in enumerate(months, 1):
                try:
                    pt.PivotFields("Month").PivotItems(m).Position = pos
                except Exception:
                    pass

        def find_cols(pt, header_row_offset, labels):
            tr = pt.TableRange1
            r0, c0 = tr.Row, tr.Column
            found = {}
            for c in range(c0, c0 + tr.Columns.Count):
                # Subtotal captions sit in the column-field rows above the last header row.
                for off in range(1, header_row_offset + 1):
                    v = summ.Cells(r0 + off, c).Value
                    if v in labels:
                        found[v] = c
            return found, r0, tr.Rows.Count, tr.Columns.Count, c0

        def add_pct(pt, header_row_offset, first_data_offset, specs):
            """specs: [(caption, numerator label, denominator label)] written right of the pivot."""
            cols, r0, nrows, ncols, c0 = find_cols(pt, header_row_offset, {l for _, n, d in specs for l in (n, d)})
            out_col = c0 + ncols
            hdr_row = r0 + header_row_offset
            for k, (cap, num, den) in enumerate(specs):
                c = out_col + k
                summ.Cells(hdr_row, c).Value = cap
                summ.Cells(hdr_row, c).Font.Bold = True
                for r in range(r0 + first_data_offset, r0 + nrows):
                    n_ref = summ.Cells(r, cols[num]).Address.replace("$", "")
                    d_ref = summ.Cells(r, cols[den]).Address.replace("$", "")
                    cell = summ.Cells(r, c)
                    cell.Formula = f"=IFERROR(N({n_ref})/{d_ref},0)"
                    cell.NumberFormat = "0.0%"

        def add_pct_by_month(pt, first_data_offset, month_list):
            """OT % per month and in total, via GETPIVOTDATA so it survives any pivot layout."""
            tr = pt.TableRange1
            r0, c0, nrows, ncols = tr.Row, tr.Column, tr.Rows.Count, tr.Columns.Count
            anchor = summ.Cells(r0, c0).Address
            hdr_row = r0 + 2
            specs = [(f"{t} % {m}", t, m) for t in ("OT",) for m in month_list + [None]]
            specs = [(cap if m else f"{t} % Total", t, m) for cap, t, m in specs]
            for k, (cap, tat, month) in enumerate(specs):
                c = c0 + ncols + k
                summ.Cells(hdr_row, c).Value = cap
                summ.Cells(hdr_row, c).Font.Bold = True
                for r in range(r0 + first_data_offset, r0 + nrows):
                    is_total = summ.Cells(r, c0).Value == "Grand Total"
                    label = "$" + summ.Cells(r, c0).Address.split("$")[1] + str(r)
                    def gpd(t):
                        parts = [f'"TAT","{t}"']
                        if month:
                            parts.append(f'"Month","{month}"')
                        if not is_total:
                            parts.append(f'"Check Name",{label}')
                        return f'IFERROR(GETPIVOTDATA("Count of TAT",{anchor},{",".join(parts)}),0)'
                    cell = summ.Cells(r, c)
                    cell.Formula = f"=IFERROR({gpd(tat)}/({gpd('IT')}+{gpd('OT')}),0)"
                    cell.NumberFormat = "0.0%"

        # Pivot 1: Month x TAT (IT / OT), plus OT %.
        p1 = make(3, "PivotMonthTAT", ["TAT"], "Month", "Month")
        order_months(p1)
        add_pct(p1, 1, 2, [("OT %", "OT", "Grand Total")])

        # Pivot 2: Check Name x (TAT > Month), plus IT % / OT %.
        p2_row = 3 + len(months) + 8
        p2 = make(p2_row, "PivotCheckTAT", ["TAT", "Month"], "check_name2", "Check Name")
        order_months(p2)
        add_pct_by_month(p2, 3, months)

        summ.Columns("A:A").ColumnWidth = 38
        summ.Columns("B:Q").ColumnWidth = 12
        summ.Activate()
        summ.Range("A1").Select()
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def to_excel(df: pd.DataFrame) -> tuple:
    """Return (xlsx bytes, has_pivot). Summary is a native Excel PivotTable when Excel is installed;
    otherwise a static summary sheet is written instead."""
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "report.xlsx")
        try:
            _write_data_workbook(df, path, with_static_summary=False)
            _add_pivots(path, _month_order(df))
            has_pivot = True
        except Exception:
            _write_data_workbook(df, path, with_static_summary=True)
            has_pivot = False
        with open(path, "rb") as f:
            return f.read(), has_pivot

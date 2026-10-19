import os
from pathlib import Path

import pandas as pd
import pymysql
import pythoncom
import win32com.client as win32
from dateutil import parser as date_parser
from dotenv import load_dotenv

XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_COUNT = -4112
XL_CENTER = -4108
XL_SRC_RANGE = 1
XL_YES = 1
LIGHT_GREY = 0xD9D9D9
MAX_COLUMN_WIDTH = 60
ROW_HEIGHT = 20

load_dotenv()

DB_HOST = os.getenv("DB_HOST")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_NAME = os.getenv("DB_NAME")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")

CLIENT_ID = 472826
CLIENT_CODE = "7084"

# Persistent Systems Limited Cleint Code-7084 (main tracker data). The literal
# 'Insufficiency_Raised' placeholder column from the portal query is omitted: the portal's
# CSV export collapsed it into the real subquery column of the same name.
MAIN_QUERY = """
SELECT
    ecf.case_flex_field1 AS 'EMP_Code',
    CONCAT(first_name,' ',middle_name,' ',last_name) AS 'Name',
    case_ars_no AS 'Vendor_ID',
    received_date AS 'IN_Date',
    `fn_case_status`(case_status) AS 'Overall_Case_status',
    DATE(`fn_latest_ir_date`(ecm.case_id)) AS 'CBC_Report_Date',
    IF(fn_last_ir_severity(ecm.case_id) COLLATE utf8mb3_unicode_ci = 'Attention Required' COLLATE utf8mb3_unicode_ci, 'Amber',
       IF(fn_last_ir_severity(ecm.case_id) COLLATE utf8mb3_unicode_ci = 'Insufficient' COLLATE utf8mb3_unicode_ci, 'Amber',
          IF(fn_last_ir_severity(ecm.case_id) COLLATE utf8mb3_unicode_ci = 'No Response Received' COLLATE utf8mb3_unicode_ci, 'Green',
             fn_last_ir_severity(ecm.case_id)
          )
       )
    ) AS 'CBC Report Color Code',
    DATE(`fn_fr_date`(ecm.case_id)) AS 'Final Report Date',
    `fn_fr_severity`(ecm.case_id) AS 'Final Report Color Code',
    fn_latest_report_date(ecm.case_id) AS 'Latest_Report_Sent',
    fn_latest_report_severity(ecm.case_id) AS 'Latest_Report_Severity',
    (SELECT GROUP_CONCAT(CONCAT(check_name,'->',fn_check_status(check_status),'->',insuff_remarks))
       FROM ec_case_checks ecc WHERE ecm.case_id=ecc.case_id AND check_status=3) AS 'Insufficiency_Raised',
    (SELECT GROUP_CONCAT(check_name)
       FROM ec_case_checks ecc WHERE ecm.case_id=ecc.case_id AND check_status NOT IN (8,9,10,11)) AS 'Pending_Checks',
    (SELECT GROUP_CONCAT(check_name)
       FROM ec_case_checks ecc WHERE ecm.case_id=ecc.case_id AND check_status <> 9) AS 'Inititated_Checks'
FROM ec_case_master ecm
LEFT JOIN ec_case_fields ecf ON ecm.case_id=ecf.case_id
LEFT JOIN ec_case_candidates ec ON ecm.candidate_id=ec.candidate_id
WHERE ecm.client_id=%s AND
    CONCAT(first_name,' ',middle_name,' ',last_name) NOT LIKE '%%Test%%' AND
    CONCAT(first_name,' ',middle_name,' ',last_name) NOT LIKE '%%Dummy%%'
    AND case_status NOT IN (8,14)
"""

# Client Wise All Cases - Bridge, trimmed to the two columns used for the Date of Joining
# lookup (case_flex_field2 holds the raw joining date); joins and filters are unchanged.
TRACKER_QUERY = """
SELECT case_ars_no, ecf3.case_flex_field2
FROM ec_case_master ecm
INNER JOIN ec_master_company emc ON ecm.client_id=emc.company_id
INNER JOIN ec_client ec ON emc.company_id=ec.client_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
INNER JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_case_fields ecf3 ON ecm.case_id=ecf3.case_id
WHERE ec.client_external_id IN (%s) AND first_name <> 'dummy'
"""

DOWNLOAD_DIR = Path(__file__).resolve().parent
REPORT_PATH = DOWNLOAD_DIR / "Persistent MIS Tracker Report.xlsx"

COLUMN_NAMES = [
    "Emp Code",
    "Name",
    "Vendor ID",
    "In Date",
    "Overall case status",
    "CBC Report Date",
    "CBC Report Color Code",
    "Final Report Date",
    "Final Report Colour code",
    "Latest_Report_Sent",
    "Latest_Report_Severity",
    "Insufficiency Raised",
    "Pending Checks",
    "Initiated Checks",
]

EXCLUDED_STATUSES = ["Duplicity", "Closed by Authbridge"]


def run_query(sql, params):
    connection = pymysql.connect(
        host=DB_HOST,
        port=DB_PORT,
        user=DB_USER,
        password=DB_PASSWORD,
        database=DB_NAME,
        charset="utf8",
    )
    try:
        with connection.cursor() as cursor:
            cursor.execute(sql, params)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()
    finally:
        connection.close()
    return columns, rows


def fetch_main_data():
    _, rows = run_query(MAIN_QUERY, (CLIENT_ID,))
    # Stringify the same way the portal's CSV export did (ISO dates, blanks for NULL) so the
    # Data Tab cell types match earlier reports.
    return pd.DataFrame(
        [["" if v is None else str(v) for v in row] for row in rows],
        columns=COLUMN_NAMES[:14],
    )


def fetch_tracker_data():
    columns, rows = run_query(TRACKER_QUERY, (CLIENT_CODE,))
    return pd.DataFrame(
        [["" if v is None else str(v) for v in row] for row in rows],
        columns=["case_ars_no", "case_flex_field2"],
    )


def parse_doj(value):
    value = value.strip()
    if not value or value.lower() == "not mentioned":
        return None
    try:
        return date_parser.parse(value).date()
    except (ValueError, OverflowError):
        return None


def build_doj_map(df):
    doj_map = {}
    for vendor_id, raw_value in zip(df["case_ars_no"], df["case_flex_field2"]):
        vendor_id = vendor_id.strip()
        if not vendor_id or vendor_id in doj_map:
            continue
        parsed = parse_doj(raw_value)
        if parsed is not None:
            doj_map[vendor_id] = parsed
    return doj_map


def normalize_emp_code(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.isdigit():
        return int(text)
    return text


def build_report_workbook(df, doj_map):
    df = df.copy()
    df.columns = COLUMN_NAMES
    df = df.mask(df == "")
    df["Emp Code"] = df["Emp Code"].map(normalize_emp_code)
    df["Date of joining"] = df["Vendor ID"].astype(str).str.strip().map(doj_map)
    df = df[~df["Overall case status"].isin(EXCLUDED_STATUSES)]
    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data Tab", index=False)
    return REPORT_PATH


def add_pivot_table(pivot_cache, sheet, anchor_cell, title, field_name):
    sheet.Range(anchor_cell).Value = title
    row, col = sheet.Range(anchor_cell).Row, sheet.Range(anchor_cell).Column
    title_range = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col + 1))
    title_range.Merge()
    title_range.Font.Bold = True
    title_range.HorizontalAlignment = XL_CENTER

    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=sheet.Cells(row + 2, col),
        TableName=f"{field_name.replace(' ', '')}Pivot",
    )
    pivot_table.TableStyle2 = "PivotStyleMedium9"
    row_field = pivot_table.PivotFields(field_name)
    row_field.Orientation = XL_ROW_FIELD
    pivot_table.AddDataField(pivot_table.PivotFields("Vendor ID"), "Count", XL_COUNT)

    for item in row_field.PivotItems():
        item.Visible = item.Name not in ("(blank)", "")

    sheet.Cells(row + 2, col).Value = "Case Status"
    header_range = sheet.Range(sheet.Cells(row + 2, col), sheet.Cells(row + 2, col + 1))
    header_range.Font.Bold = True
    header_range.Interior.Color = LIGHT_GREY


def build_pivot_summary(report_path):
    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(report_path))
        data_sheet = workbook.Sheets("Data Tab")
        last_row = data_sheet.Cells(data_sheet.Rows.Count, 1).End(-4162).Row
        last_col = data_sheet.Cells(1, data_sheet.Columns.Count).End(-4159).Column
        source_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(last_row, last_col))

        data_table = data_sheet.ListObjects.Add(XL_SRC_RANGE, source_range, None, XL_YES)
        data_table.Name = "DataTab"
        data_table.TableStyle = "TableStyleMedium9"
        header_range = data_sheet.Range(data_sheet.Cells(1, 1), data_sheet.Cells(1, last_col))
        header_range.HorizontalAlignment = XL_CENTER

        source_range.EntireColumn.AutoFit()
        for column in source_range.Columns:
            if column.ColumnWidth > MAX_COLUMN_WIDTH:
                column.ColumnWidth = MAX_COLUMN_WIDTH
                column.WrapText = True
        source_range.EntireRow.RowHeight = ROW_HEIGHT

        data_sheet.Activate()
        data_sheet.Range("A2").Select()
        excel.ActiveWindow.FreezePanes = True

        for sheet in workbook.Sheets:
            if sheet.Name == "Summary":
                sheet.Delete()
                break

        summary_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
        summary_sheet.Name = "Summary"

        pivot_cache = workbook.PivotCaches().Create(SourceType=XL_DATABASE, SourceData=source_range)
        add_pivot_table(pivot_cache, summary_sheet, "A1", "Overall Case wise Status", "Overall case status")
        add_pivot_table(pivot_cache, summary_sheet, "E1", "CBC  Severity", "CBC Report Color Code")

        summary_sheet.Columns("A:F").AutoFit()
        workbook.Save()
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


def generate_report():
    main_df = fetch_main_data()
    doj_map = build_doj_map(fetch_tracker_data())
    report_path = build_report_workbook(main_df, doj_map)
    build_pivot_summary(report_path)
    return report_path


def main():
    report_path = generate_report()
    print(f"Report saved to: {report_path}")


if __name__ == "__main__":
    main()

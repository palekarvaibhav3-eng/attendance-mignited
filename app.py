from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import calendar
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.chart import BarChart, Reference

# ==========================================================
# FLASK SETUP
# ==========================================================
app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

OUTPUT_FILE = os.path.join(UPLOAD_FOLDER, "Identix_Attendance_Report.xlsx")

COMPANY_NAME = "Mignited Technologies and Solutions Private Limited"

EMPLOYEES = {
    1: "Mukund Muley",
    5: "Pralav Phakatkar",
    8: "Rohit Nimse",
    9: "Samruddhi Kulkarni",
    10: "Mohini Shendge",
    11: "Shantanu Kulkarni"
}

# ==========================================================
# LOAD DATA
# ==========================================================
def load_data(input_file):
    df = pd.read_csv(input_file, sep=r"\s+", engine="python", header=None)

    df.columns = [
        "EmployeeID", "Date", "Time",
        "VerifyMode", "InOut", "WorkCode", "Reserved"
    ]

    df["DateTime"] = pd.to_datetime(df["Date"] + " " + df["Time"])
    df["Date"] = df["DateTime"].dt.date
    df["Month"] = df["DateTime"].dt.strftime("%B-%Y")

    return df.sort_values(["EmployeeID", "DateTime"])


# ==========================================================
# DAILY ATTENDANCE
# ==========================================================
def calculate_daily(df):

    daily = (
        df.groupby(["EmployeeID", "Date", "Month"])
        .agg(
            In_Time=("DateTime", "first"),
            Out_Time=("DateTime", "last")
        )
        .reset_index()
    )

    daily["Work_Hours"] = daily["Out_Time"] - daily["In_Time"]

    return daily


# ==========================================================
# GET MONTH DATES
# ==========================================================
def get_month_dates(daily):

    first_date = pd.to_datetime(daily["Date"]).min()
    year = first_date.year
    month = first_date.month

    total_days = calendar.monthrange(year, month)[1]

    month_dates = [
        datetime(year, month, day)
        for day in range(1, total_days + 1)
    ]

    sundays = [d for d in month_dates if d.weekday() == 6]

    return month_dates, sundays


# ==========================================================
# MATRIX + SUMMARY
# ==========================================================
def create_matrix_summary(daily, month_dates, sundays):

    matrix_rows = []
    summary_rows = []

    total_days = len(month_dates)
    total_holidays = len(sundays)
    base_working_days = total_days - total_holidays

    for emp_id, emp_name in EMPLOYEES.items():

        row = {"Employee Name": emp_name}
        emp_dates = set(
            daily[daily["EmployeeID"] == emp_id]["Date"]
        )

        present_count = 0
        sunday_work = 0

        for d in month_dates:

            col = d.strftime("%Y-%m-%d")

            if d.date() in emp_dates:
                row[col] = "P"
                present_count += 1

                if d.weekday() == 6:
                    sunday_work += 1
            else:
                row[col] = "WO" if d.weekday() == 6 else "A"

        total_working_days = base_working_days + sunday_work
        absent = max(total_working_days - present_count, 0)

        attendance_percent = round(
            (present_count / total_working_days) * 100, 2
        ) if total_working_days else 0

        summary_rows.append({
            "Employee Name": emp_name,
            "Total Days in Month": total_days,
            "Total Holidays": total_holidays,
            "Total Working Days": total_working_days,
            "Present Days": present_count,
            "Absent Days": absent,
            "Attendance %": attendance_percent
        })

        matrix_rows.append(row)

    return pd.DataFrame(matrix_rows), pd.DataFrame(summary_rows)


# ==========================================================
# WRITE EXCEL
# ==========================================================
def write_excel(daily, matrix, summary):

    if os.path.exists(OUTPUT_FILE):
        try:
            os.remove(OUTPUT_FILE)
        except PermissionError:
            print("Please close the Excel file and try again.")
            return

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

        # Detailed Sheets
        for month, month_df in daily.groupby("Month"):

            rows = []

            for emp_id, emp_data in month_df.groupby("EmployeeID"):

                emp_name = EMPLOYEES.get(emp_id, "Unknown")

                rows.append({
                    "Employee": f"{emp_id} - {emp_name}",
                    "Date": "",
                    "In Time": "",
                    "Out Time": "",
                    "Work Hours": ""
                })

                for _, r in emp_data.iterrows():
                    rows.append({
                        "Employee": "",
                        "Date": r["Date"],
                        "In Time": r["In_Time"].strftime("%H:%M:%S"),
                        "Out Time": r["Out_Time"].strftime("%H:%M:%S"),
                        "Work Hours": str(r["Work_Hours"])
                    })

                rows.append({})

            pd.DataFrame(rows).to_excel(writer, sheet_name=month[:31], index=False)

        # Matrix & Summary Per Month
        for month in summary["Month"].unique():

            month_matrix = matrix[matrix["Month"] == month].drop(columns=["Month"])
            month_summary = summary[summary["Month"] == month].drop(columns=["Month"])

            month_matrix.to_excel(
                writer,
                sheet_name=f"{month}_Matrix"[:31],
                index=False
            )

            month_summary.to_excel(
                writer,
                sheet_name=f"{month}_Summary"[:31],
                index=False
            )


# ==========================================================
# FORMAT EXCEL
# ==========================================================
def format_excel():

    wb = load_workbook(OUTPUT_FILE)

    header_fill = PatternFill("solid", fgColor="305496")
    green_fill = PatternFill("solid", fgColor="C6EFCE")
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    grey_fill = PatternFill("solid", fgColor="D9D9D9")

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for sheet in wb.sheetnames:

        ws = wb[sheet]

        if sheet.endswith("_Matrix"):
            ws.freeze_panes = "B2"
        else:
            ws.freeze_panes = "A2"

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")

        if sheet.endswith("_Matrix"):
            for row in ws.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    if cell.value == "P":
                        cell.fill = green_fill
                    elif cell.value == "A":
                        cell.fill = red_fill
                    elif cell.value == "WO":
                        cell.fill = grey_fill

        for col in ws.columns:
            max_len = max(
                len(str(cell.value)) if cell.value else 0
                for cell in col
            )
            ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 12)

    wb.save(OUTPUT_FILE)


# ==========================================================
# DASHBOARD
# ==========================================================
def create_dashboard():

    wb = load_workbook(OUTPUT_FILE)

    for sheet in wb.sheetnames:

        if sheet.endswith("_Summary"):

            summary_ws = wb[sheet]
            month = sheet.replace("_Summary", "")
            dashboard_name = f"Dashboard_{month}"

            if dashboard_name in wb.sheetnames:
                del wb[dashboard_name]

            ws = wb.create_sheet(dashboard_name, 0)

            total_employees = summary_ws.max_row - 1

            names = []
            percentages = []

            for row in summary_ws.iter_rows(min_row=2, values_only=True):
                names.append(row[0])
                percentages.append(row[6])

            avg_att = round(sum(percentages) / len(percentages), 2)
            best = names[percentages.index(max(percentages))]
            lowest = names[percentages.index(min(percentages))]

            ws.merge_cells("A1:H1")
            ws["A1"] = f"{COMPANY_NAME} - {month} Dashboard"
            ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
            ws["A1"].alignment = Alignment(horizontal="center")
            ws["A1"].fill = PatternFill("solid", fgColor="305496")

            ws["A3"] = "Total Employees:"
            ws["B3"] = total_employees

            ws["A4"] = "Average Attendance %:"
            ws["B4"] = avg_att

            ws["A5"] = "Best Performer:"
            ws["B5"] = best

            ws["A6"] = "Lowest Attendance:"
            ws["B6"] = lowest

            for r in summary_ws.iter_rows(values_only=True):
                ws.append(r)

            data_start_row = ws.max_row - total_employees + 1
            data_end_row = ws.max_row

            chart = BarChart()
            chart.title = "Attendance % by Employee"
            chart.y_axis.title = "Attendance %"

            data = Reference(ws,
                             min_col=7,
                             min_row=data_start_row,
                             max_row=data_end_row)

            cats = Reference(ws,
                             min_col=1,
                             min_row=data_start_row,
                             max_row=data_end_row)

            chart.add_data(data)
            chart.set_categories(cats)

            ws.add_chart(chart, "J3")

    wb.save(OUTPUT_FILE)


# ==========================================================
# REPORT GENERATOR
# ==========================================================
def generate_report(input_file):

    df = load_data(input_file)
    daily = calculate_daily(df)

    all_matrix = []
    all_summary = []

    for month, month_df in daily.groupby("Month"):

        month_dates, sundays = get_month_dates(month_df)
        matrix, summary = create_matrix_summary(month_df, month_dates, sundays)

        matrix.insert(0, "Month", month)
        summary.insert(0, "Month", month)

        all_matrix.append(matrix)
        all_summary.append(summary)

    final_matrix = pd.concat(all_matrix, ignore_index=True)
    final_summary = pd.concat(all_summary, ignore_index=True)

    write_excel(daily, final_matrix, final_summary)
    format_excel()
    create_dashboard()

    return OUTPUT_FILE


# ==========================================================
# FLASK ROUTES
# ==========================================================
@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    file = request.files["file"]

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        output_path = generate_report(filepath)

        return render_template("dashboard.html",
                               filename="Identix_Attendance_Report.xlsx")

    return "No file uploaded"


@app.route("/download/<filename>")
def download_file(filename):
    return send_file(os.path.join(UPLOAD_FOLDER, filename),
                     as_attachment=True)

    return "No file uploaded"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
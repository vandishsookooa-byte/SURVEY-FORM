from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from openpyxl import Workbook, load_workbook

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "RT_Knits_Transport_Survey_Responses.xlsx"

HEADERS = [
    "submittedAt",
    "employeeId",
    "shiftTiming",
    "departments",
    "pickupDropRating",
    "overtimeTransportRating",
    "overallConveyanceSatisfaction",
    "transferPlanningRating",
    "logisticsRating",
    "overallProductionSatisfaction",
    "comments",
]


def ensure_excel_file() -> None:
    if EXCEL_FILE.exists():
        return
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Survey Responses"
    sheet.append(HEADERS)
    workbook.save(EXCEL_FILE)


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.post("/submit")
def submit():
    data = request.get_json(silent=True) or {}

    if not data.get("shiftTiming"):
        return jsonify({"message": "Shift timing is required."}), 400
    if not data.get("departments"):
        return jsonify({"message": "At least one department is required."}), 400

    ensure_excel_file()
    workbook = load_workbook(EXCEL_FILE)
    sheet = workbook.active
    row = [
        datetime.now().isoformat(timespec="seconds"),
        str(data.get("employeeId", "")).strip(),
        str(data.get("shiftTiming", "")).strip(),
        str(data.get("departments", "")).strip(),
        str(data.get("pickupDropRating", "")).strip(),
        str(data.get("overtimeTransportRating", "")).strip(),
        str(data.get("overallConveyanceSatisfaction", "")).strip(),
        str(data.get("transferPlanningRating", "")).strip(),
        str(data.get("logisticsRating", "")).strip(),
        str(data.get("overallProductionSatisfaction", "")).strip(),
        str(data.get("comments", "")).strip(),
    ]
    sheet.append(row)
    workbook.save(EXCEL_FILE)

    return jsonify(
        {
            "message": "Anonymous response submitted successfully. Excel has been updated.",
        }
    )


if __name__ == "__main__":
    ensure_excel_file()
    app.run(host="0.0.0.0", port=5000, debug=False)

import os
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request
from openpyxl import Workbook, load_workbook

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "RT_Knits_Transport_Survey_Responses.xlsx"
app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")

HEADERS = [
    "submittedAt",
    "employeeName",
    "employeeId",
    "departments",
    "conveyanceAllocationRating",
    "specialTransportRating",
    "conveyanceResponsivenessRating",
    "overallConveyanceSatisfaction",
    "productionPlanningRating",
    "logisticsRequestRating",
    "productionResponsivenessRating",
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
    return app.send_static_file("index.html")


@app.get("/favicon.ico")
def favicon():
    return ("", 204)


@app.post("/submit")
def submit():
    data = request.get_json(silent=True) or {}

    if not data.get("departments"):
        return jsonify({"message": "At least one department is required."}), 400
    if not data.get("productionPlanningRating"):
        return jsonify({"message": "Production planning rating is required."}), 400
    if not data.get("logisticsRequestRating"):
        return jsonify({"message": "Logistics request management rating is required."}), 400
    if not data.get("productionResponsivenessRating"):
        return jsonify({"message": "Production responsiveness rating is required."}), 400
    if not data.get("overallProductionSatisfaction"):
        return jsonify({"message": "Overall production satisfaction is required."}), 400

    ensure_excel_file()
    workbook = load_workbook(EXCEL_FILE)
    sheet = workbook.active
    row = [
        datetime.now().isoformat(timespec="seconds"),
        str(data.get("employeeName", "")).strip(),
        str(data.get("employeeId", "")).strip(),
        str(data.get("departments", "")).strip(),
        str(data.get("conveyanceAllocationRating", "")).strip(),
        str(data.get("specialTransportRating", "")).strip(),
        str(data.get("conveyanceResponsivenessRating", "")).strip(),
        str(data.get("overallConveyanceSatisfaction", "")).strip(),
        str(data.get("productionPlanningRating", "")).strip(),
        str(data.get("logisticsRequestRating", "")).strip(),
        str(data.get("productionResponsivenessRating", "")).strip(),
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
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)

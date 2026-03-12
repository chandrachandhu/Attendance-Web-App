from flask import Flask, request, send_file
from attendance_logic import process_attendance
import os
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Google Sheet CSV export link
SHEET_URL = "https://docs.google.com/spreadsheets/d/1t1OyQbuq0kD3F9vFpo6c6ScEN-A0UWUGaYQ214qIeXw/export?format=csv"

@app.route("/")
def home():
    return open("templates/index.html").read()


# ==============================
# Extract data from Google Sheet
# ==============================
@app.route("/extract")
def extract():

    try:
        df = pd.read_csv(SHEET_URL)

        input_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")

        df.to_excel(input_path, index=False)

        return "Attendance data extracted successfully"

    except Exception as e:
        return f"Error extracting data: {str(e)}"


# ==============================
# Generate attendance
# ==============================
@app.route("/generate", methods=["POST"])
def generate():

    filename = request.form["filename"]

    input_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")

    output_path = os.path.join(
        UPLOAD_FOLDER,
        f"{filename}_attendance.xlsx"
    )

    process_attendance(input_path, output_path)

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

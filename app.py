from flask import Flask, request, send_file, render_template
from attendance_logic import process_attendance
import os
import pandas as pd
import requests

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Google Sheet Excel export link
SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTPYebkj-NlcHEdeEPIZEsEpu2F6qplpVmggh6QWxv1kkgarlWC2WvqXW9_10arvmtlVGFCl-rM3Dwm/pub?output=xlsx"


@app.route("/")
def home():
    return render_template("index.html")


# Extract data
@app.route("/extract")
def extract():
    try:

        temp_path = os.path.join(UPLOAD_FOLDER, "temp.xlsx")
        input_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")

        # Download Excel from Google Sheets
        response = requests.get(SHEET_URL)
        response.raise_for_status()

        with open(temp_path, "wb") as f:
            f.write(response.content)

        # Read sheet
        df = pd.read_excel(temp_path, header=None, dtype=str)

        # Save with correct sheet name
        with pd.ExcelWriter(input_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet2", index=False, header=False)

        return "Attendance data extracted successfully!"

    except Exception as e:
        return f"Error extracting data: {str(e)}"


# Generate attendance
@app.route("/generate", methods=["POST"])
def generate():
    try:

        filename = request.form["filename"]

        input_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")

        output_path = os.path.join(
            UPLOAD_FOLDER,
            f"{filename}_attendance.xlsx"
        )

        process_attendance(input_path, output_path)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error generating attendance: {str(e)}"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

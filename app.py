from flask import Flask, request, send_file
from attendance_logic import process_attendance
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return open("templates/index.html").read()

@app.route("/generate", methods=["POST"])
def generate():
    file = request.files["file"]
    filename = request.form["filename"]

    input_path = os.path.join(UPLOAD_FOLDER, "input.xlsx")

    output_path = os.path.join(
        UPLOAD_FOLDER,
        f"{filename}_attendance.xlsx"
    )

    file.save(input_path)

    process_attendance(input_path, output_path)

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename
import os
import comtypes.client
import pythoncom   # Required for COM threading
from fpdf import FPDF   # fallback for TXT files

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def docx_to_pdf(input_path, output_path):
    """Convert DOCX to PDF using Microsoft Word COM API (thread-safe)."""
    pythoncom.CoInitialize()

    # Ensure absolute Windows paths
    abs_input = os.path.abspath(input_path)
    abs_output = os.path.abspath(output_path)

    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False

    # Open the DOCX and export as PDF
    doc = word.Documents.Open(abs_input)
    doc.SaveAs(abs_output, FileFormat=17)  # 17 = wdFormatPDF
    doc.Close()
    word.Quit()

    pythoncom.CoUninitialize()


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert_file():
    if "file" not in request.files:
        return "❌ No file uploaded"
    file = request.files["file"]
    if file.filename == "":
        return "❌ No file selected"

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # DOCX → PDF
    if filename.lower().endswith(".docx"):
        output_path = filepath.replace(".docx", ".pdf")
        docx_to_pdf(filepath, output_path)

    # TXT → PDF (simple fallback)
    elif filename.lower().endswith(".txt"):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                pdf.multi_cell(0, 10, line.strip())
        output_path = filepath + ".pdf"
        pdf.output(output_path)

    else:
        return "❌ Unsupported file type. Please upload .docx or .txt"

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, port=8080)

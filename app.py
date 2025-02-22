from flask import Flask, request, jsonify, send_file, render_template
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import uuid
from docx import Document
from docx2pdf import convert
import pythoncom
import win32com.client

app = Flask(__name__, static_folder="static", template_folder="templates")


def generate_certificate_number():
    return f"MYC-002-{str(uuid.uuid4().int)[:6]}"


# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open_by_key("1hxpUDxgR99fBT8M_7q-APOQe8QWgO36AoOvbyloO64Y").sheet1


@app.route("/")
def home():
    return render_template("newform.html")


@app.route("/register", methods=["POST"])
def register():
    data = request.json
    email = data.get("email")

    # Check if user is already registered
    records = sheet.get_all_records()
    for record in records:
        if record["Email"] == email:
            return jsonify({"error": "User already registered"}), 400

    # Generate unique certification number
    cert_number = generate_certificate_number()

    # Save data to Google Sheets
    sheet.append_row([
        data.get("firstname"),
        data.get("lastname"),
        email,
        data.get("phone"),
        data.get("address"),
        data.get("address2"),
        data.get("state"),
        data.get("country"),
        data.get("post"),
        data.get("area"),
        cert_number
    ])

    # Generate Certificate PDF
    doc = Document("D:/trial/formfinal/certificate template.docx")
    for para in doc.paragraphs:
        if "{{NAME}}" in para.text:
            para.text = para.text.replace("{{NAME}}", data.get("firstname"))
        if "{{SURNAME}}" in para.text:
            para.text = para.text.replace("{{SURNAME}}", data.get("lastname"))
        if "{{CERT_NUMBER}}" in para.text:
            para.text = para.text.replace("{{CERT_NUMBER}}", cert_number)

        # Save the modified Word document
    cert_path = f"certificates/{cert_number}.docx"
    pdf_path = cert_path.replace(".docx", ".pdf")
    doc.save(cert_path)

    # Initialize COM before conversion
    pythoncom.CoInitialize()
    convert(cert_path)
    pythoncom.CoUninitialize()

    return send_file(pdf_path, as_attachment=True, download_name="certificate.pdf")
if __name__ == "__main__":
    app.run(debug=True)

from flask import Flask, render_template, request, send_file
import os
from report_generator import generate_vat_report

app = Flask(__name__)

# Upload folder
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Create uploads folder if not exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# -----------------------------------
# HOME PAGE
# -----------------------------------

@app.route("/")
def index():
    return render_template("upload.html")

# -----------------------------------
# GENERATE REPORT
# -----------------------------------

@app.route("/generate", methods=["POST"])
def generate():

    # Get uploaded files
    invoice = request.files["invoice"]
    bill = request.files["bill"]
    expense = request.files["expense"]
    vendor_credit = request.files["vendor_credit"]
    credit_note = request.files["credit_note"]

    # Save files
    invoice_path = os.path.join(app.config["UPLOAD_FOLDER"], invoice.filename)
    bill_path = os.path.join(app.config["UPLOAD_FOLDER"], bill.filename)
    expense_path = os.path.join(app.config["UPLOAD_FOLDER"], expense.filename)
    vendor_path = os.path.join(app.config["UPLOAD_FOLDER"], vendor_credit.filename)
    credit_path = os.path.join(app.config["UPLOAD_FOLDER"], credit_note.filename)

    invoice.save(invoice_path)
    bill.save(bill_path)
    expense.save(expense_path)
    vendor_credit.save(vendor_path)
    credit_note.save(credit_path)

    # -----------------------------------
    # CALL REPORT GENERATOR
    # -----------------------------------

    report_path = os.path.join(app.config["UPLOAD_FOLDER"], "VAT_Report.xlsx")

    generate_vat_report(
        invoice_path,
        bill_path,
        expense_path,
        vendor_path,
        credit_path,
        report_path
    )

    # -----------------------------------
    # DOWNLOAD REPORT
    # -----------------------------------

    return send_file(report_path, as_attachment=True)

# -----------------------------------
# RUN APP
# -----------------------------------

if __name__ == "__main__":
    app.run(debug=True)
import pandas as pd


# ---------------------------------------------------
# FIND COLUMN BY POSSIBLE NAMES
# ---------------------------------------------------
def find_column(df, possible_names):

    for col in df.columns:
        for name in possible_names:
            if name.lower() in col.lower():
                return col

    return None


# ---------------------------------------------------
# SAFE FILTER FUNCTION
# ---------------------------------------------------
def safe_filter(df, column, value):

    if column in df.columns:
        return df[df[column].astype(str).str.lower().str.contains(value.lower())]

    return df


# ---------------------------------------------------
# MAIN VAT REPORT FUNCTION
# ---------------------------------------------------
def generate_vat_report(invoice_path, bill_path, expense_path, vendor_credit_path, credit_note_path, output_path):

    # -----------------------------
    # READ FILES
    # -----------------------------
    invoice = pd.read_excel(invoice_path)
    bill = pd.read_excel(bill_path)
    expense = pd.read_excel(expense_path)
    credit_note = pd.read_excel(credit_note_path)
    vendor_credit = pd.read_excel(vendor_credit_path)

    # remove extra spaces from column names
    for df in [invoice, bill, expense, credit_note, vendor_credit]:
        df.columns = df.columns.str.strip()

    # -----------------------------
    # FIND IMPORTANT COLUMNS
    # -----------------------------
    invoice_tax = find_column(invoice, ["item tax amount"])
    invoice_total = find_column(invoice, ["item total"])

    bill_tax = find_column(bill, ["tax amount"])
    bill_total = find_column(bill, ["item total"])

    credit_tax = find_column(credit_note, ["tax amount"])
    credit_total = find_column(credit_note, ["item total"])

    vendor_credit_tax = find_column(vendor_credit, ["item tax amount"])
    vendor_credit_total = find_column(vendor_credit, ["item total"])

    expense_tax = find_column(expense, ["tax amount"])
    expense_total = find_column(expense, ["expense amount"])

    # -----------------------------
    # SAFE FILTERING
    # -----------------------------
    invoice = safe_filter(invoice, "VAT Treatment", "vat")
    bill = safe_filter(bill, "VAT Treatment", "vat")
    credit_note = safe_filter(credit_note, "VAT Treatment", "vat")
    vendor_credit = safe_filter(vendor_credit, "VAT Treatment", "vat")

    expense = safe_filter(expense, "ITC Eligibility", "eligible")

    # -----------------------------
    # DEBUG ROW COUNT
    # -----------------------------
    print("Invoice rows:", len(invoice))
    print("Bill rows:", len(bill))
    print("Credit note rows:", len(credit_note))
    print("Vendor credit rows:", len(vendor_credit))
    print("Expense rows:", len(expense))

    # -----------------------------
    # GROUP DATA
    # -----------------------------
    invoice_summary = pd.DataFrame()
    bill_summary = pd.DataFrame()
    credit_summary = pd.DataFrame()
    vendor_credit_summary = pd.DataFrame()
    expense_summary = pd.DataFrame()

    if invoice_tax and invoice_total:
        invoice_summary = invoice.groupby(
            ["Customer Name", "Invoice Date", "Invoice Number"],
            as_index=False
        )[[invoice_total, invoice_tax]].sum()

    if credit_tax and credit_total:
        credit_summary = credit_note.groupby(
            ["Customer Name", "Credit Note Date", "Credit Note Number"],
            as_index=False
        )[[credit_total, credit_tax]].sum()

    if bill_tax and bill_total:
        bill_summary = bill.groupby(
            ["Vendor Name", "Bill Date", "Bill Number"],
            as_index=False
        )[[bill_total, bill_tax]].sum()

    if vendor_credit_tax and vendor_credit_total:
        vendor_credit_summary = vendor_credit.groupby(
            ["Vendor Name", "Vendor Credit Date", "Vendor Credit Number"],
            as_index=False
        )[[vendor_credit_total, vendor_credit_tax]].sum()

    if expense_tax and expense_total:
        expense_summary = expense.groupby(
            ["Expense Date"],
            as_index=False
        )[[expense_total, expense_tax]].sum()

    # -----------------------------
    # WRITE REPORT
    # -----------------------------
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        sheet = "VAT Report"
        row = 0

        pd.DataFrame([["VAT REPORT"]]).to_excel(
            writer,
            sheet_name=sheet,
            startrow=row,
            index=False,
            header=False
        )

        row += 2

        # -----------------------------
        # CUSTOMER INVOICES
        # -----------------------------
        if not invoice_summary.empty:

            for cust, data in invoice_summary.groupby("Customer Name"):

                pd.DataFrame([[cust]]).to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False,
                    header=False
                )

                row += 1

                data = data.rename(columns={
                    "Invoice Date": "Date",
                    "Invoice Number": "Invoice No",
                    invoice_total: "Amount",
                    invoice_tax: "Tax"
                })

                data.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False
                )

                row += len(data) + 3

        # -----------------------------
        # CREDIT NOTES
        # -----------------------------
        if not credit_summary.empty:

            for cust, data in credit_summary.groupby("Customer Name"):

                pd.DataFrame([[f"Credit Notes - {cust}"]]).to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False,
                    header=False
                )

                row += 1

                data = data.rename(columns={
                    "Credit Note Date": "Date",
                    "Credit Note Number": "Credit No",
                    credit_total: "Amount",
                    credit_tax: "Tax"
                })

                data.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False
                )

                row += len(data) + 3

        # -----------------------------
        # VENDOR BILLS
        # -----------------------------
        if not bill_summary.empty:

            for vend, data in bill_summary.groupby("Vendor Name"):

                pd.DataFrame([[vend]]).to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False,
                    header=False
                )

                row += 1

                data = data.rename(columns={
                    "Bill Date": "Date",
                    "Bill Number": "Bill No",
                    bill_total: "Amount",
                    bill_tax: "Tax"
                })

                data.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False
                )

                row += len(data) + 3

        # -----------------------------
        # VENDOR CREDITS
        # -----------------------------
        if not vendor_credit_summary.empty:

            for vend, data in vendor_credit_summary.groupby("Vendor Name"):

                pd.DataFrame([[f"Vendor Credit - {vend}"]]).to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False,
                    header=False
                )

                row += 1

                data = data.rename(columns={
                    "Vendor Credit Date": "Date",
                    "Vendor Credit Number": "Credit No",
                    vendor_credit_total: "Amount",
                    vendor_credit_tax: "Tax"
                })

                data.to_excel(
                    writer,
                    sheet_name=sheet,
                    startrow=row,
                    index=False
                )

                row += len(data) + 3

        # -----------------------------
        # EXPENSES
        # -----------------------------
        if not expense_summary.empty:

            pd.DataFrame([["Eligible Expenses"]]).to_excel(
                writer,
                sheet_name=sheet,
                startrow=row,
                index=False,
                header=False
            )

            row += 1

            expense_summary = expense_summary.rename(columns={
                "Expense Date": "Date",
                expense_total: "Amount",
                expense_tax: "Tax"
            })

            expense_summary.to_excel(
                writer,
                sheet_name=sheet,
                startrow=row,
                index=False
            )

    return output_path

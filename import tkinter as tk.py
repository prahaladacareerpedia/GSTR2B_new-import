import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from lxml import etree as ET

# Function to generate Tally XML from Excel data
def create_tally_xml(data, output_file):
    root = ET.Element("ENVELOPE")
    header = ET.SubElement(root, "HEADER")
    ET.SubElement(header, "TALLYREQUEST").text = "Import Data"

    body = ET.SubElement(root, "BODY")
    import_data = ET.SubElement(body, "IMPORTDATA")
    request_desc = ET.SubElement(import_data, "REQUESTDESC")
    ET.SubElement(request_desc, "REPORTNAME").text = "Vouchers"

    # Set static variables like the company name if needed
    static_vars = ET.SubElement(request_desc, "STATICVARIABLES")
    ET.SubElement(static_vars, "SVCURRENTCOMPANY").text = "Your Company Name"

    request_data = ET.SubElement(import_data, "REQUESTDATA")

    grouped_data = data.groupby("Invoice number")

    for invoice_number, group in grouped_data:
        tally_message = ET.SubElement(request_data, "TALLYMESSAGE")
        voucher = ET.SubElement(tally_message, "VOUCHER", VCHTYPE="Purchase", ACTION="Create", OBJVIEW="Accounting Voucher View")

        first_entry = group.iloc[0]
        date_obj = pd.to_datetime(first_entry["Invoice Date"], dayfirst=True)
        formatted_date = date_obj.strftime('%Y%m%d')

        # Voucher level fields
        ET.SubElement(voucher, "DATE").text = formatted_date
        ET.SubElement(voucher, "VOUCHERTYPENAME").text = "Purchase"
        ET.SubElement(voucher, "VOUCHERNUMBER").text = str(invoice_number)
        ET.SubElement(voucher, "PARTYLEDGERNAME").text = first_entry["Trade/Legal name"]
        ET.SubElement(voucher, "PARTYGSTIN").text = first_entry["GSTIN of supplier"]
        ET.SubElement(voucher, "STATENAME").text = first_entry["Place of supply"]
        ET.SubElement(voucher, "COUNTRYOFRESIDENCE").text = "India"

        # Party ledger entry (Credit)
        party_ledger_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
        ET.SubElement(party_ledger_entry, "LEDGERNAME").text = first_entry["Trade/Legal name"]
        ET.SubElement(party_ledger_entry, "ISDEEMEDPOSITIVE").text = "No"
        party_amount = round(group["Taxable Value (₹)"].sum() + group["Central Tax(₹)"].sum() + group["State/UT Tax(₹)"].sum(), 2)
        ET.SubElement(party_ledger_entry, "AMOUNT").text = "{:.2f}".format(party_amount)

        # Bill allocation for the party
        bill_alloc = ET.SubElement(party_ledger_entry, "BILLALLOCATIONS.LIST")
        ET.SubElement(bill_alloc, "NAME").text = str(invoice_number)
        ET.SubElement(bill_alloc, "BILLTYPE").text = "Agst Ref"
        ET.SubElement(bill_alloc, "AMOUNT").text = "{:.2f}".format(party_amount)

        # Ledger entries for GST and other components (Debit)
        for _, row in group.iterrows():
            ledger_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(ledger_entry, "LEDGERNAME").text = str(row["Rate(%)"])
            ET.SubElement(ledger_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(ledger_entry, "AMOUNT").text = "-{:.2f}".format(row["Taxable Value (₹)"])

            # Central Tax
            central_tax_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(central_tax_entry, "LEDGERNAME").text = "Central Tax"
            ET.SubElement(central_tax_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(central_tax_entry, "AMOUNT").text = "-{:.2f}".format(row["Central Tax(₹)"])

            # State/UT Tax
            state_tax_entry = ET.SubElement(voucher, "ALLLEDGERENTRIES.LIST")
            ET.SubElement(state_tax_entry, "LEDGERNAME").text = "State/UT Tax(₹)"
            ET.SubElement(state_tax_entry, "ISDEEMEDPOSITIVE").text = "Yes"
            ET.SubElement(state_tax_entry, "AMOUNT").text = "-{:.2f}".format(row["State/UT Tax(₹)"])

    xml_str = ET.tostring(root, pretty_print=True, xml_declaration=True, encoding='UTF-8')
    with open(output_file, 'wb') as f:
        f.write(xml_str)

# Function to load Excel and convert to Tally XML
def load_excel_and_convert():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            data = pd.read_excel(file_path)
            output_file = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
            if output_file:
                create_tally_xml(data, output_file)
                messagebox.showinfo("Success", "Tally XML generated successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Setting up the Tkinter window
window = tk.Tk()
window.title("Excel to Tally XML Converter")
window.geometry("400x200")

# Adding a button to load Excel file and generate Tally XML
button = tk.Button(window, text="Load Excel and Generate Tally XML", command=load_excel_and_convert)
button.pack(pady=50)

# Running the Tkinter event loop
window.mainloop()

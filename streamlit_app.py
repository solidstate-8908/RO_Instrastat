
import streamlit as st
import pandas as pd
from yattag import Doc, indent
import os

def generate_xml(xml_admin, input_excel):
    # Read the Excel files
    inscode_xml_admin = pd.read_excel(xml_admin, sheet_name="InsCodeVersions", header=None)
    insdeclaration_xml_admin = pd.read_excel(xml_admin, sheet_name="InsDeclarationHeader", header=None)
    sales_ic_data = pd.read_excel(input_excel, sheet_name="Sales IC")

    # Initialize the XML document
    doc, tag, text = Doc().tagtext()
    xml_header = '<?xml version="1.0" encoding="UTF-8"?>'
    doc.asis(xml_header)

    with tag("InsNewDispatch", SchemaVersion=str("1.0"), xmlns=str("http://www.intrastat.ro/xml/InsSchema")):
        col_data = inscode_xml_admin.iloc[:, 1]
        tag_names = [
            "CountryVer",
            "EuCountryVer",
            "CnVer",
            "ModeOfTransportVer",
            "DeliveryTermsVer",
            "NatureOfTransactionAVer",
            "NatureOfTransactionBVer",
            "CountyVer",
            "LocalityVer",
            "UnitVer"
        ]
        with tag("InsCodeVersions"):
            for tag_name, value in zip(tag_names, col_data):
                with tag(tag_name):
                    text(str(value))

        col_data = insdeclaration_xml_admin.iloc[:, 1]
        tag_names = [
            "VatNr",
            "FirmName",
            "RefPeriod",
            "CreateDt",
            "LastName",
            "FirstName",
            "Email",
            "Phone",
            "Position"
        ]
        with tag("InsDeclarationHeader"):
            for tag_name, value in zip(tag_names, col_data):
                with tag(tag_name):
                    text(str(value))

        for _, row in sales_ic_data.iterrows():
            with tag("InsDispatchItem", OrderNr=str(row.name + 1)):
                with tag("Cn8Code"):
                    text(str(row["CodNC8"]))
                with tag("InvoiceValue"):
                    text(str(row["Sum of val facturata"]))
                with tag("StatisticalValue"):
                    text(str(row["Sum of val statistica"]))
                with tag("NetMass"):
                    text(str(row["cant"]))
                with tag("NatureOfTransactionACode"):
                    text(str(row["nat tranz A"]))
                with tag("NatureOfTransactionBCode"):
                    text(str(row["nat tranz B"]))
                with tag("DeliveryTermsCode"):
                    text(str(row["termeni livrare"]))
                with tag("ModeOfTransportCode"):
                    text(str(row["mod transport"]))
                with tag("CountryOfOrigin"):
                    text(str(row["tara origine"]))
                with tag("CountryOfDestination"):
                    text(str(row["tara de expediere"]))
                with tag("PartnerCountryCode"):
                    text(str(row["PartnerCountryCode"]))
                with tag("PartnerVatNr"):
                    text(str(row["PartnerVatNr"]))

    return indent(doc.getvalue(), indentation='   ', indent_text=False)

# Streamlit App
st.title("Romania Intrastat XML Generator")
st.write("Upload the required Excel files to generate the Intrastat XML.")

# File Uploads
xml_admin_file = st.file_uploader("Upload the XML Admin Excel File", type="xlsx")
input_excel_file = st.file_uploader("Upload the Input Excel File", type="xlsx")

# Output File Name
output_file_name = st.text_input("Enter the Output XML File Name", "sales_intrastat_output.xml")

if st.button("Generate XML"):
    if xml_admin_file and input_excel_file:
        try:
            # Generate XML
            xml_content = generate_xml(xml_admin_file, input_excel_file)

            # Save XML to a file
            with open(output_file_name, "w", encoding="utf-8") as f:
                f.write(xml_content)

            st.success(f"XML file '{output_file_name}' generated successfully!")

            # Provide download link
            with open(output_file_name, "rb") as f:
                st.download_button(
                    label="Download XML File",
                    data=f,
                    file_name=output_file_name,
                    mime="application/xml"
                )
        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload both Excel files.")

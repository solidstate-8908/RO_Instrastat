import streamlit as st
import pandas as pd
from yattag import Doc, indent
import zipfile
import io


def generate_exports_xml(xml_admin, input_excel):
    
    inscode_xml_admin = pd.read_excel(xml_admin, sheet_name="InsCodeVersions", header=None)
    insdeclaration_xml_admin = pd.read_excel(xml_admin, sheet_name="InsDeclarationHeader", header=None)
    exports_ic_data = pd.read_excel(input_excel, sheet_name="Sales IC")

    # Initialize the XML document for exports
    doc, tag, text = Doc().tagtext()
    xml_header = '<?xml version="1.0" encoding="UTF-8" ?>'
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
        # Map the data to corresponding tags
        data_map = {
            "VatNr": col_data[0],
            "FirmName": col_data[1],
            "RefPeriod": col_data[2],
            "CreateDt": col_data[3],
            "LastName": col_data[4],
            "FirstName": col_data[5],
            "Email": col_data[6],
            "Phone": col_data[7],
            "Position": col_data[8]
        }
        with tag("InsDeclarationHeader"):
            for key in ["VatNr", "FirmName", "RefPeriod", "CreateDt"]:
                with tag(key):
                    text(str(data_map[key]))
            with tag("ContactPerson"):
                for key in ["LastName", "FirstName", "Email", "Phone", "Position"]:
                    with tag(key):
                        text(str(data_map[key]))

        for _, row in exports_ic_data.iterrows():
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

def generate_imports_xml(xml_admin, input_excel):
    inscode_xml_admin = pd.read_excel(xml_admin, sheet_name="InsCodeVersions", header=None)
    insdeclaration_xml_admin = pd.read_excel(xml_admin, sheet_name="InsDeclarationHeader", header=None)
    imports_ic_data = pd.read_excel(input_excel, sheet_name="Aquisitions IC")
    
    # Initialize the XML document for imports
    doc, tag, text = Doc().tagtext()
    xml_header = '<?xml version="1.0" encoding="UTF-8" ?>'
    doc.asis(xml_header)

    with tag("InsNewArrival", SchemaVersion=str("1.0"), xmlns=str("http://www.intrastat.ro/xml/InsSchema")):
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
        # Map the data to corresponding tags
        data_map = {
            "VatNr": col_data[0],
            "FirmName": col_data[1],
            "RefPeriod": col_data[2],
            "CreateDt": col_data[3],
            "LastName": col_data[4],
            "FirstName": col_data[5],
            "Email": col_data[6],
            "Phone": col_data[7],
            "Position": col_data[8]
        }
        with tag("InsDeclarationHeader"):
            for key in ["VatNr", "FirmName", "RefPeriod", "CreateDt"]:
                with tag(key):
                    text(str(data_map[key]))
            with tag("ContactPerson"):
                for key in ["LastName", "FirstName", "Email", "Phone", "Position"]:
                    with tag(key):
                        text(str(data_map[key]))

        for _, row in imports_ic_data.iterrows():
                with tag("InsArrivalItem", OrderNr=str(row.name + 1)):
                    with tag("Cn8Code"):
                        text(str(row["CodNC8"]))
                    with tag("InvoiceValue"):
                        text(str(row["Total"]))
                    with tag("StatisticalValue"):
                        text(str(row["Total"]))
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
                    with tag("CountryOfConsignment"):
                        text(str(row["tara de expediere"]))

    return indent(doc.getvalue(), indentation='   ', indent_text=False)

# Streamlit App
st.title("Romania Intrastat XML Generator")
st.write("Upload the required Excel files to generate the Intrastat XML files for imports & exports.")

# File Uploads
xml_admin_file = st.file_uploader("Upload the XML Admin Excel File", type="xlsx")
input_excel_file = st.file_uploader("Upload the Input Excel File", type="xlsx")

# Output File Name
exports_output_fname = st.text_input("Enter the exports XML File Name including .xml at the end", "exports_intrastat_output.xml")
imports_output_fname = st.text_input("Enter the imports XML File Name including .xml at the end", "imports_intrastat_output.xml")

if st.button("Generate XML files"):
    if xml_admin_file and input_excel_file:
        try:
            # Generate XML files
            exports_xml_content = generate_exports_xml(xml_admin_file, input_excel_file)
            imports_xml_content = generate_imports_xml(xml_admin_file, input_excel_file)

            # Save XML files
            with open(exports_output_fname, "w", encoding="utf-8") as f:
                f.write(exports_xml_content)
            with open(imports_output_fname, "w", encoding="utf-8") as f:
                f.write(imports_xml_content)

            # Create ZIP in memory
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr(exports_output_fname, exports_xml_content)
                zip_file.writestr(imports_output_fname, imports_xml_content)

            zip_buffer.seek(0)
            st.success("XML files generated successfully!")
            # Provide single download button for ZIP
            st.download_button(
                label="Download XML Files as ZIP",
                data=zip_buffer,
                file_name="intrastat_xml_files.zip",
                mime="application/zip"
            )

        except Exception as e:
            st.error(f"An error occurred: {e}")
    else:
        st.warning("Please upload both Excel files.")


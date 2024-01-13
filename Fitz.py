import fitz  # PyMuPDF
import pandas as pd
import os
import traceback
import sys

# Function: Get User Input
def get_user_input(default_excel='test excel 1', default_template='template'):
    print("Please enter the required file names without extensions or type 'default' to use the preset names.")
    data_file = input(f"Enter the name of the Excel file without extension (default: '{default_excel}'): ").strip()
    if data_file.lower() == 'default' or not data_file:
        data_file = default_excel
    data_file += '.xlsx'

    template_file = input(f"Enter the name of the PDF template file without extension (default: '{default_template}'): ").strip()
    if template_file.lower() == 'default' or not template_file:
        template_file = default_template
    template_file += '.pdf'
    return data_file, template_file

# Function: Setup File Paths
def setup_file_paths(data_file='test excel 1.xlsx',
                     electrical_folder='Filled Electrical PDFs',
                     structural_folder='Filled Structural PDFs'):
    current_path = os.getcwd()
    data_path = os.path.join(current_path, data_file)
    electrical_output_folder = os.path.join(current_path, electrical_folder)
    os.makedirs(electrical_output_folder, exist_ok=True)
    structural_output_folder = os.path.join(current_path, structural_folder)
    os.makedirs(structural_output_folder, exist_ok=True)
    return data_path, electrical_output_folder, structural_output_folder



# Function: Generate PDF with fitz
def generate_pdf_with_fitz(template, output_folder, file_name, excel_fields_data, static_data, field_name_mapping):
    try:
        doc = fitz.open(template)

        # Update field names if mapping is provided
        if field_name_mapping:
            for page in doc:
                for widget in page.widgets():
                    old_field_name = widget.field_name
                    if old_field_name in field_name_mapping:
                        new_field_name = field_name_mapping[old_field_name]
                        widget.field_name = new_field_name
                        widget.update()

        for page_num in range(len(doc)):
            page = doc[page_num]
            widgets = [w for w in page.widgets()]  # Collect all widgets on the page

            for widget in widgets:
                field_name = widget.field_name
                if field_name:
                    # Directly use the field name for matching
                    if field_name in excel_fields_data:
                        field_value = str(excel_fields_data[field_name])
                    elif field_name in static_data:
                        field_value = str(static_data[field_name])
                    else:
                        continue  # Skip to next widget

                    # Update the field value
                    if widget.field_type in [fitz.PDF_WIDGET_TYPE_TEXT, fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX]:
                        widget.field_value = field_value
                    elif widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:

                        # For button widgets, set to on_state if applicable
                        if field_value.lower() in ['yes', 'true', 'checked']:
                            widget.field_value = True  # Check the box
                        else:
                            widget.field_value = False  # Uncheck the box

                    widget.update()

        output_file_path = os.path.join(output_folder, file_name)
        doc.save(output_file_path)
        doc.close()
    except Exception as e:
        print(f"Error in generate_pdf_with_fitz: {e}")
        traceback.print_exc()


# Function: Process PDFs
def process_pdfs(template, data_path,
                 electrical_output_folder,
                 structural_output_folder,
                 electrical_static_data, 
                 structural_static_data,
                 field_name_mapping):
    try:
        print("Processing... please wait.")
        excel_data = pd.read_excel(data_path)
        for index, row in excel_data.iterrows():
            excel_fields_data = {
                'Job Address': row['Job Address'],
                'City': row['City'],
                'Tax Folio No': row['Tax Folio No'],
                'Job Value': str(row['Job Value']),
                'Legal Description': row['Legal Description'],
                'Property Owner': row['Property Owner'],
                'Phone': row['Phone'],
                'Email': row['Email'],
                'Owners Address': row['Owners Address'],
                'State': row['State'],
                'Zip': row['Zip'],
                'City_2': row['City_2']
            }
            property_owner = row['Property Owner']
            electrical_file_name = f'{property_owner} Filled Electrical Form.pdf'
            structural_file_name = f'{property_owner} Filled Structural Form.pdf'
            generate_pdf_with_fitz(template, electrical_output_folder, electrical_file_name, excel_fields_data, electrical_static_data, field_name_mapping)
            generate_pdf_with_fitz(template, structural_output_folder, structural_file_name, excel_fields_data, structural_static_data, field_name_mapping)
        print("PDFs generated successfully.")
    except FileNotFoundError as e:
        print(f"Error: File not found. Please check the file names and try again. ")
        print(f"File causing error: {e.filename}")
        traceback.print_exc()
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()

# Main Function
def main():
    print("Pandas Version: ", pd.__version__)
    print("PyMuPDF Version: ", fitz.__version__)
    print("Python Version: ", sys.version)
    try:
        data_file_name, template_file_name = get_user_input()
        data_path, electrical_output_folder, structural_output_folder = setup_file_paths(data_file=data_file_name)
        template = 'template.pdf'

        field_name_mapping = {

        }

        electrical_static_data = {
            'TRADE-ELECTRICALCheck Box': 'Yes',
            'Building Use': 'Residential',
            'Dropdown4': 'VB', #Construction Type
            'Occupancy Group': 'Residential', #Not a choice in the dropdown menu
            'Present Use': 'Residential',
            'Proposed Use': 'Residential',
            'Description of Work': 'Solar System Roof Mount and Interconnection',
            'WORK-NEWCheck Box': 'Yes',
            'Contracting Co': 'MES Electric',
            'Phone_2': '(571) 422-0970',
            'Email_2': 'jackson.mcinerney@smartroofinc.com',
            'Company Address': '2083 Guadelupe Dr',
            'City_3': 'Wellington',
            'State_2': 'FL',
            'Zip_2': '33414',
            'Qualifiers Name': 'Mark Spoor',
            'License Number': 'EC13001707',
            'TypePrint Property Owner or Agent Name_2': 'Mark Spoor',
            'Notary Name_2': 'Jackson McInerney'
        }

    # Pre-filled structural contractor data with appropriate PDF form field names gathered from PDF Tester.py
        structural_static_data = {
            'TRADE-BUILDINGCheck Box': 'Yes',
            'Building Use': 'Residential',
            'Dropdown4': 'VB',
            # 'Occupancy Group': 'Residential',
            'Present Use': 'Residential',
            'Proposed Use': 'Residential',
            'Description of Work': 'Solar PV System Roof Mount and Interconnection',
            'WORK-NEWCheck Box': 'Yes',
            'Contracting Co': 'Smart Roof LLC',
            'Phone_2': '(571) 422-0970',
            'Email_2': 'jackson.mcinerney@smartroofinc.com',
            'Company Address': '6413 Congress Ave #225',
            'City_3': 'Boca Raton',
            'State_2': 'FL',
            'Zip_2': '33487',
            'Qualifiers Name': 'Juan David Castro Marino',
            'License Number': 'CGC1528586',
            'TypePrint Property Owner or Agent Name_2': 'Juan David Castro Marino',
            # 'Notary Name_2': 'Jackson McInerney'
            
        }
        process_pdfs(template, data_path, electrical_output_folder, structural_output_folder, electrical_static_data, structural_static_data, field_name_mapping)

    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    main()
    
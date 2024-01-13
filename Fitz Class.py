import fitz  # PyMuPDF
import pandas as pd
import os
import traceback
import sys

class PDFProcessor:
    def __init__(self, default_excel='test excel 1', default_template='template'):

        self.data_file, self.template_file = self.get_user_input(default_excel, default_template)

        self.data_path, self.electrical_output_folder, self.structural_output_folder = self.setup_file_paths()

        self.field_name_mapping = {
            # 'New': 'Old'
            'Job Address': 'Job Address',
            #'City': '', #Can't find on new form
            'Folio': 'Tax Folio No',
            #'Job Value': '',#Can't find on new form
            'Contractor Name': 'Contracting Co',
            'Qualifier Name': 'Qualifiers Name',
            'Address': 'Company Address',
            'City': 'City_3',
            'State': 'State_2',
            'Zip': 'Zip_2',
            'Check Box1': 'WORK-NEWCheck Box',
            'Current use of property 2': 'Present Use',
            'Description of Work 2': 'Description of Work',
            # Legal description unmapped
            'Owner': 'Property Owner',
            'Address_2': "Owners Address",
            'City_2': "City_2",
            'State_2': 'State',
            'Zip_2': 'Zip',
            # Dont use Contracting Phone number, contracting email, or license number
            # I don't have last 4 of owners social, or their phone number
            # or last 4 of qualifier # or contractor number
            'Print': 'TypePrint Property Owner or Agent Name_2',
            'Print_2': 'Notary Name_2',
            'Check Box20': 'TRADE-BUILDINGCheck Box',
            'Check Box21': 'TRADE-ELECTRICALCheck Box'
        }

        # Initialize your static data here
        self.electrical_static_data = {
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
        self.structural_static_data = {
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

    def generate_file_name(self, property_owner, folder_type):
        if folder_type == 'electrical':
            return f'{property_owner} Filled Electrical Form.pdf'
        elif folder_type == 'structural':
            return f'{property_owner} Filled Structural Form.pdf'
        else:
            raise ValueError("Invalid folder type")

    def get_user_input(self, default_excel, default_template):
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

    def setup_file_paths(self):
        current_path = os.getcwd()
        data_path = os.path.join(current_path, self.data_file)
        electrical_output_folder = os.path.join(current_path, 'Filled Electrical PDFs')
        os.makedirs(electrical_output_folder, exist_ok=True)
        structural_output_folder = os.path.join(current_path, 'Filled Structural PDFs')
        os.makedirs(structural_output_folder, exist_ok=True)
        return data_path, electrical_output_folder, structural_output_folder

    def generate_pdf_with_fitz(self, folder_type, excel_fields_data):
        output_folder = self.electrical_output_folder if folder_type == 'electrical' else self.structural_output_folder
        property_owner = excel_fields_data['Property Owner']
        file_name = self.generate_file_name(property_owner, folder_type)

        try:
            doc = fitz.open(self.template_file)

            # Update field names if mapping is provided
            if self.field_name_mapping:
                for page in doc:
                    for widget in page.widgets():
                        old_field_name = widget.field_name
                        if old_field_name in self.field_name_mapping:
                            new_field_name = self.field_name_mapping[old_field_name]
                            widget.field_name = new_field_name
                            widget.update()

            for page_num in range(len(doc)):
                page = doc[page_num]
                widgets = [w for w in page.widgets()]

                for widget in widgets:
                    field_name = widget.field_name
                    if field_name:
                        field_value = excel_fields_data.get(field_name, None)
                        static_data = self.electrical_static_data if folder_type == 'electrical' else self.structural_static_data
                        field_value = field_value or static_data.get(field_name, None)
                        if field_value is None:
                            continue

                        if widget.field_type in [fitz.PDF_WIDGET_TYPE_TEXT, fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX]:
                            widget.field_value = field_value
                        elif widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
                            widget.field_value = True if field_value.lower() in ['yes', 'true', 'checked'] else False

                        widget.update()

            output_file_path = os.path.join(output_folder, file_name)
            doc.save(output_file_path)
            doc.close()
        except Exception as e:
            print(f"Error in generate_pdf_with_fitz: {e}")
            traceback.print_exc()

    def process_pdfs(self):
        try:
            print("Processing... please wait.")
            excel_data = pd.read_excel(self.data_path)
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


                self.generate_pdf_with_fitz('electrical', excel_fields_data)
                self.generate_pdf_with_fitz('structural', excel_fields_data)

            print("PDFs generated successfully.")
        except FileNotFoundError as e:
            print(f"Error: File not found. Please check the file names and try again. ")
            print(f"File causing error: {e.filename}")
            traceback.print_exc()
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            traceback.print_exc()

def main():
    print("Pandas Version: ", pd.__version__)
    print("PyMuPDF Version: ", fitz.__version__)
    print("Python Version: ", sys.version)
    processor = PDFProcessor()
    processor.process_pdfs()

if __name__ == "__main__":
    main()

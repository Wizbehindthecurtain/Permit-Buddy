from pypdf import PdfReader, PdfWriter, generic
import pandas as pd
import os
import pkg_resources
from pypdf.generic import BooleanObject

def get_user_input(default_excel='test excel 1.xlsx', default_template='template.pdf'):
    print("Please enter the required file names or type 'default' to use the preset names.")

    data_file = input(f"Enter the name of the Excel file (default: '{default_excel}'): ").strip()
    if data_file.lower() == 'default' or not data_file:
        data_file = default_excel

    template_file = input(f"Enter the name of the PDF template file (default: '{default_template}'): ").strip()
    if template_file.lower() == 'default' or not template_file:
        template_file = default_template

    # Add additional file name inputs and checks here as needed

    return data_file, template_file

def setup_file_paths(data_file='test excel 1.xlsx', electrical_folder='Filled Electrical PDFs', structural_folder='Filled Structural PDFs'):
    current_path = os.getcwd()
    data_path = os.path.join(current_path, data_file)

    electrical_output_folder = os.path.join(current_path, electrical_folder)
    os.makedirs(electrical_output_folder, exist_ok=True)

    structural_output_folder = os.path.join(current_path, structural_folder)
    os.makedirs(structural_output_folder, exist_ok=True)

    return data_path, electrical_output_folder, structural_output_folder

def generate_pdf(template, output_folder, file_name, pdf_fields_data, static_data):
    output_file_path = os.path.join(output_folder, file_name)
    with open(template, 'rb') as file:
        reader = PdfReader(file)
        writer = PdfWriter()

        if '/AcroForm' in reader.trailer["/Root"]:
            reader.trailer["/Root"]["/AcroForm"].update({
                generic.NameObject("/NeedAppearances"): BooleanObject(True)
            })

        for page in reader.pages:
            if '/Annots' in page:
                for annot in page['/Annots']:
                    obj = annot.get_object()
                    field_name = obj.get("/T")
                    field_type = obj.get("/FT")

                    if field_name and field_type:
                        field_name = field_name.strip('()')

                        # Check for text fields in both pdf_fields_data and static_data
                        if field_type == '/Tx':
                            if field_name in pdf_fields_data:  # First preference to dynamic data
                                obj.update({
                                    generic.NameObject("/V"): generic.create_string_object(pdf_fields_data[field_name])
                                })
                            elif field_name in static_data:  # Then static data
                                obj.update({
                                    generic.NameObject("/V"): generic.create_string_object(static_data[field_name])
                                })

                        # Check for dropdown fields in static_data
                        elif field_type == '/Ch' and field_name in static_data:
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object(static_data[field_name])
                            })

                        # Check for checkbox fields in static_data
                        elif field_type == '/Btn' and field_name in static_data:
                            check_value = '/Yes' if static_data[field_name] == 'Yes' else '/Off'
                            obj.update({
                                generic.NameObject("/V"): generic.NameObject(check_value),
                                generic.NameObject("/AS"): generic.NameObject(check_value)
                            })

            writer.add_page(page)

        with open(output_file_path, 'wb') as output_file:
            writer.write(output_file)



def process_and_generate_pdfs(template, data_path, electrical_output_folder, structural_output_folder, electrical_static_data, structural_static_data):
    excel_data = pd.read_excel(data_path)

    for index, row in excel_data.iterrows():
        pdf_fields_data = {
            # ... (existing code to extract data from the row)
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

        generate_pdf(template, electrical_output_folder, electrical_file_name, pdf_fields_data, electrical_static_data)
        generate_pdf(template, structural_output_folder, structural_file_name, pdf_fields_data, structural_static_data)

# Main execution
print("Pandas Version:", pd.__version__)
print("pypdf Version:", pkg_resources.get_distribution('pypdf').version)

data_file_name, template_file_name = get_user_input()
data_path, electrical_output_folder, structural_output_folder = setup_file_paths(data_file=data_file_name)

template = 'template.pdf'
data_path, electrical_output_folder, structural_output_folder = setup_file_paths()

electrical_static_data = {
    'TRADE-ELECTRICALCheck Box': 'Yes',
    'Building Use': 'Residential',
    'Dropdown4': 'VB',
    'Occupancy Group': 'Residential',
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

process_and_generate_pdfs(template, data_path, electrical_output_folder, structural_output_folder, electrical_static_data, structural_static_data)

print("PDFs generated successfully.")

from pypdf import PdfReader, PdfWriter, generic
import pandas as pd
import os

# Example data
template = 'Filled Electrical Form 24.pdf'
current_path = os.getcwd()

data_file_name = 'test excel 1.xlsx'
data_path = os.path.join(current_path, data_file_name)

# Read Excel data into a DataFrame
excel_data = pd.read_excel(data_path)

# Create a new folder for output PDFs
electrical_output_folder = os.path.join(current_path, 'Filled Electrical PDFs')
os.makedirs(electrical_output_folder, exist_ok=True)

structural_output_folder = os.path.join(current_path, 'Filled Structural PDFs')
os.makedirs(structural_output_folder, exist_ok=True)

electrical_static_data = {
    'TRADE-ELECTRICALCheck Box': 'Yes',
    'Building Use': 'Residential',
    'Construction Type': 'VB',
    'Occupancy Group': 'Residential',
    'Present Use': 'Residential',
    'Proposed Use': 'Residential',
    'Description of Work': 'Solar System Roof Mount and Interconnection',
    'WORK-NEWCheck Box': '/On',
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
    'Construction Type': 'VB',
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

# Iterate over rows in the DataFrame
for index, row in excel_data.iterrows():
    # Example data from the current row
    pdf_fields_data = {
        'Job Address': row['Job Address'],
        'City': row['City'],
        'Tax Folio No': row['Tax Folio No'],
        'Job Value': str(row['Job Value']),  # Assuming 'Job Value' is numeric; convert to string
        'Legal Description': row['Legal Description'],
        'Property Owner': row['Property Owner'],
        'Phone': row['Phone'],
        'Email': row['Email'],
        'Owners Address': row['Owners Address'],
        'State': row['State'],
        'Zip': row['Zip']
    }

    # Process PDF for the current row
    electrical_file_name = f'Filled Electrical Form {index + 1}.pdf'  # Output PDF file name
    output_file_path = os.path.join(electrical_output_folder, electrical_file_name)

    with open(template, 'rb') as file:
        reader = PdfReader(file)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]

            # reader.get_fields()

            if '/Annots' in page:
                for field in page['/Annots']:
                    obj = field.get_object()
                    if '/T' in obj and '/FT' in obj and obj['/FT'] == '/Tx':
                        field_name = obj['/T']
                        if field_name in pdf_fields_data:
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object(pdf_fields_data[field_name])
                            })
                        if field_name in electrical_static_data:
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object(electrical_static_data[field_name])
                            })
                    elif '/T' in obj and '/FT' in obj and obj['/FT'] == '/Btn':
                        field_name = obj['/T']
                        if field_name == 'TRADE-ELECTRICALCheck Box':
                            # Update the checkbox value
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object('Yes'),
                                generic.NameObject("/AS"): generic.create_string_object('Yes')
                            })  
                            

            writer.add_page(page)

        # Update form field values
        writer.update_page_form_field_values(page, pdf_fields_data)
        writer.update_page_form_field_values(page, electrical_static_data)

        # Save the filled PDF to a new file
        with open(output_file_path, 'wb') as output_file:
            writer.write(output_file)
            output_file.close()

    structural_file_name = f'Filled Structural Form {index + 1}.pdf'  # Output PDF file name
    output_file_path = os.path.join(structural_output_folder, structural_file_name)

    with open(template, 'rb') as file:
        reader = PdfReader(file)
        writer = PdfWriter()

        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]

            # reader.get_fields()

            if '/Annots' in page:
                for field in page['/Annots']:
                    obj = field.get_object()
                    if '/T' in obj and '/FT' in obj and obj['/FT'] == '/Tx':
                        field_name = obj['/T']
                        if field_name in pdf_fields_data:
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object(pdf_fields_data[field_name])
                            })
                        if field_name in structural_static_data:
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object(structural_static_data[field_name])
                            })
                    elif '/T' in obj and '/FT' in obj and obj['/FT'] == '/Btn':
                        field_name = obj['/T']
                        if field_name == 'TRADE-BUILDINGCheck Box':
                            # Update the checkbox value
                            obj.update({
                                generic.NameObject("/V"): generic.create_string_object('Yes'),
                                generic.NameObject("/AS"): generic.create_string_object('Yes')
                            })  
                            

            writer.add_page(page)

        # Update form field values
        writer.update_page_form_field_values(page, pdf_fields_data)
        writer.update_page_form_field_values(page, electrical_static_data)

        # Save the filled PDF to a new file
        with open(output_file_path, 'wb') as output_file:
            writer.write(output_file)
            output_file.close()

print("PDFs generated successfully.")

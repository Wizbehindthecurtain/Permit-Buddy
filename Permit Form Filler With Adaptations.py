from pypdf import PdfReader, PdfWriter, generic
import pandas as pd
import os
import pkg_resources
from pypdf.generic import BooleanObject
import traceback


##############################################################################################################################
###########################################################################################################################
### Author:     Jonathan Mazurkiewicz
###
### Date:       1/1/2024
###
### Purpose:    Gather input from excel, fill out, and create new PDFs for structural and electrical contractors.
###
### Functions:  get_user_input(excel_file_name, template_file_name)
###             setup_file_paths(excel_file_name, electrical_folder_name, structural_folder_name)
###             generate_pdf(template_name, output_folder, file_name, excel_data, static_data)
###             process_pdfs(template, data_path, electrical_output_folder, structural_output_folder,
###                 electrical_static_data, structural_static_data)
###
### Notes:      (1) Utilizes pandas to read from excel. The field names MUST be named exactly the way they are in the template.
###             Naming the fields wrong in row 1 of excel will result in the program not working properly.
###
###             (2) Utilizes pypdf Version 3.8.1. Ensure that you are installing the proper version as newer versions cause the
###             code to be unstable. Will check for updates in the future. The command line script for installing is:
###             pip install pypdf==3.8.1
###
###             (3) When filling out the form fields in a PDF, you need the assigned 'name' field. These entries are made when
###             the PDF is created. This program will only work with the current template since other PDFs will have different
###             field names. There is a supplementary file named 'PDF Tester' that will accept the pdf file name as an input,
###             then return the form field names. These values can then be used for excel sheet column titles and for static
###             info to allow you to play with the code yourself, if you need to change anything.
###
###             (4) Uses Python Version 3.10.11 as of 1/1/2024
###
###             (5) This code accepts user input for file names. The template pdf file and the excel file MUST be inside of
###             the project folder. Enter the name of the file as it appears in the folder. Otherwise, enter nothing or
###             'default' to use the preset file name.
###
###             (6) Exploring functionality on macOS. Here are the steps for installing on macOS currently (from Terminal):
###                 ** You can find the terminal in Applications -> Utilities -> Terminal
###                 1. Install Homebrew
###                     /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"                 
###                 2. Install Python
###                     brew install python@3.10
###                 3. Check the version to make sure it was installed correctly
###                     python3 --version
###                 4. Install Git for cloning this file
###                     brew install git
###                 5. Install VSCode: Go to the website, install, drag to applications, then launch.
###                 6. Open the terminal again, and make a folder for your Python project.
###                     mkdir ~/Documents/MyPythonProject
###                     cd ~/Documents/MyPythonProject
###                 7. Create virtual environment
###                     python3 -m venv venv
###                 8. Activate virtual environment
###                     source venv/bin/activate
###                 9. While virtual environment is active, install dependencies.
###                     pip install pypdf==3.8.1
###                     pip install pandas
###                 10. Clone from github
###                     Make sure you are in the project folder, then, in the Terminal:
###                         git clone https://github.com/jonmaz4410/Jackson-PDF1.git
###                     IF you need to update the code to the most recent version, enter:
###                         git pull
###                 11. (Optional) Set up your name and email if its the first time using Git
###                     git config --global user.name "Your Name"
###                     git config --global user.email "youremail@example.com"
###                 12. Install necessary extensions in VSCode (Python, Pylance)
###                 13. Run the code! Hopefully it works. If not, send me an email at jonmaz4410@gmail.com and I'll try to help.
###
##############################################################################################################################
##############################################################################################################################

##############################################################################################################################
### FUNCTION DEFINITIONS
##############################################################################################################################


##############################################################################################################################
### Name:       get_user_input()
### Arguments:  default_excel, default_template
### Purpose:    Gather user input for file names for template PDF and excel sheet, in case they change
### Notes:      Provides default values for file names. Only requires the name and not the extension.
   
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

##############################################################################################################################
### Name:       setup_file_paths()
### Arguments:  data_file, electrical_folder, structural_folder,
### Purpose:    Create the names of the output file folders. Currently there is no user input functionality
### Notes:      Provides default values for file names. Only requires the name and not the extension.

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

##############################################################################################################################
### Name:       generate_pdf()
### Arguments:  template, output_folder, file_name, excel_fields_data, static_data
### Purpose:    Read template, update AcroForm for visibility, update text, dropdown and button fields
###             with data from excel and pre-filled static values. Then, write a new PDF. This function is called by
###             the next function, process_pdfs().
### Notes:      None

def generate_pdf(template, output_folder, file_name, excel_fields_data, static_data):
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

                        # Check for text fields in both excel_fields_data and static_data
                        if field_type == '/Tx':
                            if field_name in excel_fields_data:  # First preference to dynamic data
                                obj.update({
                                    generic.NameObject("/V"): generic.create_string_object(excel_fields_data[field_name])
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

##############################################################################################################################
### Name:       process_pdf()
### Arguments:  template, data_path, electrical_output_folder,structural_output_folder, electrical_static_data, structural_static_data
### Purpose:    Read from excel file, row by row based on names of column headers. Get the property owner name to have for saving
###             the file. Then, call the generate_pdf() function.
### Notes:      None

def process_pdfs(template, data_path,
                 electrical_output_folder,
                 structural_output_folder,
                 electrical_static_data, 
                 structural_static_data):
    try:
        print("Processing... please wait.")
        with open(template, 'rb') as file:
            reader = PdfReader(file)
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

                generate_pdf(template, electrical_output_folder, electrical_file_name, excel_fields_data, electrical_static_data)
                generate_pdf(template, structural_output_folder, structural_file_name, excel_fields_data, structural_static_data)
    except FileNotFoundError:
        print(f"Error: File not found. Please check the file names and try again.")
        return
    
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        traceback.print_exc()
        return
    
    print("PDFs generated successfully.")
    
##############################################################################################################################

##############################################################################################################################

### Main Function ###

# Print versions of dependencies to troubleshoot in case the program doesn't work.    
print("Pandas Version:", pd.__version__)
print("pypdf Version:", pkg_resources.get_distribution('pypdf').version)


try:
    # Prompt user for input of file names, if necessary.
    data_file_name, template_file_name = get_user_input()

    # Set up the file folders for the output files for structural and electrical filled PDFs 
    data_path, electrical_output_folder, structural_output_folder = setup_file_paths(data_file=data_file_name)

    template = 'template.pdf'

    # Pre-filled electrical contractor data with appropriate PDF form field names gathered from PDF Tester.py
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

    # Function call to fill out and save PDF based on each row of data from Excel sheet.
    process_pdfs(template, data_path, electrical_output_folder, structural_output_folder, electrical_static_data, structural_static_data)

except Exception as e:
    print(f"An unexpected error occurred: {e}")
    traceback.print_exc()

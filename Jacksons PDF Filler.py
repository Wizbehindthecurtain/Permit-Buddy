import fitz  # PyMuPDF
import pandas as pd
import os
import traceback
import sys

##############################################################################################################################
###########################################################################################################################
### Author:     Jonathan Mazurkiewicz
###
### Date:       1/27/2024
###
### Purpose:    Gather input from excel, fill out, and create new PDFs for structural and electrical contractors.
###
### Notes:      (1) Utilizes pandas to read from excel. The field names MUST be named exactly the way they are in the template.
###             Naming the fields wrong in row 1 of excel will result in the program not working properly.
###
###             (2) Utilizes fitz (PyMuPDF) to fill in the templates. This is a comprehensive library for PDF modifications.
###
###             (3) When filling out the form fields in a PDF, you need the assigned 'name' field. The file is made to work with
###             other templates by adding the correct fields into "field_mappings" and "fields_to_skip_dict". For each new template,
###             please use the PDF Tester file to find the correct field names until a GUI can be set up.
###
###             (4) Uses Python Version 3.10.11 as of 1/1/2024. It should work fine with other versions of Python, but if in
###             doubt, revert back to this version.
###
###             (5) This code accepts user input for file names. The template pdf file and the excel file MUST be inside of
###             the project folder. Enter the name of the file as it appears in the folder. Otherwise, enter nothing or
###             'default' to use the preset file name.
###
###             (6) Exploring functionality on macOS. Here are the steps for installing on macOS currently (from Terminal):
###                 ** You can find the terminal in Applications -> Utilities -> Terminal
###                 1.  Install Homebrew
###                         /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
###       
###                 2.  Install Python
###                         brew install python@3.10
###                 *   If you find that 'brew' commands dont work, that means you need to edit the PATH
###                 *   Check if you are using Bash or Zsh for your SHELL. Usually, if you see a '%' symbol at the end of
###                 *   the terminal prompt, that means Zsh while '$' means Bash. If you are unsure, you can run this line:
###                 *   echo $SHELL
###                 *   macOS switched to Zsh recently. Depending on the SHELL, write the following into Terminal:
###                 *   For Bash:
###                 *       echo 'export PATH="/usr/local/bin:$PATH"' >> ~/.bash_profile
###                         source ~/.bash_profile
###                 *   For Zsh:
###                 *       echo 'export PATH="/usr/local/bin:$PATH"' >> ~/.zshrc
###                 *       source ~/.zshrc
###
###                 3.  Check the version to make sure it was installed correctly
###                         python3 --version
###                 *   If the version does not read python3.10.xx, you will need to edit the PATH as mentioned above.
###                 *   Feel free to try to run the code with your existing version of Python, but if it does not work,
###                 *   come back here for troubleshooting.
###
###                 4.  Install Git for cloning this file
###                         brew install git
###
###                 5.  Install VSCode: Go to the website, install, drag to applications, then launch.
###
###                 6.  Open the terminal again, and make a folder for your Python project.
###                         mkdir ~/Documents/MyPythonProject
###                         cd ~/Documents/MyPythonProject      
###
###                 7.  Create virtual environment
###                         python3 -m venv venv
###
###                 8.  Activate virtual environment
###                         source venv/bin/activate
###                 *   Now, the terminal should say (venv)
###
###                 9.  While virtual environment is active, install dependencies.
###                         sudo pip3 install PyMuPDF pandas openpyxl
###
###                 10. Clone files from github
###                 *   Make sure you are in the project folder in Terminal. If you closed the terminal, then reopen and type:
###                 *       cd ~/Documents/MyPythonProject
###                 *   Then, paste this line:
###                 *       git clone https://github.com/jonmaz4410/Jackson-PDF1.git
###                 *   IF you need to update the code to the most recent version in the future, enter:
###                 *       git pull
###
###                 11. Next, copy all of the files that were cloned and paste them in the PROJECT FOLDER. THE CODE WILL NOT WORK
###                 otherwise! To be more clear, the cloned files will be in the folder "Jackson-PDF1". Copy the contents inside
###                 of that folder into the place where the FOLDER exists (your project folder). Then, optionally, you can delete
###                 the Jackson-PDF1 folder.
###
###                 12. (Optional) Set up your name and email if its the first time using Git
###                     git config --global user.name "Your Name"
###                     git config --global user.email "youremail@example.com"
###
###                 13. Install necessary extensions in VSCode, if not already installed (Python, Pylance)
###
###                 14. Open the Project Folder in VSCode and set up the virtual environment
###                 *   Press 'Cmd + Shift + P'
###                 *   Type Python: Select Interpreter
###                 *   Choose the one that points to the virtual environment folder
###                 *       Should look like '.venv/bin/python'
###
###                 13. Run the code! Hopefully it works. If not, send me an email at jonmaz4410@gmail.com and I'll try to help.
###
##############################################################################################################################
#############################################################################################################################


class PDFProcessor:
    def __init__(self, default_excel='test excel 1', default_template='template'):

        self.fields_to_skip_dict = {
            'miami': ['City_3', 'Phone_2'],
            'palm beach': ['a', 'b']
            # etc etc
        }

        self.data_file, self.template_file, self.fields_to_skip_current = self.get_user_input(default_excel, default_template)

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

        
        jurisdiction = input("Enter the name of your jurisdiction (e.g., 'miami'): ").strip().lower()
        
        # Retrieve the correct list of fields to skip based on the jurisdiction
        fields_to_skip_current = self.fields_to_skip_dict.get(jurisdiction, [])
        
        return data_file, template_file, fields_to_skip_current

    def setup_file_paths(self):
        current_path = os.path.dirname(os.path.abspath(__file__))
        os.chdir(current_path)
        data_path = os.path.join(current_path, self.data_file)
        electrical_output_folder = os.path.join(current_path, 'Filled Electrical PDFs')
        os.makedirs(electrical_output_folder, exist_ok=True)
        structural_output_folder = os.path.join(current_path, 'Filled Structural PDFs')
        os.makedirs(structural_output_folder, exist_ok=True)
        print(current_path)
        return data_path, electrical_output_folder, structural_output_folder
    

    def generate_pdf_with_fitz(self, folder_type, excel_fields_data):
        output_folder = self.electrical_output_folder if folder_type == 'electrical' else self.structural_output_folder
        property_owner = excel_fields_data['Property Owner']
        file_name = self.generate_file_name(property_owner, folder_type)

        try:
            doc = fitz.open(self.template_file)
            for page in doc:
                for widget in page.widgets():
                    original_field_name = widget.field_name

                    # Skip fields based on the current jurisdiction's list
                    if original_field_name in self.fields_to_skip_current:
                        continue
                    # Remap field names using the extended mapping
                    new_field_name = self.field_name_mapping.get(original_field_name, original_field_name)

                    # Determine the appropriate data source based on folder_type
                    static_data = self.electrical_static_data if folder_type == 'electrical' else self.structural_static_data

                    # Check if the field should be filled
                    if new_field_name in static_data or new_field_name in excel_fields_data:
                        field_value = excel_fields_data.get(new_field_name, static_data.get(new_field_name))
                        # Fill the field with the determined value
                        self.fill_widget(widget, field_value)

            # Save the updated PDF
            output_file_path = os.path.join(output_folder, file_name)
            doc.save(output_file_path)
            doc.close()
        except Exception as e:
            print(f"Error in generate_pdf_with_fitz: {e}")
            traceback.print_exc()

    def fill_widget(self, widget, field_value):
        if field_value is None:
            return  # Skip if no value to fill

        if widget.field_type in [fitz.PDF_WIDGET_TYPE_TEXT, fitz.PDF_WIDGET_TYPE_COMBOBOX, fitz.PDF_WIDGET_TYPE_LISTBOX]:
            widget.field_value = field_value
        elif widget.field_type == fitz.PDF_WIDGET_TYPE_CHECKBOX:
            widget.field_value = field_value.lower() in ['yes', 'true', 'checked']
        widget.update()


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

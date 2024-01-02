from pypdf import PdfReader

#Change name of the PDF you want to test here.
template = 'template.pdf'
with open(template, 'rb') as file:
    reader = PdfReader(file)

    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]

        if '/Annots' in page:
            for field in page['/Annots']:
                obj = field.get_object()
                
                # Check if the object is a form field
                if '/T' in obj and '/FT' in obj:
                    field_name = obj['/T']
                    field_type = obj['/FT']
                    
                    print(f"Field Name: {field_name}")
                    print(f"Field Type: {field_type}")
                    
                    # Additional information depending on the type of field
                    if field_type == '/Tx':  # Text field
                        print(f"Field Value (/V): {obj.get('/V', None)}")
                        print(f"Field Default Value (/DV): {obj.get('/DV', None)}")
                        print(f"Field Flags (/Ff): {obj.get('/Ff', None)}")
                        
                        # Look for visibility-related information in the /DA string
                        da_string = obj.get('/DA', None)
                        if da_string:
                            print(f"Field Default Appearance (/DA): {da_string}")
                            # Extract additional information from the /DA string if needed
                        
                    elif field_type == '/Btn':  # Button field
                        print(f"Field Value (/V): {obj.get('/V', None)}")
                        print(f"Field Flags (/Ff): {obj.get('/Ff', None)}")
                        print(f"Field Default Appearance (/DA): {obj.get('/DA', None)}")
                        print(f"Field Appearance State (/AS): {obj.get('/AS', None)}")
                    
                    print("-" * 30)  # Separator between fields

# Note: This is a basic example, and you may need to adapt it based on your specific PDF structure and the details you want to extract.

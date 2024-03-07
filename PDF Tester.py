import fitz
import traceback


def get_pdf_form_field_names(pdf_path):
    field_types = {
        fitz.PDF_WIDGET_TYPE_TEXT: "Text Field",
        fitz.PDF_WIDGET_TYPE_CHECKBOX: "Checkbox",
        fitz.PDF_WIDGET_TYPE_COMBOBOX: "Combobox",
        fitz.PDF_WIDGET_TYPE_LISTBOX: "Listbox",
        fitz.PDF_WIDGET_TYPE_RADIOBUTTON: "Radiobutton",
        fitz.PDF_WIDGET_TYPE_SIGNATURE: "Signature",
        fitz.PDF_WIDGET_TYPE_BUTTON: "Button"
    }

    fields_info = {}
    try:
        doc = fitz.open(pdf_path)
        for page in doc:
            for widget in page.widgets():
                field_name = widget.field_name
                field_type = widget.field_type
                field_type_name = field_types.get(field_type, "Unknown")

                if field_name:
                    if field_type_name not in fields_info:
                        fields_info[field_type_name] = []
                    if field_name not in fields_info[field_type_name]:
                        fields_info[field_type_name].append(field_name)
        doc.close()
    except Exception as e:
        print(f"Error while reading PDF form fields: {e}")
        traceback.print_exc()

    return fields_info

# Example usage
pdf_path = 'miami template.pdf'  # Replace with the path to your new PDF
fields_info = get_pdf_form_field_names(pdf_path)

# Printing the fields in a readable format
for field_type, names in fields_info.items():
    print(f"{field_type}:")
    for name in names:
        print(f"  - {name}")

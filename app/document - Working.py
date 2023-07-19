import os
from docx import Document
import openpyxl
import sys
from gptkey import generate_prompt


def generate_email_documents(excel_sheet_name):
    # Loading the Excel workbook
    workbook = openpyxl.load_workbook(excel_sheet_name)

    # Selecting a sheet within the Excel workbook
    Leads = workbook['Leads']
    Variations = workbook['Variations']

    # Create the output folder if it doesn't exist
    output_folder = "OutputDocs"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Iterate over rows in the Leads sheet, starting from row 2
    for row_index, row in enumerate(Leads.iter_rows(min_row=2, values_only=True), start=2):
        # Create a new Word document
        doc = Document()

        # Read values from specific cells and store them in variables
        company_name = Leads[f'B{row_index}'].value
        TG = Leads[f'C{row_index}'].value
        E_Variation = Leads[f'D{row_index}'].value
        F_Name = Leads[f'E{row_index}'].value
        L_Name = Leads[f'F{row_index}'].value
        Email_1 = Leads[f'G{row_index}'].value
        Email_2 = Leads[f'H{row_index}'].value
        website = Leads[f'I{row_index}'].value

        # Assigning the Variation
        tosend = ""
        V1 = Variations['A2'].value
        V2 = Variations['B2'].value
        V3 = Variations['C2'].value
        # Ensuring the correct email is sent
        if E_Variation == 1:
            tosend = V1
        elif E_Variation == 2:
            tosend = V2
        else:
            tosend = V3

        # Redirect standard output to Word document
        output = sys.stdout
        sys.stdout = open("output.txt", "w")
        prompt = f"You are an email marketing expert. You already have the email templates, but are tasked with visiting this website: {website} and coming out with a 40-50 complete word compliment about their services and accomplishments to include in the email."
        # Testing to see - PUTS THE CONTENT TOGETHER
        print(f"Hey {F_Name},")
        print()
        print(f"I was just going through {company_name}'s website and I just had to reach out to you.")
        print()
        print(generate_prompt(prompt))
        print()
        print(tosend)
        print()
        print("Thanks & Regards,")

        # Restore standard output
        sys.stdout = output

        # Read the content from the temporary file
        with open("output.txt", "r") as file:
            content = file.read()

        # Add the content to the Word document
        doc.add_paragraph(content)

        # Generate a unique filename based on the company name
        filename = f"{company_name}_output.docx"

        # Saving the document with company name as file name in the output folder
        doc.save(os.path.join(output_folder, filename))

    print("The documents have been generated")


# Call the function with the Excel sheet name as an argument
generate_email_documents('PV_Lead Management Sheet_Python.xlsx')

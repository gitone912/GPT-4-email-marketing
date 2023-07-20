from django.http import HttpResponse
from django.shortcuts import render
from django.core.mail import send_mail
from django.conf import settings
from rest_framework.decorators import api_view
import os
from docx import Document
import pandas as pd
from .gptkey import generate_prompt

def generate_and_send_email_documents(excel_sheet_name):
    # Loading the Excel workbook using pandas
    df_leads = pd.read_excel(excel_sheet_name, sheet_name='Leads')
    df_variations = pd.read_excel(excel_sheet_name, sheet_name='Variations')

    # Create the output folder if it doesn't exist
    output_folder = "OutputDocs"
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Iterate over rows in the Leads sheet
    for _, row in df_leads.iterrows():
        # Read values from specific cells and store them in variables
        company_name = row['Company Name']
        TG = row['TG1/TG2']
        E_Variation = row['Email Variation']
        Lead1 = row['Lead_1']
        Lead2 = row['Lead_2']
        Email_1 = row['Email']
        Email_2 = row['Email 2']
        website = row['Website']

        # Assigning the Variation
        tosend = ""
        V1 = df_variations.at[0, 'V1']
        V2 = df_variations.at[0, 'V2']
        V3 = df_variations.at[0, 'V3']
        # Ensuring the correct email is sent
        if E_Variation == 1:
            tosend = V1
        elif E_Variation == 2:
            tosend = V2
        else:
            tosend = V3

        # Generate the email content
        prompt = f"Hey ChatGPT, as an email marketing expert, you possess exceptional insight into crafting engaging messages. Today, your task is to visit {company_name}'s website: {website} and craft a compliment that specifically references the exact highlights the company's outstanding services/accomplishments and the importance of said services in today's market. WORDS TO AVOID: 'their',' they', 'congratulations ', 'thank you'. APPROACHES TO AVOID: Talking about customer service. WORDS TO USE: 'your' TRY TO: make the paragraph fluent, reflecting genuine admiration for the company's offerings. MUST BE: be specific about their service offerings. MUST BE: within 30-35 words. DO NOT: miss any of these instructions."
        greeting = f"Hi {Lead1}" if pd.isna(Lead2) else f"Hi {Lead1} & {Lead2}"

        content = f"{greeting}\n\n"
        content += f"I was just going through {company_name}'s website and I just had to reach out to you.\n\n"
        content += f"{generate_prompt(prompt)}\n\n"
        content += f"{tosend}\n\n"
        content += "Thanks & Regards,"

        # Create the Word document and add the content
        doc = Document()
        doc.add_paragraph(content)

        # Generate a unique filename based on the company name
        filename = f"{company_name}_output.docx"

        # Saving the document with the company name as the file name in the output folder
        doc.save(os.path.join(output_folder, filename))

        # Send email to the respective email IDs
        send_mail(
            f'{company_name} Email Template',
            content,
            settings.EMAIL_HOST_USER,
            [Email_1, Email_2],  # You can add more recipients or customize this list based on your needs
            fail_silently=True,
        )

    print("The documents have been generated and emails have been sent.")

def home(request):
    if request.method == 'POST':
        # Handle the file upload
        if 'excel_sheet' in request.FILES:
            excel_sheet = request.FILES['excel_sheet']
            # Save the uploaded file temporarily
            with open('temp.xlsx', 'wb') as destination:
                for chunk in excel_sheet.chunks():
                    destination.write(chunk)

            # Call the function with the uploaded Excel sheet
            generate_and_send_email_documents('temp.xlsx')

            # Clean up the temporary file
            os.remove('temp.xlsx')

            return HttpResponse("The documents have been generated and emails have been sent.")
        else:
            return HttpResponse("No file was uploaded.")

    return render(request, 'index.html')

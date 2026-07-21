import pdfplumber

from cleaner import clean_text
from cleaner import clean_date
from cleaner import get_month
from utils import extract


def read_pdf(pdf_path):

    text = ""

    with pdfplumber.open(pdf_path) as pdf:

        for page in pdf.pages:

            page_text = page.extract_text()

            if page_text:
                text += page_text + "\n"

    data = {

        "INQUIRY_NO":
            extract(r"RFQ NO\s+([A-Za-z0-9\-\/]+)", text),

        "CLIENT":
            extract(r"CUSTOMER NAME\s+(.*?)\s+SENT DATE", text),

        "LOCATION":
            extract(r"LOCATION\s+(.*?)\s+ENQUIRY DATE", text),

        "CONSULTANT":
            extract(r"CONSULTANT\s+(.*?)\s+REV NO", text),

        "DESIGN_RELEASE_DATE":
            extract(r"SENT DATE\s+([0-9A-Za-z\-]+)", text),

        "QUERIES_RECEIVED_DATE":
            extract(r"ENQUIRY DATE\s+([0-9A-Za-z\-]+)", text),

        "ABR_DESIGNER":
            extract(r"DESIGN BY\s+([A-Za-z]+)", text),

        "CHECKED_BY":
            extract(r"CHECKED BY\s+([A-Za-z]+)", text),

        "ABR_SALES":
            extract(r"SALES ENGG\.\s*([^\r\n]+)", text),

        "BUILDING":
            extract(r"BLDG\s+(.*?)\s+DESIGN BY", text),

        "REVISION":
            extract(r"REV NO\.\s*([0-9]+)", text),

        "COMPLEXITY":
            extract(r"Complexity Index\s*:\s*([A-Z]+)", text),

        "DESIGN_WEIGHT_MT":
            extract(r"TOTAL\s+([0-9]+\.[0-9]+)", text)
            

    }
    for key in data:
        if "DATE" in key:
            data[key] = clean_date(data[key])
        else:
            data[key] = clean_text(data[key])

    # Extract and add the specific month 
    data["MONTHS"] = get_month(data["QUERIES_RECEIVED_DATE"])        

    return data
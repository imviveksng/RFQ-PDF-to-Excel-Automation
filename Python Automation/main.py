from config import NEW_PDF_FOLDER
from pdf_reader import read_pdf
from excel_writer import write_to_excel
from utils import move_pdf

def main():
    # Find all PDFs in the New PDFs folder
    pdf_files = list(NEW_PDF_FOLDER.glob("*.pdf"))
    
    if not pdf_files:
        print("No PDF Found.")
        return

    # Process each PDF one by one
    for pdf in pdf_files:
        print(f"\nProcessing : {pdf.name}")
        
        # 1. Read and clean the data
        data = read_pdf(pdf)
        
        # 2. Write the data to Excel
        write_to_excel(data)
        
        # 3. Rename and move the PDF to the Processed folder
        move_pdf(pdf, data["INQUIRY_NO"])
        
    print("\nFinished.")

if __name__ == "__main__":
    main()
import os
import comtypes.client

# Set the folder path where your DOCX files are stored
folder_path = r"C:\path\to\your\docx\files"  # Change this to your folder

# Initialize Word application
word = comtypes.client.CreateObject("Word.Application")
word.Visible = False  # Run in the background

# Convert all DOCX files to PDF
for file in os.listdir(folder_path):
    if file.endswith(".docx"):
        docx_path = os.path.join(folder_path, file)
        pdf_path = os.path.join(folder_path, file.replace(".docx", ".pdf"))

        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = PDF format
        doc.Close()

# Quit Word application
word.Quit()

print("All DOCX files have been converted to PDF successfully!")

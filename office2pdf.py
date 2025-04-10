import os
import comtypes.client
import win32com.client

# Function to convert PPTX to PDF
def pptx_to_pdf(input_path, output_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
    presentation.SaveAs(output_path, 32)  # 32 = PDF format
    presentation.Close()
    powerpoint.Quit()

    print(f"✔ Saved: {output_path}")

# Function to convert DOCX to PDF
def docx_to_pdf(input_path, output_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF format
    doc.Close()
    word.Quit()

    print(f"✔ Saved: {output_path}")

# Function to convert XLSX to PDF
def xlsx_to_pdf(input_path, output_path):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(input_path)
    wb.ExportAsFixedFormat(0, output_path)  # 0 = PDF format
    wb.Close()
    excel.Quit()

    print(f"✔ Saved: {output_path}")

# Function to convert all supported files in the folder
def convert_all_files_in_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            full_path = os.path.join(root, file)
            output_path = os.path.splitext(full_path)[0] + ".pdf"

            if file.lower().endswith(".pptx"):
                print(f"Converting: {full_path}")
                pptx_to_pdf(full_path, output_path)
            elif file.lower().endswith(".docx"):
                print(f"Converting: {full_path}")
                docx_to_pdf(full_path, output_path)
            elif file.lower().endswith(".xlsx"):
                print(f"Converting: {full_path}")
                xlsx_to_pdf(full_path, output_path)
            else:
                print(f"Skipping: {full_path} (Unsupported file type)")

# Main program function
if __name__ == "__main__":
    folder = input("Enter the path to the folder with files: ")

    if os.path.isdir(folder):
        convert_all_files_in_folder(folder)
    else:
        print("❌ The provided path does not exist.")

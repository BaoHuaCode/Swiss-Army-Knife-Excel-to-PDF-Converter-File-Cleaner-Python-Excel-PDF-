import ezsheets, send2trash
from pathlib import Path



def The_swiss_army_knife(folder):
    """Convert excel to pdf and delete empty pdf files."""
    folder_path = Path(folder)
    if not folder_path.exists():
        print(f"{folder_path} not exist.")
        return

    # Create a sheet for record change and delete files.
    ss1 = ezsheets.Spreadsheet()
    sheet = ss1.sheets[0]
    sheet[1, 1] = "Converted files"
    sheet[2, 1] = "Deleted files"

    # Get the files in the folder, set the convert number and delete number as 0.
    convert_num = []
    delete_num = []
    for file in folder_path.iterdir():
        if not file.is_file():
            continue
        # The excel file upload to google sheet and convert to pdf form.
        if file.suffix.lower() == ('.xlsx'):
            
            ss = ezsheets.upload(str(file))
            print(f"Uploading {file.name}...")
            ss.downloadAsPDF()
            print(f"Converting {file.name} to PDF...")
            convert_num.append(file.name)

        # Delete the empty pdf files.
        elif file.suffix.lower() == ('.pdf') and file.stat().st_size < 10:
            
            send2trash.send2trash(file)
            print(f"Deleting empty {file.name}...")
            delete_num.append(file.name)           

    # Record the change to sheet
    for index, file in enumerate(convert_num,start=2):
        sheet[1, index] = file
    for index,file in enumerate(delete_num, start=2):
        sheet[2, index] = file

    print(f"Converted {len(convert_num)} Excel files to PDF, deleted {len(delete_num)} empty PDF files.")

The_swiss_army_knife("your folder path")
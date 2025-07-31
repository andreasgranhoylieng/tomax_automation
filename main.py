import zipfile
import os
from pdfminer.high_level import extract_text
import shutil
import pandas as pd
from datetime import datetime, timedelta
import PyPDF2
import matplotlib.pyplot as plt


def unzip_file(zip_path, extract_path):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
        print(f'File unzipped to {extract_path}')



def find_CoC_excel(folder_name, numbers):
    initial_dir = os.getcwd()
    os.chdir(folder_name)
    files_in_folder = os.listdir()
    
    for file_name in files_in_folder:
        if "CoC" in file_name and (file_name.endswith('.xlsx') or file_name.endswith('.xls')) and numbers in file_name:
            excel_path = os.path.join(folder_name, file_name) 
            os.chdir(initial_dir) 
            return excel_path
    
    os.chdir(initial_dir)
    return None

def extract_data_from_excel(excel_path):
    data_dict = {}
    xls = pd.ExcelFile(excel_path)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        header_row = None
        for i in range(len(df)):
            if 'Serial number' in df.iloc[i].values and 'Heatno' in df.iloc[i].values:
                header_row = i
                break

        if header_row is not None:
            df.columns = df.iloc[header_row]
            df = df.drop(df.index[header_row])

            for index, row in df.iterrows():
                serial_number = row['Serial number']
                heatno = row['Heatno']

                # Check if the row contains valid data
                if pd.notna(serial_number) and pd.notna(heatno):
                    if serial_number not in data_dict:
                        try:
                            data_dict[serial_number] = int(heatno)
                        except:
                            data_dict[serial_number] = heatno

    return data_dict


def extract_metadata(pdf_path):
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf = PyPDF2.PdfReader(pdf_file)
            metadata = pdf.metadata
            
            if metadata is None:
                return None
            
            s = metadata.get('/ModDate', metadata.get('/CreationDate', None))

            if s is None:
                return None

            s = s[2:]

            dt_str = s[:14]

            dt = datetime.strptime(dt_str, "%Y%m%d%H%M%S")

            offset_str = s[14:]
            if offset_str == 'Z':  # UTC
                offset_hours = 0
                offset_minutes = 0
            else:
                offset_hours = int(offset_str[:3])
                offset_minutes = int(offset_str[4:6])
                if offset_str[0] == '+':
                    dt -= timedelta(hours=offset_hours, minutes=offset_minutes)
                else:
                    dt += timedelta(hours=offset_hours, minutes=offset_minutes)

            return dt
    except:
        return None


def search_pdfs_for_string(root_folder, search_string, custom_folder, pdf_cache):
    matching_pdfs = []

    ignore_folder = custom_folder
    

    for dirpath, dirnames, filenames in os.walk(root_folder):

        if ignore_folder in dirpath:
            continue


        for filename in filenames:
            if filename.endswith('.pdf') and ("CoC" in filename or "MTC" in filename):
                pdf_path = os.path.join(dirpath, filename)
                
                # Check the cache first
                text = pdf_cache.get(pdf_path, None)
                
                if text is None:
                    try:
                        text = extract_text(pdf_path)
                        pdf_cache[pdf_path] = text  # Store into cache
                    except Exception as e:
                        print(f"Failed to process {pdf_path} due to {e}")
                        continue
            

                if f"{search_string}" in text:
                    matching_pdfs.append(pdf_path)


    return matching_pdfs


def copy_and_rename_pdfs(pdf_paths, new_folder, search_for):
    if not os.path.exists(new_folder):
        os.makedirs(new_folder)
        
    latest_mtc_time = datetime.min
    latest_mtc_path = ''

    # Check the latest updated PDF for CoC and MTC
    for pdf_path in pdf_paths:
        modified_time = extract_metadata(pdf_path)
        
        try:
            if 'MTC' in pdf_path and modified_time > latest_mtc_time:

                latest_mtc_time = modified_time
                latest_mtc_path = pdf_path
        except:
            pass
    

    try:
        modified_search_for = search_for.replace("/","")
        shutil.copy(latest_mtc_path, os.path.join(new_folder, f"{modified_search_for} -MTC.pdf"))
        print(f'Finished saving MTC for: {search_for}')
    except:
        print(f"Did not find MTC containing: {search_for}")



def find_and_mark_excel(coc_excel, serial_number, save_folder):
    xls = pd.ExcelFile(coc_excel)
    
    for sheet_name in xls.sheet_names:
        print(f"Searching in sheet: {sheet_name}") # NEW
        df = xls.parse(sheet_name)
        
        if df.isin([serial_number]).any().any():
            row_idx = df[df.eq(serial_number).any(axis=1)].index.tolist()[0]

            if not os.path.exists(save_folder):
                os.makedirs(save_folder)
            
            modified_serial_number = serial_number.replace("/","")
            save_path = os.path.join(save_folder, f'{modified_serial_number} -CoC.xlsx')
            writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            green_format = workbook.add_format({'bg_color': '#C6EFCE'})
            last_column_idx = df.shape[1] 
            last_column_letter = chr(65 + last_column_idx - 1)
            excel_range = f"A{row_idx + 2}:{last_column_letter}{row_idx + 2}" 
            worksheet.conditional_format(excel_range, {'type': 'no_errors', 'format': green_format})
        
            writer.close()


            df = pd.read_excel(save_path)
            
            plt.figure(figsize=(12, 8))
            plt.table(cellText=df.values, colLabels=df.columns, cellLoc='center', loc='center')
            plt.axis('off')
            
            pdf_filename = f'{modified_serial_number} -CoC.pdf'
            pdf_path = os.path.join(save_folder, pdf_filename)
            plt.savefig(pdf_path)
            
            plt.close()


            return

    print(f"Could not find {serial_number} in the excel file. Something is wrong")


all_folders = [item for item in os.listdir() if os.path.isdir(item)]

# Remove the specified folders
folders_to_remove = ['Documents', '.DS_Store']
for folder in folders_to_remove:
    if folder in all_folders:
        all_folders.remove(folder)



for folder in all_folders:
    dir_content = os.listdir(folder)
    
    # Filter out the zip file
    zip_file = next((file for file in dir_content if file.endswith('.zip')), None)
    if zip_file:
        # Create a directory name based on the zip file name without the extension
        folder_name = os.path.splitext(zip_file)[0]
        extract_path = os.path.join(os.getcwd(), folder, folder_name)
        
        
        zip_file_path = os.path.join(os.getcwd(), folder, zip_file)
        unzip_file(zip_file_path, extract_path)
        
        numbers = ''.join([char for char in folder_name if char.isdigit()])
        
        # Search for the CoC Excel file in the extracted folder
        coc_excel_path = find_CoC_excel(extract_path, numbers)  # Replace "numbers" with the actual search string or variable
        if coc_excel_path:
            print(f"Found CoC Excel file: {coc_excel_path}")
            
            data = extract_data_from_excel(coc_excel_path)
            #print(f"Extracted data: {data}")
            print("Saving data to cache....")
            
            pdf_cache = {}
            
            custom_folder = "Documents"
            for key in data:
                find_and_mark_excel(coc_excel_path, key, custom_folder)
                matching_pdfs = search_pdfs_for_string(os.getcwd(), str(data[key]), custom_folder, pdf_cache)
                copy_and_rename_pdfs(matching_pdfs, custom_folder, key)


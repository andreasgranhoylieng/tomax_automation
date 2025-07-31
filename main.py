import tomllib
import shutil
import pandas as pd
import PyPDF2
import matplotlib.pyplot as plt
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


def load_config(config_path: Path) -> dict[str, Any]:
    """Loads the configuration from a .toml file."""
    print(f"Loading configuration from: {config_path}")
    with open(config_path, "rb") as f:
        return tomllib.load(f)


def unzip_file(zip_path: Path, extract_path: Path) -> None:
    """Unzips a file to the specified extraction path."""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    print(f"File unzipped to: {extract_path}")


def find_coc_excel(folder_path: Path, numbers: str, coc_keyword: str) -> Path | None:
    """Finds the Certificate of Conformity Excel file in a given folder."""
    for file_path in folder_path.iterdir():
        if (coc_keyword in file_path.name and
                file_path.suffix in ['.xlsx', '.xls'] and
                numbers in file_path.name):
            return file_path
    return None


def extract_data_from_excel(excel_path: Path, serial_header: str, heatno_header: str) -> dict[str, int | str]:
    """Extracts serial numbers and heat numbers from the CoC Excel file."""
    data_dict: dict[str, int | str] = {}
    xls = pd.ExcelFile(excel_path)

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        header_row_index = None
        for i, row in df.iterrows():
            if serial_header in row.values and heatno_header in row.values:
                header_row_index = i
                break

        if header_row_index is not None:
            df.columns = df.iloc[header_row_index]
            df = df.drop(df.index[header_row_index])

            for _, row in df.iterrows():
                serial_number = row.get(serial_header)
                heatno = row.get(heatno_header)

                if pd.notna(serial_number) and pd.notna(heatno):
                    try:
                        data_dict[str(serial_number)] = int(heatno)
                    except (ValueError, TypeError):
                        data_dict[str(serial_number)] = str(heatno)
    return data_dict


def extract_metadata_date(pdf_path: Path) -> datetime | None:
    """Extracts the modification or creation date from PDF metadata."""
    try:
        with open(pdf_path, "rb") as pdf_file:
            pdf = PyPDF2.PdfReader(pdf_file)
            metadata = pdf.metadata
            if not metadata:
                return None

            date_str_raw = metadata.get('/ModDate') or metadata.get('/CreationDate')
            if not date_str_raw:
                return None

            # Format is D:YYYYMMDDHHMMSS+HH'MM'
            dt_str = date_str_raw.split('Z')[0].split('+')[0].split('-')[0].strip("D:")
            dt_obj = datetime.strptime(dt_str, "%Y%m%d%H%M%S")
            return dt_obj
    except Exception as e:
        print(f"Could not extract metadata from {pdf_path.name}: {e}")
        return None


def search_pdfs_for_string(
    root_folder: Path,
    search_string: str,
    output_folder: Path,
    pdf_cache: dict[Path, str],
    coc_keyword: str,
    mtc_keyword: str
) -> list[Path]:
    """Searches all relevant PDFs for a specific string."""
    matching_pdfs: list[Path] = []
    for pdf_path in root_folder.rglob('*.pdf'):
        if output_folder in pdf_path.parents:
            continue

        if coc_keyword in pdf_path.name or mtc_keyword in pdf_path.name:
            text = pdf_cache.get(pdf_path)
            if text is None:
                try:
                    # Using pdfminer.high_level.extract_text as in the original script
                    from pdfminer.high_level import extract_text
                    text = extract_text(str(pdf_path))
                    pdf_cache[pdf_path] = text
                except Exception as e:
                    print(f"Failed to process {pdf_path.name} due to: {e}")
                    continue

            if search_string in text:
                matching_pdfs.append(pdf_path)
    return matching_pdfs


def copy_and_rename_pdfs(pdf_paths: list[Path], output_folder: Path, serial_number: str, mtc_keyword: str) -> None:
    """Finds the latest MTC PDF and copies it with a new name."""
    output_folder.mkdir(exist_ok=True)
    latest_mtc_time = datetime.min
    latest_mtc_path = None

    for pdf_path in pdf_paths:
        if mtc_keyword in pdf_path.name:
            modified_time = extract_metadata_date(pdf_path)
            if modified_time and modified_time > latest_mtc_time:
                latest_mtc_time = modified_time
                latest_mtc_path = pdf_path

    if latest_mtc_path:
        sanitized_serial = serial_number.replace("/", "_").replace("\\", "_")
        new_filename = f"{sanitized_serial} -{mtc_keyword}.pdf"
        try:
            shutil.copy(latest_mtc_path, output_folder / new_filename)
            print(f"Finished saving MTC for: {serial_number}")
        except Exception as e:
            print(f"Could not copy file {latest_mtc_path.name}: {e}")
    else:
        print(f"Did not find MTC containing the required data for: {serial_number}")


def find_and_mark_excel(coc_excel_path: Path, serial_number: str, save_folder: Path) -> None:
    """Finds a serial number in an Excel file, highlights the row, and saves it as PDF and XLSX."""
    xls = pd.ExcelFile(coc_excel_path)
    save_folder.mkdir(exist_ok=True)
    sanitized_serial = serial_number.replace("/", "_").replace("\\", "_")

    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        if df.isin([serial_number]).any().any():
            row_idx = df[df.eq(serial_number).any(axis=1)].index.tolist()[0]
            save_path_xlsx = save_folder / f'{sanitized_serial} -CoC.xlsx'
            writer = pd.ExcelWriter(str(save_path_xlsx), engine='xlsxwriter')
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            green_format = workbook.add_format({'bg_color': '#C6EFCE'})
            worksheet.conditional_format(row_idx + 1, 0, row_idx + 1, df.shape[1] - 1, {
                'type': 'no_errors', 'format': green_format
            })
            writer.close()

            # Save as PDF
            plt.figure(figsize=(12, 8))
            ax = plt.subplot(111, frame_on=False)
            ax.xaxis.set_visible(False)
            ax.yaxis.set_visible(False)
            table = pd.plotting.table(ax, df, loc='center', colWidths=[0.1] * len(df.columns))
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1.2, 1.2)
            
            pdf_path = save_folder / f'{sanitized_serial} -CoC.pdf'
            plt.savefig(str(pdf_path), bbox_inches='tight', pad_inches=0.1)
            plt.close()
            return

    print(f"Could not find '{serial_number}' in {coc_excel_path.name}.")


def main() -> None:
    """Main function to run the document processing workflow."""
    try:
        config = load_config(Path("config.toml"))
    except FileNotFoundError:
        print("Error: config.toml not found. Please create it in the script's directory.")
        return
    except Exception as e:
        print(f"Error loading configuration: {e}")
        return

    # --- Setup paths and terms from config ---
    root_path = Path(config["paths"]["root_directory"])
    output_folder_name = config["paths"]["output_folder_name"]
    output_path = root_path / output_folder_name
    
    terms = config["search_terms"]
    coc_kw = terms["certificate_of_conformity"]
    mtc_kw = terms["material_test_certificate"]
    serial_header = terms["excel_serial_header"]
    heatno_header = terms["excel_heatno_header"]
    
    folders_to_ignore = set(config["settings"]["ignore_list"] + [output_folder_name])

    if not root_path.is_dir():
        print(f"Error: The root directory '{root_path}' does not exist.")
        return

    all_folders = [item for item in root_path.iterdir() if item.is_dir() and item.name not in folders_to_ignore]

    for folder in all_folders:
        print(f"\n--- Processing folder: {folder.name} ---")
        zip_file = next(folder.glob('*.zip'), None)
        if not zip_file:
            print(f"No .zip file found in {folder.name}. Skipping.")
            continue

        extract_path = folder / zip_file.stem
        unzip_file(zip_file, extract_path)

        numbers = ''.join(filter(str.isdigit, extract_path.name))
        coc_excel_path = find_coc_excel(extract_path, numbers, coc_kw)

        if coc_excel_path:
            print(f"Found CoC Excel file: {coc_excel_path.name}")
            data = extract_data_from_excel(coc_excel_path, serial_header, heatno_header)
            print("Extracted serial and heat numbers. Processing each entry...")
            
            pdf_cache: dict[Path, str] = {}
            for serial, heat_no in data.items():
                print(f"-> Searching for Serial: {serial}, Heat No: {heat_no}")
                find_and_mark_excel(coc_excel_path, serial, output_path)
                matching_pdfs = search_pdfs_for_string(root_path, str(heat_no), output_path, pdf_cache, coc_kw, mtc_kw)
                copy_and_rename_pdfs(matching_pdfs, output_path, serial, mtc_kw)
        else:
            print(f"Could not find a CoC Excel file in {extract_path.name}.")

    print("\n--- Script finished ---")


if __name__ == "__main__":
    main()
# Automated Document Processor for CoC and MTC

This script automates the process of finding, processing, and organizing Certificate of Conformity (CoC) and Material Test Certificate (MTC) documents. It unzips project folders, reads serial numbers from a master CoC Excel file, finds the corresponding MTC PDF files by searching for heat numbers, and saves organized copies of the relevant documents into a designated output folder.

## Features

-   **Configuration Driven**: All paths, folder names, and search terms are managed in a `config.toml` file, making the script adaptable without code changes.
-   **Automated Unzipping**: Automatically finds and extracts `.zip` files in project subdirectories.
-   **Excel Data Extraction**: Reads serial numbers and corresponding heat numbers from a CoC Excel file.
-   **PDF Content Search**: Scans MTC and CoC PDF files for specific heat numbers to find correct documentation.
-   **Document Generation**: Creates highlighted copies of the CoC Excel file for each serial number and saves them as both `.xlsx` and `.pdf`.
-   **File Organization**: Copies and renames the relevant MTC and CoC files into a centralized output directory for easy access.

## Prerequisites

-   **Windows Operating System or Mac**
-   **Python 3.11**: The script has been tested with this version.
-   **Miniconda**: A minimal installer for Conda. If you don't have it, you can download it here:
    -   [**Download Miniconda**](https://docs.conda.io/en/latest/miniconda.html).

## 1. Setup and Installation

Follow these steps to set up the required environment using Conda.

### Step 1: Create the Conda Environment

1.  Open the **Anaconda Prompt** from the Start Menu (or terminal on mac).
2.  Navigate to the project directory where you saved the files (`main.py`, `config.toml`, etc.).
    ```bash
    cd C:\path\to\your\project_folder
    ```
3.  Create a new Conda environment from the `environment_windows.yml` file (or `environment_mac.yml` on mac). This command sets up an environment named `doc_processor` with Python 3.11 and all necessary libraries.
    ```bash
    conda env create -f environment_windows.yml
    ```
   

### Step 2: Activate the Environment

Before running the script, you must activate the Conda environment you just created.

```bash
conda activate doc_processor
```

You will need to run this activation command every time you open a new terminal to work on this project.

## 2. Configuration

Before running the script, you **must** configure the `config.toml` file.

1.  Open `config.toml` in a text editor.
2.  Modify the `root_directory` path to the **absolute path** of the folder containing your project subdirectories (e.g., `C:\\Users\\YourUser\\Documents\\Projects`).
3.  Adjust the `output_folder_name` and other search terms if they differ from the defaults.


## 3. How to Run the Script

Once the environment is activated and the `config.toml` file is set up, run the script from the Anaconda Prompt (or terminal):

```bash
python main.py
```

The script will print its progress in the terminal, indicating which folders and files it is processing.

## 4. Output

The script will create a new folder inside your specified `root_directory` (e.g., `Processed_Documents`). Inside this folder, you will find the generated documents, named according to their serial number:

-   `{serial_number} -CoC.pdf`: A PDF version of the CoC with the relevant row highlighted.
-   `{serial_number} -CoC.xlsx`: An Excel version of the CoC with the relevant row highlighted.
-   `{serial_number} -MTC.pdf`: The corresponding Material Test Certificate.


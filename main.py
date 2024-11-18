import os
import logging
import time
import json
from tqdm import tqdm
from ds.converter import PDF_Process, PPTX_Process, DOCX_Process, load_and_flatten_urls
import pandas as pd
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

File_Converter_Run = True

# ############################################### Configuration ###############################################################
json_scrape_export_path = 'processed/proj_data.json'                # Path to store the JSON file
json_proj_files_export_path = 'processed/proj_files.json'           # Path to store the JSON file
json_url_import_path = 'urls.json'                                  # Path to import the URL JSON file
project_info_csv = 'ds/PBI.csv'                                     # Path to the project information CSV file
project_base_directory = 'downloads'                                # Path to the base directory containing the project files
File_Converter_Run = False                                          # Flag to run text extraction from PDF, DOCX, and  to JSON converter
# #############################################################################################################################

# Setup logging
start_time = time.time()
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
# Ensure the logs directory exists
if not os.path.exists('logs'):
    os.makedirs('logs')

# Setup file handler for logging
file_handler = logging.FileHandler(f'logs/file_converter_{time.strftime("%Y%m%d_%H%M%S")}.log')
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

#setup logger handler
logger.addHandler(file_handler)
logger.info('Starting XML to JSON conversion')


# Run the file converter
if File_Converter_Run:
    logger.addHandler(file_handler)
    start_time = time.time()
    logger.info('Starting File Conversion')

    # Load the flattened URLs data
    flattened_urls_data = load_and_flatten_urls(json_url_import_path)

    # Define the base directory
    bpi_df = pd.read_csv(project_info_csv)

    # Define the base directory
    base_directory = os.path.join(project_base_directory)

    # Dictionary to store the extracted text
    proj_files = {}

    # Total projects
    total_dirs = sum([len(dirs) for _, dirs, _ in os.walk(base_directory)])

    # Walk through the directory
    i = 0
    for root,_,files in tqdm(os.walk(base_directory), desc='Scanning directories', unit='dir', total=total_dirs, position=0, leave=True, ncols=100):
        for file in files:
            pid = root.split('\\')[1]
            proj_name = bpi_df.loc[bpi_df['ProjectId'] == pid]['ProjectName'].values[0].replace(' ', '%20')
            proj_link = f'<your_project_url>{proj_name}.aspx'
            # PDF file processing
            if file.endswith('.pdf'):
                pdf_file = os.path.join(root, file)
                # Check to see if file URL exists in the JSON file; if not, use the project link
                try:
                    link = flattened_urls_data[pid][file]
                except KeyError:
                    link = proj_link
                # Extract text from the PDF
                try:
                    PDF_Obj = PDF_Process(pdf_path=pdf_file)
                    extracted_doc = PDF_Obj.process_text
                    pdf_type = 'pdf_ocr' if PDF_Obj.req_ocr else 'pdf'
                    # Store the extracted text in the dictionary
                    if pid not in proj_files:
                        proj_files[pid] = []
                    proj_files[pid].append({'project_link': proj_link, 'link': link, 'root': root, 'filetype' : pdf_type ,'filename' : file ,'extracted_text': extracted_doc})
                except Exception as e:
                    logger.warning(f'Warning on obtaining url for {pdf_file}: {e}')

            # DOCX file processing
            elif file.endswith('.docx'):
                docx_file = os.path.join(root, file)
                # Check to see if file URL exists in the JSON file; if not, use the project link
                try:
                    link = flattened_urls_data[pid][file]
                except KeyError:
                    link = proj_link
                # Extract text from the DOCX
                try:
                    extracted_doc = DOCX_Process(docx_file).extracted_text
                    # Store the extracted text in the dictionary
                    if pid not in proj_files:
                        proj_files[pid] = []
                    proj_files[pid].append({'project_link': proj_link, 'link': link,'root': root, 'filetype' : 'docx', 'filename' : file, 'extracted_text': extracted_doc})
                except Exception as e:
                    logger.warning(f'Warning on obtaining url for {docx_file}: {e}')

            # PPTX file processing
            elif file.endswith('.pptx'):
                pptx_file = os.path.join(root, file)
                # Check to see if file URL exists in the JSON file; if not, use the project link
                try:
                    link = flattened_urls_data[pid][file]
                except KeyError:
                    link = proj_link
                # Extract text from the PowerPoint
                try:
                    extracted_ppt = PPTX_Process(pptx_file).extracted_text
                    # Store the extracted text in the dictionary
                    if pid not in proj_files:
                        proj_files[pid] = []
                    proj_files[pid].append({'project_link': proj_link, 'link': link,'root': root, 'filetype' : 'pptx','filename' : file, 'extracted_text': extracted_ppt})
                except Exception as e:
                    logger.warning(f'Warning on obtaining url for {pptx_file}: {e}')
        # i += 1
        # if i == 1: break

    # Export the extracted text to a JSON file
    with open(json_proj_files_export_path, 'w') as file:
        json.dump(proj_files, file, indent=4)

    # Log total time taken
    logger.info(f'Total time taken: {round((time.time() - start_time) / 60, 1)} minutes')
    logger.info('File Conversion Completed')
else:
    logger.info('Skipping File Conversion')

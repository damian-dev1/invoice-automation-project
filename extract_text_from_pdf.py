import os
import re
import sys
import csv
import time
import logging
import subprocess
from datetime import datetime
from pathlib import Path
import fitz

def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    logging.critical("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

sys.excepthook = handle_exception

logging.basicConfig(
    filename='invoice_extraction.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logging.getLogger().addHandler(console_handler)

input_dir = Path('./pdf_invoices')
ocr_output_dir = Path('./output')
ocr_output_dir.mkdir(parents=True, exist_ok=True)
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
csv_output_path = Path(f'./extracted_orders_{timestamp}.csv')

ORDER_PATTERN = re.compile(
    r'(?i)(?:purchase\s*order|po|order\s*no\.?)?\s*[:\-]?\s*(3100\d{7})'
)

def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()
        return text
    except Exception as e:
        logging.error(f"Failed to extract text from {pdf_path.name}: {e}", exc_info=True)
        return ""

def run_ocr(input_pdf_path, output_pdf_path):
    try:
        result = subprocess.run(
            [
                'ocrmypdf',
                '--force-ocr',
                '--optimize', '3',
                '--deskew',
                '--clean',
                '--output-type', 'pdf',
                str(input_pdf_path),
                str(output_pdf_path)
            ],
            check=True,
            capture_output=True,
            text=True
        )
        logging.info(f"OCR completed for: {input_pdf_path.name}")
        if result.stdout:
            logging.info(f"OCR stdout for {input_pdf_path.name}:\n{result.stdout}")
        if result.stderr:
            logging.warning(f"OCR stderr for {input_pdf_path.name}:\n{result.stderr}")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"OCR failed for {input_pdf_path.name}: {e}", exc_info=True)
        if e.stderr:
            logging.error(f"OCR stderr:\n{e.stderr}")
        return False

def extract_order_numbers(text):
    matches = ORDER_PATTERN.findall(text)
    return list(set(matches))

def process_pdf(pdf_path):
    logging.info(f"Processing: {pdf_path.name}")
    try:
        text = extract_text_from_pdf(pdf_path)
        if text.strip():
            logging.info(f"Text found in {pdf_path.name}, extracting order numbers.")
            return extract_order_numbers(text)
        else:
            logging.info(f"No text found in {pdf_path.name}, running OCR.")
            ocr_pdf_path = ocr_output_dir / pdf_path.name
            if run_ocr(pdf_path, ocr_pdf_path):
                ocr_text = extract_text_from_pdf(ocr_pdf_path)
                return extract_order_numbers(ocr_text)
            else:
                return []
    except Exception as e:
        logging.error(f"Error during processing of {pdf_path.name}: {e}", exc_info=True)
        return []

def main():
    start_time = time.time()
    pdf_files = list(input_dir.glob('*.pdf'))
    total_files = len(pdf_files)
    extracted_data = []

    if not pdf_files:
        logging.warning("No PDF files found in the input directory.")
        return

    for pdf_file in pdf_files:
        try:
            order_numbers = process_pdf(pdf_file)
            if order_numbers:
                logging.info(f"Order numbers found in {pdf_file.name}: {order_numbers}")
                for order in order_numbers:
                    extracted_data.append([pdf_file.name, order])
            else:
                logging.info(f"No order numbers found in {pdf_file.name}.")
        except Exception as e:
            logging.error(f"Unexpected error while processing {pdf_file.name}: {e}", exc_info=True)

    try:
        with open(csv_output_path, mode='w', newline='') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(['Filename', 'Order Number'])
            writer.writerows(extracted_data)
        logging.info(f"CSV export completed: {csv_output_path}")
    except Exception as e:
        logging.error(f"Failed to write CSV file: {e}", exc_info=True)

    elapsed_time = time.time() - start_time
    logging.info(f"Processing complete. Total files processed: {total_files}. Time taken: {elapsed_time:.2f} seconds.")

if __name__ == '__main__':
    main()

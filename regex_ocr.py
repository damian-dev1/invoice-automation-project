import re
import csv
import logging
import subprocess
from pathlib import Path
import fitz
import pandas as pd

logging.basicConfig(
    filename='invoice_extraction.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

input_dir = Path('invoices_in')
ocr_output_dir = Path('ocr_output')
csv_summary_path = Path('extracted_invoice_data.csv')
csv_line_items_path = Path('line_items.csv')
onedrive_path = Path('C:/Users/YourUsername/OneDrive/Invoices')
onedrive_path.mkdir(parents=True, exist_ok=True)
ocr_output_dir = Path('ocr_output')
ocr_output_dir.mkdir(parents=True, exist_ok=True)

ORDER_PATTERN = re.compile(r'(?i)(?:purchase\s*order|po|order\s*no\.?)?\s*[:\-]?\s*(3100\d{7})')
INVOICE_PATTERN = re.compile(r'(?i)(invoice[\s:_#-]*no\.?|inv[\s:_#-]*number)?\s*[:#-]?\s*([A-Z0-9\-\/]{4,})')
DATE_PATTERN = re.compile(r'(?i)(invoice\s*date|date\s*of\s*issue)\s*[:#-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})')
DUE_DATE_PATTERN = re.compile(r'(?i)(due\s*date)\s*[:#-]?\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})')
TOTAL_PATTERN = re.compile(r'(?i)(grand\s*total|total\s*amount|amount\s*due)\s*[:$AUD\s]*([$€£]?\s*\d{1,3}(?:,\d{3})*(?:\.\d{2})?)')
FREIGHT_PATTERN = re.compile(r'(?i)(freight[\s_]*(inc)?[\s_]*gst)?\s*[:\-]?\s*([$€£]?\s*\d+(?:\.\d{2})?)')
SUPPLIER_PATTERN = re.compile(r'(?i)(from|seller|vendor|supplier)\s*[:\-]?\s*(.+)')
ABN_PATTERN = re.compile(r'(?i)(ABN|GST\s*number|VAT\s*number|Tax\s*ID)[\s:]*([A-Z0-9\- ]{8,})')
PO_PATTERN = re.compile(r'(?i)(PO[\s_-]?Number|Purchase\s*Order|Reference)\s*[:#-]?\s*([A-Z0-9\-\/]{4,})')

def extract_text_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        return "\n".join(page.get_text() for page in doc)
    except:
        return ""

def run_ocr(input_pdf_path, output_pdf_path):
    try:
        subprocess.run(
            ['ocrmypdf', '--force-ocr', '--optimize', '3', '--deskew', '--clean', '--output-type', 'pdf', str(input_pdf_path), str(output_pdf_path)],
            check=True
        )
        return True
    except:
        return False

def extract_fields(text):
    order = ORDER_PATTERN.findall(text)
    invoice = INVOICE_PATTERN.search(text)
    date = DATE_PATTERN.search(text)
    due = DUE_DATE_PATTERN.search(text)
    total = TOTAL_PATTERN.search(text)
    supplier = SUPPLIER_PATTERN.search(text)
    abn = ABN_PATTERN.search(text)
    po = PO_PATTERN.search(text)
    freight = FREIGHT_PATTERN.findall(text)
    return {
        'order_number': ', '.join(sorted(set(order))) if order else '',
        'invoice_number': invoice.group(2).strip() if invoice else '',
        'invoice_date': date.group(2).strip() if date else '',
        'due_date': due.group(2).strip() if due else '',
        'total_amount': total.group(2).strip() if total else '',
        'freight_inc_gst': freight[-1][2].strip() if freight else '0.00',
        'supplier': supplier.group(2).split('\n')[0].strip() if supplier else '',
        'abn': abn.group(2).strip() if abn else '',
        'po_number': po.group(2).strip() if po else ''
    }

def extract_line_items(text):
    lines = text.splitlines()
    items = []
    header_keywords = ['description', 'sku', 'qty', 'quantity', 'unit', 'price', 'amount']
    header_found = False
    for i, line in enumerate(lines):
        if not header_found and all(any(h in part.lower() for part in line.lower().split()) for h in header_keywords):
            header_found = True
            continue
        if header_found:
            if line.strip() == "" or re.match(r"(?i)(total|subtotal|gst|grand)", line):
                break
            fields = re.split(r'\s{2,}|\t', line.strip())
            if len(fields) >= 4:
                items.append(fields[:5])
    return items

def process_pdf(pdf_path):
    text = extract_text_from_pdf(pdf_path)
    if not text.strip():
        ocr_pdf_path = ocr_output_dir / pdf_path.name
        if run_ocr(pdf_path, ocr_pdf_path):
            text = extract_text_from_pdf(ocr_pdf_path)
    if text.strip():
        fields = extract_fields(text)
        items = extract_line_items(text)
        return fields, items
    return {}, []

def write_to_csv(results, summary_csv, line_csv):
    summary_data = []
    line_data = []

    for filename, (fields, items) in results.items():
        fields['pdf_filename'] = filename
        summary_data.append(fields)
        orders = fields['order_number'].split(', ')
        for idx, item in enumerate(items):
            clean = item + [''] * (5 - len(item))
            line_data.append({
                'pdf_filename': filename,
                'line_index': idx,
                'order_number': orders[0] if len(orders) == 1 else ';'.join(orders),
                'invoice_number': fields['invoice_number'],
                'sku': clean[0],
                'description': clean[1],
                'qty': clean[2],
                'unit_price': clean[3],
                'amount': clean[4],
                'freight_inc_gst': fields.get('freight_inc_gst', '0.00')
            })

    pd.DataFrame(summary_data).to_csv(summary_csv, index=False)
    pd.DataFrame(line_data).to_csv(line_csv, index=False)

    try:
        (onedrive_path / summary_csv.name).write_bytes(summary_csv.read_bytes())
        (onedrive_path / line_csv.name).write_bytes(line_csv.read_bytes())
        logging.info(f"Exported to OneDrive: {onedrive_path}")
    except Exception as e:
        logging.error(f"Failed to copy to OneDrive: {e}")


def main():
    pdf_files = list(input_dir.glob('*.pdf'))
    results = {}
    for pdf in pdf_files:
        fields, lines = process_pdf(pdf)
        if fields:
            results[pdf.name] = (fields, lines)
    write_to_csv(results, csv_summary_path, csv_line_items_path)

if __name__ == '__main__':
    main()

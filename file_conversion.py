import argparse
import os
import platform
import subprocess
from fpdf import FPDF
import pandas as pd
from docx import Document
from markdown2 import markdown
import pdfkit
from pptx import Presentation

def parse_arguments():
    parser = argparse.ArgumentParser(
        description="File Conversion Task",
        epilog="""
Available conversions:
  - .txt   →  .pdf
  - .docx  →  .pdf
  - .md    →  .html
  - .html  →  .pdf
  - .xlsx  →  .csv
  - .pptx  →  .pdf
""",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("--dir", required=True, help="Directory containing files to convert")
    parser.add_argument("--ext", required=True, help="File extension to look for")
    parser.add_argument("--format", required=True, help="Target conversion format")
    parser.add_argument("--frequency", type=int, required=True, help="Frequency of the task")
    parser.add_argument("--unit", choices=["minute", "hour", "day"], required=True, help="Unit of time")
    return parser.parse_args()

def check_files(directory, extension):
    files = [f for f in os.listdir(directory) if f.endswith(extension)]
    if not files:
        print(f"❌ No files with extension {extension} found in {directory}")
    return files

def txt_to_pdf(txt_file, pdf_file):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        with open(txt_file, 'r', encoding='utf-8') as file:
            for line in file:
                pdf.multi_cell(0, 10, line)
        pdf.output(pdf_file)
        print(f"✅ Converted {txt_file} → {pdf_file}")
    except Exception as e:
        print(f"❌ Error converting {txt_file} to PDF: {e}")

def docx_to_pdf(docx_file, pdf_file):
    try:
        doc = Document(docx_file)
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Arial", size=12)
        for para in doc.paragraphs:
            pdf.multi_cell(0, 10, para.text)
        pdf.output(pdf_file)
        print(f"✅ Converted {docx_file} → {pdf_file}")
    except Exception as e:
        print(f"❌ Error converting {docx_file} to PDF: {e}")

def convert_file(file_path, target_format):
    base, ext = os.path.splitext(file_path)
    conversions = {
        (".txt", ".pdf"): txt_to_pdf,
        (".docx", ".pdf"): docx_to_pdf,
    }
    if (ext, target_format) in conversions:
        conversions[(ext, target_format)](file_path, base + target_format)
    else:
        print(f"❌ Conversion from {ext} to {target_format} is not supported.")

def schedule_task_windows(script_path, directory, extension, target_format, frequency, unit):
    task_name = "FileConversionTask"
    command = f'python "{script_path}" --dir "{directory}" --ext "{extension}" --format "{target_format}"'
    schedule_unit = {"minute": "MINUTE", "hour": "HOURLY", "day": "DAILY"}
    schedule_command = (
        f'schtasks /create /tn "{task_name}" /tr "{command}" '
        f'/sc {schedule_unit[unit]} /mo {frequency} /f'
    )
    try:
        subprocess.run(schedule_command, shell=True, check=True)
        print(f"✅ Scheduled task '{task_name}' in Windows every {frequency} {unit}(s).")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to schedule task in Windows: {e}")

def schedule_task_linux(script_path, directory, extension, target_format, frequency, unit):
    cron_time = f"*/{frequency} * * * *" if unit == "minute" else f"0 */{frequency} * * *" if unit == "hour" else f"0 0 */{frequency} * *"
    cron_job = f'{cron_time} python3 {script_path} --dir "{directory}" --ext "{extension}" --format "{target_format}"'
    subprocess.run(f'(crontab -l; echo "{cron_job}") | crontab -', shell=True, executable='/bin/bash')
    print(f"✅ Scheduled task in Linux every {frequency} {unit}(s).")

def main():
    args = parse_arguments()
    files = check_files(args.dir, args.ext)
    for file in files:
        convert_file(os.path.join(args.dir, file), args.format)

    script_path = os.path.abspath(__file__)
    if platform.system() == "Windows":
        schedule_task_windows(script_path, args.dir, args.ext, args.format, args.frequency, args.unit)
    elif platform.system() == "Linux":
        schedule_task_linux(script_path, args.dir, args.ext, args.format, args.frequency, args.unit)
    else:
        print("❌ Unsupported operating system.")

if __name__ == "__main__":
    main()

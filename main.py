#### main.py
####
#### This is the main entry point of the report generation system.
#### It handles the overall workflow, including:
####   - Checking for required Python packages.
####   - Prompting the user to regenerate an existing report or create a new one.
####   - Calling the report generation functions from the 'analysis' package.
####   - Converting the generated markdown report to a Word document.
####   - Logging key events and errors throughout the process.


import os
import subprocess
import sys
import time
import logging
from md_to_word import convert_md_to_docx
from analysis.report_generation import generate_report

# Configuration
file_path = './1k_lines_sales_data.xlsx'
model_name = "phi4:latest" # gives best results, takes 16 GB of GPU RAM
# model_name = "gemma3:12b" # gives 2nd best results, takes 12 GB of GPU RAM
# model_name = "deepseek-r1:1.5b" #smallest model takes 3 GB of GPU RAM adds <THINKING>
output_dir = "./output"
temperature = 0.1
md_report = os.path.join(output_dir, "executive_report.md")


def check_requirements():
    """Verify required packages are installed"""
    try:
        import pandas
        import docx
        import matplotlib
        import seaborn
        import requests
        return True
    except ImportError as e:
        print(f"Missing requirement: {str(e)}")
        print("Please install requirements: pip install -r requirements.txt")
        return False


def prompt_user(prompt, default=None):
    """Helper function for user prompts"""
    while True:
        response = input(f"{prompt} [y/n]: ").lower()
        if response in ('y', 'yes'):
            return True
        elif response in ('n', 'no'):
            return False
        elif default is not None and response == '':
            return default
        print("Please enter 'y' or 'n'")


def main():
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        filename='./report_generator.log'
    )

    # Start timing
    start_time = time.time()

    # Check requirements
    if not check_requirements():
        logging.error("Missing required packages")
        return

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Check if report exists
    if os.path.exists(md_report):
        print("\nExisting report found:", md_report)
        logging.info(f"Found existing report at {md_report}")
        if not prompt_user("Do you want to regenerate the report?"):
            print("Using existing report...")
            logging.info("Using existing report")
        else:
            print("Regenerating report...")
            logging.info("Regenerating report")
            report_md = generate_report(file_path, output_dir, model_name, temperature)
            if report_md:
                with open(md_report, 'w') as f:
                    f.write(report_md)
            else:
                print("Report generation failed")
                logging.error("Report generation failed")
                return
    else:
        print("No existing report found - generating new report...")
        logging.info("Generating new report")
        report_md = generate_report(file_path, output_dir, model_name, temperature)
        if report_md:
            with open(md_report, 'w') as f:
                f.write(report_md)
        else:
            print("Report generation failed")
            logging.error("Report generation failed")
            return

    # Convert to Word
    print("\nConverting markdown to Word document...")
    logging.info("Starting markdown to Word conversion")
    if convert_md_to_docx(md_report, output_dir):
        print("\nConversion complete!")
        print(f"Word document saved to: {os.path.join(output_dir, 'executive_report.docx')}")
        logging.info(f"Conversion successful. DOCX saved to {os.path.join(output_dir, 'executive_report.docx')}")
    else:
        print("\nConversion failed")
        logging.error("Markdown to Word conversion failed")

    # Log completion time
    logging.info(f"Completed in {time.time() - start_time:.2f} seconds")


if __name__ == "__main__":
    print("=== Report Generation System ===")
    main()
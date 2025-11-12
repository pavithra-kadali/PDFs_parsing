import os
import re
import sys
import json
import logging
import pdfplumber
import tkinter as tk
from tkinter import messagebox
import pandas as pd


def get_base_dir():
    """Get the base directory (supports PyInstaller .exe)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

def setup_logging(base_dir):
    """Set up logging directory and file."""
    logs_dir = os.path.join(base_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    log_path = os.path.join(logs_dir, "extract_log.log")
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        filemode="w"
    )
    return log_path

def load_config(base_dir):
    """Load config.json or create one if missing."""
    config_path = os.path.join(base_dir, "config.json")
    if not os.path.exists(config_path):
        default_config = {
            "input_path": "",
            "output_dir": ""
        }
        with open(config_path, "w") as f:
            json.dump(default_config, f, indent=4)
        raise FileNotFoundError("config.json created. Please edit it with your paths and re-run.")
    with open(config_path, "r") as f:
        return json.load(f)


def generate_file_name(pdf_name):
    base = re.sub(r"\.pdf$", "", os.path.basename(pdf_name), flags=re.IGNORECASE)
    safe_name = re.sub(r"[\\/:?*\[\]]", "_", base)
    return safe_name


def extract_metadata(pdf_path):
    """Extract metadata such as AUDIT ID, NCPDP, Date, etc."""
    metadata = {
        "AUDIT ID": "",
        "NCPDP": "",
        "Date": "",
        "Address": "",
        "Subject": ""
    }
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if not pdf.pages:
                return metadata
            full_text = "\n".join([page.extract_text() or "" for page in pdf.pages])

        audit_match = re.search(r"AUDIT\s*ID[:\s]*([A-Za-z0-9\-]+)", full_text, re.IGNORECASE)
        ncpdp_match = re.search(r"NCPDP[:\s]*([0-9]+)", full_text, re.IGNORECASE)
        date_match = re.search(r"Date[:\s]*(\d{1,2}/\d{1,2}/\d{2,4})", full_text, re.IGNORECASE)

        metadata["AUDIT ID"] = audit_match.group(1).strip() if audit_match else ""
        metadata["NCPDP"] = ncpdp_match.group(1).strip() if ncpdp_match else ""
        metadata["Date"] = date_match.group(1).strip() if date_match else ""

        addr_match = re.search(
            r"(?m)^(?:[A-Z\s]*\bPHARMACY\b[^\n]*\n(?:.*\n){0,5}?FAX[:\s]*\d+)",
            full_text, re.IGNORECASE
        )
        if addr_match:
            addr_text = addr_match.group(0)
            addr_text = re.sub(r"\s*\n\s*", ", ", addr_text)
            addr_text = re.sub(r"\s{2,}", " ", addr_text).strip(" ,")
            metadata["Address"] = addr_text

        subject_match = re.search(r"RE[:\s].{5,}", full_text)
        if subject_match:
            metadata["Subject"] = subject_match.group(0).strip()

        return metadata
    except Exception as e:
        logging.error(f"Metadata extraction failed for {pdf_path}: {e}")
        return metadata


def extract_tables(pdf_path):
    """Extract tables from the PDF using tabula + pdfplumber."""
    try:
        import tabula
    except Exception:
        logging.error("tabula-py or Java not found. Please ensure Java is installed and added to PATH.")
        show_message("Error", "Java not found.\nPlease install Java and try again.")
        return []

    try:
        all_tables = []
        seen_titles = set()

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                page_tables = tabula.read_pdf(pdf_path, pages=page_num, multiple_tables=True, lattice=True)
                if not page_tables:
                    continue

                detected_tables = page.find_tables()
                words = page.extract_words()
                line_map = {}
                for w in words:
                    y_key = round(w["bottom"], 1)
                    if y_key not in line_map:
                        line_map[y_key] = []
                    line_map[y_key].append(w["text"])
                line_texts = {k: " ".join(v).strip() for k, v in line_map.items()}

                for idx, table in enumerate(page_tables):
                    if not isinstance(table, pd.DataFrame) or table.empty:
                        continue

                    table = table.dropna(how='all').reset_index(drop=True)

                    title = ""
                    try:
                        if detected_tables and len(detected_tables) > idx:
                            x0, top, x1, bottom = detected_tables[idx].bbox
                            above_lines = {y: text for y, text in line_texts.items() if y < top and (top - y) < 120}
                            if above_lines:
                                nearest_y = max(above_lines.keys())
                                title = above_lines[nearest_y].strip()
                    except Exception as te:
                        logging.warning(f"Could not extract table title on page {page_num}: {te}")

                    if not title:
                        possible_title = str(table.columns[0])
                        if re.search(r"[A-Za-z]", possible_title):
                            title = possible_title
                        else:
                            title = f"Table_{page_num}_{idx+1}"

                    if title in seen_titles:
                        continue

                    seen_titles.add(title)
                    all_tables.append({
                        "title": title,
                        "data": table.fillna("").to_dict(orient="records")
                    })

        logging.info(f"Extracted {len(all_tables)} tables from {os.path.basename(pdf_path)}")
        return all_tables
    except Exception as e:
        logging.error(f"Table extraction failed for {pdf_path}: {e}")
        return []


def write_to_json(pdf_data, output_json):
    """Write extracted data to a JSON file."""
    try:
        json_serializable = {}
        for pdf_file, content in pdf_data.items():
            json_serializable[os.path.basename(pdf_file)] = {
                "metadata": content.get("metadata", {}),
                "tables": content.get("tables", [])
            }

        with open(output_json, "w", encoding="utf-8") as f:
            json.dump(json_serializable, f, indent=4, ensure_ascii=False)

        logging.info(f"✅ JSON file saved: {output_json}")
        return output_json
    except Exception as e:
        logging.error(f"Failed to write JSON: {e}")
        return None


def show_message(title, message):
    """Show popup messages for non-technical users."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()


def main():
    base_dir = get_base_dir()
    log_path = setup_logging(base_dir)
    logging.info("Starting PDF Table Extractor (JSON Output)...")

    try:
        config = load_config(base_dir)
        input_path = config.get("input_path")
        output_dir = config.get("output_dir")

        if not input_path or not os.path.exists(input_path):
            msg = f"Invalid input path: {input_path}"
            logging.error(msg)
            show_message("Error", msg)
            return

        if not output_dir:
            msg = "Output directory not specified in config.json."
            logging.error(msg)
            show_message("Error", msg)
            return

        os.makedirs(output_dir, exist_ok=True)
        output_json = os.path.join(output_dir, "extracted_table_data.json")

        # Identify PDF files
        if os.path.isdir(input_path):
            pdf_files = [os.path.join(input_path, f) for f in os.listdir(input_path) if f.lower().endswith(".pdf")]
        elif input_path.lower().endswith(".pdf"):
            pdf_files = [input_path]
        else:
            pdf_files = []

        if not pdf_files:
            msg = f"No PDF files found in {input_path}"
            logging.warning(msg)
            show_message("Warning", msg)
            return

        pdf_data = {}
        for pdf_file in pdf_files:
            logging.info(f"Processing: {pdf_file}")
            metadata = extract_metadata(pdf_file)
            tables = extract_tables(pdf_file)
            pdf_data[pdf_file] = {"metadata": metadata, "tables": tables}

        if pdf_data:
            json_path = write_to_json(pdf_data, output_json)
            if json_path:
                show_message("Success", f"Extraction completed.\nJSON saved to:\n{json_path}")
        else:
            msg = "No data extracted from any files."
            logging.warning(msg)
            show_message("Warning", msg)

        logging.info(f"Log file: {log_path}")

    except Exception as e:
        err = f"Unexpected error: {e}"
        logging.error(err)
        show_message("Error", err)


if __name__ == "__main__":
    main()

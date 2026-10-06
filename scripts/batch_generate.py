import os
import sys
import csv
import argparse
from typing import List

# Import create_certificate and get_default_paths from generate_certificate
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from generate_certificate import create_certificate, get_default_paths

def find_default_name_file() -> str:
    """
    Search for default names file in the project directory.
    Checks name.csv, input/name.xlsx, input/name.csv, etc.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)

    candidates = [
        os.path.join(project_root, "name.csv"),
        os.path.join(project_root, "input", "name.xlsx"),
        os.path.join(project_root, "input", "name.csv"),
        "/files/name.csv",
        "/files/input/name.xlsx",
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    return None

def load_names_from_csv(file_path: str) -> List[str]:
    """Load names from a CSV file."""
    names = []
    with open(file_path, mode="r", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = None
        name_col_idx = 0

        for row in reader:
            if not row or not any(cell.strip() for cell in row):
                continue
            cleaned_row = [cell.strip() for cell in row]

            if header is None:
                # Check if this row is a header
                possible_headers = ["name", "student name", "student", "full name", "participant"]
                for idx, cell in enumerate(cleaned_row):
                    if cell.lower() in possible_headers:
                        name_col_idx = idx
                        header = cleaned_row
                        break
                if header is not None:
                    continue  # Skip header row
                else:
                    # No recognizable header, treat first row as data if not matching common keywords
                    names.append(cleaned_row[0])
            else:
                if name_col_idx < len(cleaned_row) and cleaned_row[name_col_idx]:
                    names.append(cleaned_row[name_col_idx])

    return names

def load_names_from_excel(file_path: str) -> List[str]:
    """Load names from an Excel (.xlsx, .xls) file."""
    try:
        import pandas as pd
    except ImportError:
        raise ImportError("pandas and openpyxl are required to read Excel files. Run: pip install pandas openpyxl")

    df = pd.read_excel(file_path)
    if df.empty:
        return []

    # Identify name column
    name_col = None
    possible_names = ["name", "student name", "student", "full name", "participant"]
    for col in df.columns:
        if str(col).strip().lower() in possible_names:
            name_col = col
            break

    if name_col is None:
        # Default to the first column
        name_col = df.columns[0]

    names = df[name_col].dropna().astype(str).str.strip().tolist()
    # Filter out empty strings
    names = [n for n in names if n]
    return names

def load_names(file_path: str) -> List[str]:
    """Load names based on file extension."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return load_names_from_excel(file_path)
    elif ext in [".csv", ".txt"]:
        return load_names_from_csv(file_path)
    else:
        raise ValueError(f"Unsupported file format: {ext}. Please use .csv or .xlsx")

def run_batch_generation(file_path: str = None, template_path: str = None, output_dir: str = None, convert_to_pdf_flag: bool = False):
    """
    Read list of names and generate a certificate for each one.
    """
    if file_path is None:
        file_path = find_default_name_file()
        if not file_path:
            raise FileNotFoundError(
                "Could not find a names file (name.csv or input/name.xlsx). "
                "Please specify one using --file <path>."
            )

    print(f"Reading names from: {file_path}")
    names = load_names(file_path)

    if not names:
        print("Warning: No names found in the specified file.")
        return

    default_template, default_output = get_default_paths()
    template_path = template_path or default_template
    output_dir = output_dir or default_output

    print(f"Template path:    {template_path}")
    print(f"Output directory: {output_dir}")
    print(f"Convert to PDF:   {convert_to_pdf_flag}")
    print(f"Total names:      {len(names)}")
    print("-" * 50)

    success_count = 0
    fail_count = 0

    for idx, name in enumerate(names, 1):
        try:
            cert_path = create_certificate(name, template_path=template_path, output_dir=output_dir)
            msg = f"[{idx}/{len(names)}] Success: {name} -> {os.path.basename(cert_path)}"
            if convert_to_pdf_flag:
                from convert_to_pdf import convert_pptx_to_pdf
                pdf_path = convert_pptx_to_pdf(cert_path)
                msg += f" & {os.path.basename(pdf_path)}"
            print(msg)
            success_count += 1
        except Exception as e:
            print(f"[{idx}/{len(names)}] Failed for '{name}': {e}")
            fail_count += 1

    print("-" * 50)
    print(f"Batch generation completed: {success_count} succeeded, {fail_count} failed.")
    print(f"Certificates saved to: {output_dir}")

def main():
    parser = argparse.ArgumentParser(description="Batch generate certificates for a list of names.")
    parser.add_argument(
        "--file", "-f",
        help="Path to CSV or Excel file containing student names (default: name.csv or input/name.xlsx)",
        default=None
    )
    parser.add_argument(
        "--template", "-t",
        help="Path to PPTX template file (default: template/Certificate_template.pptx)",
        default=None
    )
    parser.add_argument(
        "--output", "-o",
        help="Directory to save generated certificates (default: output/pptx)",
        default=None
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Also convert generated PPTX certificates to PDF (saved in ouput_pdf/)"
    )

    args = parser.parse_args()
    run_batch_generation(
        file_path=args.file,
        template_path=args.template,
        output_dir=args.output,
        convert_to_pdf_flag=args.pdf
    )

if __name__ == "__main__":
    main()

import os
import sys
import shutil
import subprocess

def get_default_pdf_dir() -> str:
    """Determine output PDF directory."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    if os.path.exists("/files"):
        return "/files/ouput_pdf"
    return os.path.join(project_root, "ouput_pdf")

def convert_pptx_to_pdf(pptx_path: str, output_pdf_dir: str = None) -> str:
    """
    Converts a single PPTX file to PDF.
    - If LibreOffice is installed (standard in Linux/Docker): uses libreoffice --headless
    - On Windows: uses Microsoft PowerPoint COM automation
    """
    abs_pptx = os.path.abspath(pptx_path)
    if not os.path.exists(abs_pptx):
        raise FileNotFoundError(f"PPTX file not found: {abs_pptx}")

    output_pdf_dir = output_pdf_dir or get_default_pdf_dir()
    os.makedirs(output_pdf_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(abs_pptx))[0]
    abs_pdf = os.path.abspath(os.path.join(output_pdf_dir, f"{base_name}.pdf"))

    # 1. Check for LibreOffice (standard in Linux / n8n Docker containers)
    soffice_cmd = shutil.which("libreoffice") or shutil.which("soffice")
    if soffice_cmd:
        cmd = [soffice_cmd, "--headless", "--convert-to", "pdf", abs_pptx, "--outdir", output_pdf_dir]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode == 0 and os.path.exists(abs_pdf):
            return abs_pdf
        else:
            raise RuntimeError(f"LibreOffice conversion failed: {res.stderr}")

    # 2. Windows fallback: Use installed Microsoft PowerPoint via PowerShell COM
    if sys.platform == "win32":
        ps_script = f"""
$ppt = New-Object -ComObject PowerPoint.Application
try {{
    $deck = $ppt.Presentations.Open('{abs_pptx}', 0, 0, 0)
    $deck.SaveAs('{abs_pdf}', 32)
    $deck.Close()
}} finally {{
    $ppt.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
}}
"""
        res = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script], capture_output=True, text=True)
        if res.returncode == 0 and os.path.exists(abs_pdf):
            return abs_pdf
        else:
            raise RuntimeError(f"PowerPoint COM conversion failed: {res.stderr}")

    raise RuntimeError(
        "No PDF converter available. Please install LibreOffice:\n"
        "  - On Ubuntu/Debian/Docker: apt-get update && apt-get install -y libreoffice\n"
        "  - On Windows: Ensure PowerPoint or LibreOffice is installed."
    )

def convert_all_in_folder(input_dir: str = None, output_pdf_dir: str = None):
    """Convert all PPTX files in a directory to PDF."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)

    if input_dir is None:
        if os.path.exists("/files/output/pptx"):
            input_dir = "/files/output/pptx"
        else:
            input_dir = os.path.join(project_root, "output", "pptx")

    output_pdf_dir = output_pdf_dir or get_default_pdf_dir()
    os.makedirs(output_pdf_dir, exist_ok=True)

    pptx_files = [
        os.path.join(input_dir, f)
        for f in os.listdir(input_dir)
        if f.lower().endswith(".pptx") and not f.startswith("~$")
    ]

    if not pptx_files:
        print(f"No PPTX files found in: {input_dir}")
        return

    print(f"Found {len(pptx_files)} PPTX file(s) to convert...")
    for idx, pptx_file in enumerate(pptx_files, 1):
        try:
            pdf_path = convert_pptx_to_pdf(pptx_file, output_pdf_dir)
            print(f"[{idx}/{len(pptx_files)}] Converted: {os.path.basename(pdf_path)}")
        except Exception as e:
            print(f"[{idx}/{len(pptx_files)}] Failed {os.path.basename(pptx_file)}: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
        if os.path.isdir(target):
            convert_all_in_folder(target)
        else:
            out = convert_pptx_to_pdf(target)
            print("Converted:", out)
    else:
        convert_all_in_folder()

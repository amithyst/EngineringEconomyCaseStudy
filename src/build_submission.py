#!/usr/bin/env python3

import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.analysis.case_model import main as build_model_outputs
from src.excel.build_step_by_step_workbook import main as build_workbook
from src.reporting.generate_report_figures import main as build_figures


REPO_ROOT = Path(__file__).resolve().parents[1]
ANALYSIS_DIR = REPO_ROOT / "docs" / "analysis"
SUBMISSION_DIR = REPO_ROOT / "outputs" / "submission"


def run_latex() -> None:
    for _ in range(2):
        subprocess.run(
            ["xelatex", "-interaction=nonstopmode", "-halt-on-error", "final-report.tex"],
            cwd=ANALYSIS_DIR,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )


def package_submission() -> None:
    SUBMISSION_DIR.mkdir(parents=True, exist_ok=True)
    report_pdf = ANALYSIS_DIR / "final-report.pdf"
    excel_file = SUBMISSION_DIR / "EE2025_Case_Study_Calculations.xlsx"
    target_pdf = SUBMISSION_DIR / "EE2025_Case_Study_Report.pdf"
    zip_path = SUBMISSION_DIR / "EE2025_Case_Study_Submission.zip"
    target_pdf.write_bytes(report_pdf.read_bytes())
    with ZipFile(zip_path, "w") as zf:
        zf.write(target_pdf, arcname=target_pdf.name)
        zf.write(excel_file, arcname=excel_file.name)


def main() -> None:
    build_model_outputs()
    build_figures()
    build_workbook()
    run_latex()
    package_submission()


if __name__ == "__main__":
    main()

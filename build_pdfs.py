#!/usr/bin/env python3
"""Render the HTML CVs to PDF.

Engine by platform: WeasyPrint on Linux, headless Chromium (Puppeteer) on
Windows, where WeasyPrint's GTK/Pango libraries are not available.
Override with --engine weasyprint|chromium.

Usage: python build_pdfs.py [--engine ENGINE] [html_file ...]
"""
import argparse
import json
import pathlib
import platform
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent

TARGETS = {
    "cv_final.html": "Md_Niaj_Shahriar_Shishir_Odoo_FullStack_Architect.pdf",
    "cv_fullstack_odoo_developer.html": "Md_Niaj_Shahriar_Shishir_FullStack_Odoo_Developer.pdf",
    "cv_senior_project_manager_data_ltd.html": "Md_Niaj_Shahriar_Shishir_Senior_Project_Manager.pdf",
    "cv_tech_lead_integrations.html": "Md_Niaj_Shahriar_Shishir_Tech_Lead_Odoo_Customization_Integrations.pdf",
}
EXPECTED_PAGES = 2


def build_weasyprint(jobs):
    from weasyprint import HTML
    for html, pdf in jobs:
        HTML(filename=str(ROOT / html)).write_pdf(str(ROOT / pdf))


def build_chromium(jobs):
    payload = json.dumps([[str(ROOT / h), str(ROOT / p)] for h, p in jobs])
    subprocess.run(["node", str(ROOT / "build_pdfs_chromium.mjs"), payload], check=True, cwd=ROOT)


def page_count(pdf):
    try:
        from pypdf import PdfReader
    except ImportError:
        return None
    return len(PdfReader(str(ROOT / pdf)).pages)


def main():
    default_engine = "chromium" if platform.system() == "Windows" else "weasyprint"
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--engine", choices=["weasyprint", "chromium"], default=default_engine)
    parser.add_argument("files", nargs="*", help="HTML sources to build (default: all)")
    args = parser.parse_args()

    unknown = [f for f in args.files if f not in TARGETS]
    if unknown:
        sys.exit(f"Unknown source(s): {', '.join(unknown)}. Known: {', '.join(TARGETS)}")
    jobs = [(h, TARGETS[h]) for h in (args.files or TARGETS)]

    (build_chromium if args.engine == "chromium" else build_weasyprint)(jobs)

    status = 0
    for html, pdf in jobs:
        pages = page_count(pdf)
        flag = "" if pages in (None, EXPECTED_PAGES) else f"  <-- expected {EXPECTED_PAGES} pages"
        print(f"[{args.engine}] {html} -> {pdf} ({pages if pages is not None else '?'} pages){flag}")
        if flag:
            status = 1
    sys.exit(status)


if __name__ == "__main__":
    main()

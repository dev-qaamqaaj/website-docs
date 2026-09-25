#!/usr/bin/env python3
"""Build the OBD client onboarding kit.

    python3 build.py            # render every document + deck into dist/
    python3 build.py --no-deck  # documents only

Reads config.json, renders templates/*.html to build/*.html, prints them to PDF
with headless Chromium, then builds the kickoff deck (PPTX + PDF).
"""
import json
import shutil
import subprocess
import sys
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape

ROOT = Path(__file__).resolve().parent
BUILD = ROOT / "build"
DIST = ROOT / "dist"

DOCUMENTS = [
    ("01_welcome_pack.html", "01 - Welcome Pack.pdf"),
    ("02_client_onboarding_form.html", "02 - Onboarding Form.pdf"),
    ("03_service_agreement.html", "03 - Service Agreement.pdf"),
    ("04_handover.html", "04 - Go-Live & Handover.pdf"),
]

CHROME_CANDIDATES = [
    "/opt/pw-browsers/chromium-1194/chrome-linux/chrome",
    "chromium", "chromium-browser", "google-chrome", "google-chrome-stable",
]


def money(value, currency="PKR"):
    return f"{currency} {value:,.0f}"


def find_chrome():
    for c in CHROME_CANDIDATES:
        path = shutil.which(c) or (c if Path(c).exists() else None)
        if path:
            return path
    sys.exit("Chromium/Chrome not found — install one or add it to CHROME_CANDIDATES.")


def render_documents(cfg):
    env = Environment(loader=FileSystemLoader(ROOT / "templates"), autoescape=select_autoescape(["html"]))
    env.filters["money"] = lambda v: money(v, cfg["pricing"]["currency"])
    BUILD.mkdir(exist_ok=True)
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir()
    chrome = find_chrome()
    ctx = {"c": cfg}
    for template, pdf_name in DOCUMENTS:
        html_path = BUILD / template
        html_path.write_text(env.get_template(template).render(**ctx), encoding="utf-8")
        out = DIST / pdf_name
        subprocess.run(
            [chrome, "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
             "--allow-file-access-from-files", "--virtual-time-budget=4000",
             f"--print-to-pdf={out}", html_path.as_uri()],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        print(f"  ✓ {out.relative_to(ROOT)}")


def build_deck(cfg):
    sys.path.insert(0, str(ROOT / "deck"))
    from build_deck import build  # noqa: E402

    pptx = build(cfg, DIST / "00 - Kickoff Deck.pptx")
    print(f"  ✓ {pptx.relative_to(ROOT)}")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(DIST), str(pptx)],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print(f"  ✓ {pptx.with_suffix('.pdf').relative_to(ROOT)}")
    else:
        print("  ! LibreOffice not found — skipped deck PDF export")


def package(cfg):
    """Collect every document as PDF and PowerPoint in one client-named folder."""
    folder = ROOT / f"{cfg['client']['business']} Onboarding Kit"
    if folder.exists():
        shutil.rmtree(folder)
    pdf_dir, ppt_dir = folder / "PDF", folder / "PPT"
    pdf_dir.mkdir(parents=True)
    ppt_dir.mkdir()
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    for src in sorted(DIST.iterdir()):
        if src.suffix == ".pdf":
            shutil.copy2(src, pdf_dir / src.name)
            if soffice and not (DIST / (src.stem + ".pptx")).exists():
                subprocess.run([soffice, "--headless", "--infilter=impress_pdf_import", "--convert-to", "pptx",
                                "--outdir", str(ppt_dir), str(src)],
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        elif src.suffix == ".pptx":
            shutil.copy2(src, ppt_dir / src.name)
    print(f"  ✓ {folder.relative_to(ROOT)}/ (PDF + PPT)")


if __name__ == "__main__":
    config = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
    print("Building OBD onboarding kit…")
    render_documents(config)
    if "--no-deck" not in sys.argv:
        build_deck(config)
    package(config)
    print("Done → dist/ and the client kit folder")

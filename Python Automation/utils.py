import re
import shutil
from pathlib import Path

from config import PROCESSED_FOLDER


def extract(pattern, text):

    match = re.search(pattern, text, re.IGNORECASE)

    if match:
        return match.group(1).strip()

    return ""


def move_pdf(pdf_path: Path, inquiry_no: str):

    inquiry_no = inquiry_no.strip()

    new_name = f"{inquiry_no}.pdf"

    destination = PROCESSED_FOLDER / new_name

    shutil.move(str(pdf_path), str(destination))

    print("✓ PDF moved.")
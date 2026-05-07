import fitz
import re

def read_pdf_text(pdf_path: str) -> str:
    doc = fitz.open(pdf_path)
    parts = []

    for page in doc:
        text = page.get_text("text")
        if text:
            parts.append(text)

    doc.close()
    return "\n".join(parts)

def extract_contract_no(self, text: str) -> str:
    match = re.search(r"Số[:\s]*([0-9]+[-–][0-9]+/?HĐKT[-\sA-Z]*)", text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def extract_customer_name(self, text: str) -> str:
    match = re.search(r"Bên A.*?:\s*([^,]+)", text, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extract_install_address(self, text: str) -> str:
    match = re.search(r"lắp đặt tại[:\s]*(.+?)Điều", text, re.IGNORECASE)
    return match.group(1).strip() if match else ""

def extract_motor_power(self, text: str) -> str:
    match = re.search(r"(\d+(?:[.,]\d+)?\s*KW)", text, re.IGNORECASE)
    return match.group(1).strip() if match else ""
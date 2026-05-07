# import re
# from models import ContractData
# from pdf_reader import read_pdf_text


# def normalize_text(text: str) -> str:
#     text = text.replace("\xa0", " ")
#     text = re.sub(r"[ \t]+", " ", text)
#     text = re.sub(r"\n+", "\n", text)
#     return text.strip()

# def get_section(text: str, start: str, end: str = None) -> str:
#     start = re.escape(start)
#     end = re.escape(end) if end else None

#     if end:
#         pattern = rf"{start}(.+?){end}"
#     else:
#         pattern = rf"{start}(.+)$"

#     match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
#     return match.group(1).strip() if match else ""


# class ContractExtractor:
  
     
#     def extract(self, pdf_path: str) -> ContractData:
#         raw_text = read_pdf_text(pdf_path)
#         text = normalize_text(raw_text)

   

#         header = text.split("Điều 01")[0]
#         dieu_01 = get_section(text, "Điều 01", "Điều 02")
#         dieu_04 = get_section(text, "Điều 04", "Điều 05")

#         return ContractData(
#             source_file=pdf_path,
#             contract_no=self.extract_contract_no(header),
#             customer_name=self.extract_customer_name(header),
#             installation_address=self.extract_installation_address(dieu_01),
#             product_days=self.extract_product_days(dieu_04),
#             installation_days=self.extract_installation_days(dieu_04),
#             raw_text=text,
#         )
#     def extract_contract_no(self, text: str) -> str:
#         match = re.search(r"Số\s*:\s*(.+)", text, re.IGNORECASE)
#         return match.group(1).strip() if match else ""
    
#     def extract_customer_name(self, text: str) -> str:
#         match = re.search(
#             r"Bên\s*A.*?:\s*(?:Ông|Bà)?\s*[-:]*\s*([^\n]+)",
#             text,
#             re.IGNORECASE
#         )
#         return match.group(1).strip() if match else ""
    
#     def extract_installation_address(self, text: str) -> str:
#         for line in text.splitlines():
#             if "lắp đặt tại" in line.lower():
#                 line = re.sub(r".*lắp đặt tại\s*:\s*", "", line, flags=re.IGNORECASE)
#                 return line.strip(" .:-")
#         return ""
#     def extract_product_days(self, text: str) -> int:
#         match = re.search(r"Nhập thiết bị.*?(\d+)\s*ngày", text, re.IGNORECASE | re.DOTALL)
#         return int(match.group(1)) if match else 0


#     def extract_installation_days(self, text: str) -> int:
#         match = re.search(r"Vận chuyển.*?(\d+)\s*ngày", text, re.IGNORECASE | re.DOTALL)
#         return int(match.group(1)) if match else 0


import re
from models import ContractData
from pdf_reader import read_pdf_text_native, read_pdf_ocr


def normalize_text(text: str) -> str:
    if not text:
        return ""

    text = text.replace("\xa0", " ")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def find_header(text: str) -> str:
    """
    Try to cut before section 1, but tolerate OCR noise.
    If not found, just keep the first chunk of the document.
    """
    patterns = [
        r"^(.*?)(?=\b(?:Điều|Dieu|DIEU)\s*0?1\b)",
        r"^(.*?)(?=\n\s*1\s*['`.:]\s*1\s*[:.])",   # OCR-style "1 ' 1 :"
        r"^(.*?)(?=\n\s*1\s*[.:])",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()

    return text[:3000].strip()


def get_section(text: str, start_num: int, end_num: int = None) -> str:
    """
    Best-effort only. OCR may destroy section titles.
    """
    dieu_word = r"(?:Điều|Dieu|DIEU)"
    start_patterns = [
        rf"\b{dieu_word}\s*0?{start_num}\b",
        rf"\n\s*{start_num}\s*['`.:]\s*{start_num}\s*[:.]",
        rf"\n\s*{start_num}\s*[.:]",
    ]

    if end_num is not None:
        end_patterns = [
            rf"\b{dieu_word}\s*0?{end_num}\b",
            rf"\n\s*{end_num}\s*['`.:]\s*{end_num}\s*[:.]",
            rf"\n\s*{end_num}\s*[.:]",
        ]
    else:
        end_patterns = []

    for start_pattern in start_patterns:
        if end_num is not None:
            for end_pattern in end_patterns:
                pattern = rf"{start_pattern}(.*?){end_pattern}"
                match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
                if match:
                    return match.group(1).strip()
        else:
            pattern = rf"{start_pattern}(.*)$"
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1).strip()

    return ""


class ContractExtractor:
    def extract(self, pdf_path: str) -> ContractData:
        native_raw = read_pdf_text_native(pdf_path)
        native_text = normalize_text(native_raw)
        native_data = self._extract_from_text(pdf_path, native_text)
        native_score = self._score_result(native_data)

        print(f"[Parser] Native score: {native_score}")

        if native_score >= 3:
            print("[Parser] Using native parsed result")
            return native_data

        print("[Parser] Native parse weak -> fallback to OCR")
        ocr_raw = read_pdf_ocr(pdf_path, lang="vie+eng")
        ocr_text = normalize_text(ocr_raw)
        ocr_data = self._extract_from_text(pdf_path, ocr_text)
        ocr_score = self._score_result(ocr_data)

        print(f"[Parser] OCR score: {ocr_score}")

        if ocr_score > native_score:
            print("[Parser] Using OCR parsed result")
            return ocr_data

        print("[Parser] OCR not better -> keep native result")
        return native_data

    def _extract_from_text(self, pdf_path: str, text: str) -> ContractData:
        header = find_header(text)
        dieu_01 = get_section(text, 1, 2)
        dieu_04 = get_section(text, 4, 5)

        address_source = dieu_01 if dieu_01 else text
        days_source = dieu_04 if dieu_04 else text

        customer_name = self.extract_customer_name(header)
        if not customer_name:
            customer_name = self.extract_customer_name(text)

        return ContractData(
            source_file=pdf_path,
            contract_no=clean_contract_no(self.extract_contract_no(header)),
            customer_name=clean_customer_name(customer_name),
            installation_address=clean_address(self.extract_installation_address(address_source)),
            product_days=self.extract_product_days(days_source),
            installation_days=self.extract_installation_days(days_source),
            raw_text=text,
        )

    def _score_result(self, data: ContractData) -> int:
        score = 0
        if data.contract_no:
            score += 1
        if data.customer_name:
            score += 1
        if data.installation_address:
            score += 1
        if data.product_days > 0:
            score += 1
        if data.installation_days > 0:
            score += 1
        return score

    def extract_contract_no(self, text: str) -> str:
        patterns = [
            r"(?:Số|So|S)\s*[:\-5]?\s*([0-9]{2,4}\s*-\s*[0-9]{2,4}[^\s\n,]*)",
            r"([0-9]{2,4}\s*-\s*[0-9]{2,4}\s*[A-Z0-9\/\-.]*)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value = match.group(1).strip(" .:-,")
                value = re.sub(r"\s+", "", value)
                # reject phone-number-like junk
                if len(re.sub(r"\D", "", value)) >= 8 and "-" not in value:
                    continue
                return value
        return ""

    def extract_customer_name(self, text: str) -> str:
        patterns = [
            # Standard / semi-clean forms
            r"(?:Bên|Ben|Be)\s*A[^\n:]{0,40}[:\-]\s*(?:Ông|Bà|Anh|Chị|Ong|Ba|Bdr)?\s*[-:]*\s*([^\n]+)",

            # OCR-noisy "Bên A ... name" on same line
            r"(?:Bên|Ben|Be)\s*A[^\n]{0,80}?([A-ZÀ-Ỹ][A-ZÀ-Ỹ\s\.'\-]{5,})",

            # Alternate phrase
            r"(?:Chủ\s*đầu\s*tư|Chu\s*dau\s*tu)[^\n:]{0,20}[:\-]\s*([^\n]+)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value = match.group(1).strip(" .:-,'\"")
                value = re.sub(r"\(.*?\)", "", value).strip()
                value = re.sub(r"\s{2,}", " ", value)

                # Remove common prefixes accidentally captured
                value = re.sub(r"^(Ông|Bà|Anh|Chị|Ong|Ba|Bdr)\s+", "", value, flags=re.IGNORECASE)

                # Keep only likely name-like chunk before extra junk
                value = re.split(r"\b(?:Địa chỉ|Dia chi|CMND|CCCD|MST|Điện thoại|Dien thoai)\b", value, maxsplit=1, flags=re.IGNORECASE)[0].strip()

                if len(value) >= 4:
                    return value

        return ""
    def clean_customer_name(value: str) -> str:
        if not value:
            return ""

        value = value.strip(" .:-,'\"")
        value = re.sub(r"\s{2,}", " ", value)

        replacements = {
            "Ttfl": "THỊ",
            "TtIi": "THỊ",
            "NHU'": "NHƯ",
            "XU,|N": "XUÂN",    
            "XUAN": "XUÂN",
            "LE ": "LÊ ",
        }

        for wrong, right in replacements.items():
            value = value.replace(wrong, right)

        return value
    def extract_installation_address(self, text: str) -> str:
        patterns = [
            r"lắp đặt tại\s*[:.]?\s*([^\n]+)",
            r"được lắp đặt tại\s*[:.]?\s*([^\n]+)",
            r"thi công tại\s*[:.]?\s*([^\n]+)",
            r"tại\s*:\s*([^\n]+Q[u]?ận\s*\d+[^\n]*)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value = match.group(1).strip(" .:-")
                if len(value) > 5:
                    return value

        for line in text.splitlines():
            lower_line = line.lower()
            if "lắp đặt tại" in lower_line or "thi công tại" in lower_line:
                cleaned = re.sub(r".*(lắp đặt tại|thi công tại)\s*[:.]?\s*", "", line, flags=re.IGNORECASE)
                cleaned = cleaned.strip(" .:-")
                if cleaned:
                    return cleaned

        return ""

    def extract_product_days(self, text: str) -> int:
        patterns = [
            r"nhập thiết bị.*?(\d+)\s*ngày",
            r"sản xuất.*?(\d+)\s*ngày",
            r"thiết bị.*?được sản xuất.*?(\d+)\s*ngày",
            r"(\d+)\s*ngày[^\n]{0,100}sản xuất",
            r"(\d+)\s*ngày[^\n]{0,100}thiết bị",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_installation_days(self, text: str) -> int:
        patterns = [
            r"vận chuyển.*?lắp đặt.*?vận hành.*?(\d+)\s*ngày",
            r"vận chuyển.*?lắp đặt.*?(\d+)\s*ngày",
            r"lắp đặt.*?vận hành.*?(\d+)\s*ngày",
            r"(\d+)\s*ngày[^\n]{0,120}lắp đặt",
            r"(\d+)\s*ngày[^\n]{0,120}vận chuyển",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0
import re
from models import ContractData
from pdf_reader import read_pdf_text


def normalize_text(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.strip()


def get_section(text: str, start: str, end: str = None) -> str:
    start = re.escape(start)
    end = re.escape(end) if end else None

    if end:
        pattern = rf"{start}(.+?){end}"
    else:
        pattern = rf"{start}(.+)$"

    match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
    return match.group(1).strip() if match else ""


class ContractExtractor:
    def extract(self, pdf_path: str) -> ContractData:
        raw_text = read_pdf_text(pdf_path)
        text = normalize_text(raw_text)

        header = text.split("Điều 01")[0]
        dieu_01 = get_section(text, "Điều 01", "Điều 02")
        dieu_04 = get_section(text, "Điều 04", "Điều 05")

        return ContractData(
            source_file=pdf_path,
            contract_no=self.extract_contract_no(header),
            customer_name=self.extract_customer_name(header),
            customer_address=self.extract_customer_address(header),
            installation_address=self.extract_installation_address(dieu_01),
            product_days=self.extract_product_days(dieu_04),
            installation_days=self.extract_installation_days(dieu_04),
            load_capacity_kg=self.extract_load_capacity_kg(text),
            speed_mpm=self.extract_speed_mpm(text),
            motor_brand=self.extract_motor_brand(text),
            motor_power=self.extract_motor_power(text),
            control_system_brand=self.extract_control_system_brand(text),
            number_of_stops=self.extract_number_of_stops(text),
            number_of_floors_served=self.extract_number_of_floors_served(text),
            electrical_box_brand=self.extract_electrical_box_brand(text),
            pit_depth_mm=self.extract_pit_depth_mm(text),
            overhead_mm=self.extract_overhead_mm(text),
            steel_frame=self.extract_steel_frame(text),
            shaft_wall_material=self.extract_shaft_wall_material(text),
            cabin_door_material=self.extract_cabin_door_material(text),
            landing_door_material=self.extract_landing_door_material(text),
            raw_text=text,
        )

    def extract_contract_no(self, text: str) -> str:
        match = re.search(r"Số\s*:\s*(.+)", text, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    def extract_customer_name(self, text: str) -> str:
        patterns = [
            r"Bên\s*A.*?:\s*(?:Ông|Bà)?\s*[-:]*\s*([^\n]+)",
            r"Bên\s*A\s*\(.*?\)\s*:\s*([^\n]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip(" .:-")
        return ""

    def extract_customer_address(self, text: str) -> str:
        patterns = [
            r"Bên\s*A.*?Địa\s*chỉ\s*:\s*([^\n]+)",
            r"Địa\s*chỉ\s*:\s*([^\n]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                value = match.group(1).strip(" .:-")
                if len(value) > 5:
                    return value
        return ""

    def extract_installation_address(self, text: str) -> str:
        patterns = [
            r"lắp đặt tại\s*:\s*([^\n]+)",
            r"lắp đặt tại\s*([^\n]+)",
            r"được lắp đặt tại\s*:\s*([^\n]+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip(" .:-")
        return ""

    def extract_product_days(self, text: str) -> int:
        patterns = [
            r"Nhập thiết bị.*?(\d+)\s*ngày",
            r"Thời gian .*?sản xuất.*?(\d+)\s*ngày",
            r"sản xuất.*?(\d+)\s*ngày",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_installation_days(self, text: str) -> int:
        patterns = [
            r"Vận chuyển.*?(\d+)\s*ngày",
            r"lắp đặt.*?(\d+)\s*ngày",
            r"vận hành.*?(\d+)\s*ngày",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_load_capacity_kg(self, text: str) -> int:
        patterns = [
            r"Tải\s*trọng.*?(\d{2,4})\s*kg",
            r"load.*?(\d{2,4})\s*kg",
            r"(\d{2,4})\s*kg.*?Tải\s*trọng",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_speed_mpm(self, text: str) -> int:
        patterns = [
            r"Tốc\s*độ.*?(\d+)\s*(?:m/phút|m/phut|m/phút|m/p|m/min)",
            r"speed.*?(\d+)\s*(?:m/phút|m/p|m/min)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_motor_brand(self, text: str) -> str:
        brands = [
            "MONTANARI",
            "SICOR",
            "TORIN",
            "ZIEHL-ABEGG",
            "HITACHI",
            "FUJI",
        ]

        upper_text = text.upper()
        motor_region_patterns = [
            r"ĐỘNG\s*CƠ.*?(MONTANARI|SICOR|TORIN|HITACHI|FUJI|ZIEHL[- ]ABEGG)",
            r"MÁY\s*KÉO.*?(MONTANARI|SICOR|TORIN|HITACHI|FUJI|ZIEHL[- ]ABEGG)",
        ]

        for pattern in motor_region_patterns:
            match = re.search(pattern, upper_text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1).replace(" ", "-").upper()

        for brand in brands:
            if brand in upper_text:
                return brand

        return ""

    def extract_control_system_brand(self, text: str) -> str:
        patterns = [
            r"(?:điều khiển|VVVF|inverter).*?(YASKAWA|STEP|MONARCH|FUJI|MITSUBISHI)",
            r"(YASKAWA|STEP|MONARCH|FUJI|MITSUBISHI).*?(?:điều khiển|VVVF|inverter)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1).upper()
        return ""

    def extract_number_of_stops(self, text: str) -> int:
        patterns = [
            r"Số\s*điểm\s*dừng.*?(\d+)",
            r"điểm\s*dừng.*?(\d+)",
            r"stops.*?(\d+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_number_of_floors_served(self, text: str) -> int:
        patterns = [
            r"Số\s*tầng\s*phục\s*vụ.*?(\d+)",
            r"tầng\s*phục\s*vụ.*?(\d+)",
            r"floors\s*served.*?(\d+)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def extract_electrical_box_brand(self, text: str) -> str:
        patterns = [
            r"(?:tủ điện|control cabinet|electrical box).*?(YASKAWA|STEP|MONARCH|FUJI|MITSUBISHI|DELTA|SCHNEIDER)",
            r"(YASKAWA|STEP|MONARCH|FUJI|MITSUBISHI|DELTA|SCHNEIDER).*?(?:tủ điện|control cabinet|electrical box)",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1).upper()

        if self.extract_control_system_brand(text):
            return self.extract_control_system_brand(text)

        return ""

    def extract_motor_power(self, text: str) -> str:
        patterns = [
            r"Công\s*suất\s*[:\-]?\s*(\d+(?:[.,]\d+)?\s*K\s*W)",
            r"Công\s*suất.*?(\d+(?:[.,]\d+)?\s*K\s*W)",
            r"(\d+(?:[.,]\d+)?\s*K\s*W)",
        ]

        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return re.sub(r"\s+", " ", match.group(1)).strip()

        return ""

    def extract_pit_depth_mm(self, text: str) -> int:
        patterns = [
            r"Pit\s*\([^\n]*?\)\s*(\d{3,5})\s*mm",
            r"Pit[^\n]{0,40}?(\d{3,5})\s*mm",
            r"chiều\s*âm\s*hố\s*thang[^\n]{0,40}?(\d{3,5})\s*mm",
        ]
        return self._extract_first_int(text, patterns)

    def extract_overhead_mm(self, text: str) -> int:
        patterns = [
            r"OH\s*\(Overhead\)\s*(\d{3,5})\s*mm",
            r"OH[^\n]{0,40}?(\d{3,5})\s*mm",
            r"Overhead[^\n]{0,40}?(\d{3,5})\s*mm",
        ]
        return self._extract_first_int(text, patterns)

    def extract_steel_frame(self, text: str) -> bool:
        text_lower = text.lower()

        steel_keywords = [
            "khung hố thép",
            "khung thép",
            "sắt chấn",
            "khung sắt",
            "steel frame",
        ]
        concrete_keywords = [
            "hố bê tông",
            "bê tông cốt thép",
            "btct",
        ]

        if any(keyword in text_lower for keyword in steel_keywords):
            return True
        if any(keyword in text_lower for keyword in concrete_keywords):
            return False
        return False

    def extract_shaft_wall_material(self, text: str) -> str:
        block = self._extract_block_after_anchor(
            text,
            anchor_patterns=[
                r"Vách kính lắp đặt quanh khung hố",
                r"Vách\s*kính\s*lắp\s*đặt\s*quanh\s*khung\s*hố",
            ],
            stop_patterns=[
                r"\n\s*05\s+Pit",
                r"\n\s*VI\.\s*THIẾT\s*KẾ\s*CABIN",
                r"\n\s*VII\.\s*CỬA\s*CABIN",
            ],
        )
        return self._summarize_material_block(block)

    def extract_cabin_door_material(self, text: str) -> str:
        block = self._extract_block_after_anchor(
            text,
            anchor_patterns=[r"Vật liệu cửa cabin"],
            stop_patterns=[r"\n\s*VIII\.\s*CỬA\s*TẦNG", r"\n\s*01\s+Loại cửa"],
        )
        return self._summarize_material_block(block)

    def extract_landing_door_material(self, text: str) -> str:
        block = self._extract_block_after_anchor(
            text,
            anchor_patterns=[r"Vật liệu cửa các tầng"],
            stop_patterns=[r"\n\s*03\s+Bao che các tầng", r"\n\s*IX\.\s*THIẾT\s*BỊ\s*AN\s*TOÀN"],
        )
        return self._summarize_material_block(block)

    def _extract_first_int(self, text: str, patterns: list[str]) -> int:
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return int(match.group(1))
        return 0

    def _extract_block_after_anchor(
        self,
        text: str,
        anchor_patterns: list[str],
        stop_patterns: list[str],
    ) -> str:
        for anchor_pattern in anchor_patterns:
            anchor_match = re.search(anchor_pattern, text, re.IGNORECASE)
            if not anchor_match:
                continue

            start = anchor_match.end()
            tail = text[start:]
            stop_index = len(tail)
            for stop_pattern in stop_patterns:
                stop_match = re.search(stop_pattern, tail, re.IGNORECASE)
                if stop_match:
                    stop_index = min(stop_index, stop_match.start())
            block = tail[:stop_index].strip()
            if block:
                return block
        return ""

    def _summarize_material_block(self, block: str) -> str:
        if not block:
            return ""

        lines = [line.strip(" -•\t") for line in block.splitlines() if line.strip()]
        cleaned_lines: list[str] = []
        for line in lines:
            line = re.sub(r"\s+", " ", line)
            if re.fullmatch(r"\d{1,2}", line):
                continue
            cleaned_lines.append(line)

        if not cleaned_lines:
            return ""

        material_keywords = (
            "inox",
            "kính",
            "nhôm",
            "khung",
            "cường lực",
            "sơn tĩnh điện",
            "10mm",
            "1mm",
        )
        selected = [line for line in cleaned_lines[:4] if any(keyword in line.lower() for keyword in material_keywords)]
        if not selected:
            selected = cleaned_lines[:2]

        summary = " | ".join(selected)
        summary = re.sub(r"\s+", " ", summary).strip(" |")
        return summary[:300]

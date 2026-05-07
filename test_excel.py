from extractor import ContractExtractor
from planner import generate_plan
from excel_repo import ExcelRepository

pdf_path = r"C:\Users\admin\Desktop\Capstone\072-26 HĐKT-Ông NGUYỄN TRỌNG BÌNH.pdf"
excel_path = r"C:\Users\admin\Desktop\Capstone\contracts.xlsx"

extractor = ContractExtractor()
data = extractor.extract(pdf_path)

plan = generate_plan(data.product_days, data.installation_days)

repo = ExcelRepository(excel_path)
repo.save_contract_and_plan(data, plan)

print("Saved to Excel successfully.")
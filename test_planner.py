from extractor import ContractExtractor
from planner import generate_plan

pdf_path = r"C:\Users\admin\Desktop\Capstone\072-26 HĐKT-Ông NGUYỄN TRỌNG BÌNH.pdf"

extractor = ContractExtractor()
data = extractor.extract(pdf_path)

plan = generate_plan(data.product_days, data.installation_days)

print("Start:", plan.start_date.date())
print("Product End:", plan.product_end.date())
print("Installation End:", plan.installation_end.date())
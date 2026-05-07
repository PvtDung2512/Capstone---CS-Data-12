from extractor import ContractExtractor

pdf_path = r"C:\Users\admin\Desktop\Capstone\072-26 HĐKT-Ông NGUYỄN TRỌNG BÌNH.pdf"

extractor = ContractExtractor()
data = extractor.extract(pdf_path)


print("\n=== EXTRACTED ===")
print("Contract No:", repr(data.contract_no))
print("Customer:", repr(data.customer_name))
print("Customer Address:", repr(data.customer_address))
print("Installation Address:", repr(data.installation_address))
print("Product Days:", data.product_days)
print("Installation Days:", data.installation_days)
print("Load Capacity (kg):", data.load_capacity_kg)
print("Speed (m/min):", data.speed_mpm)
print("Motor Brand:", repr(data.motor_brand))
print("Control System Brand:", repr(data.control_system_brand))
print("Number of Stops:", data.number_of_stops)
print("Number of Floors Served:", data.number_of_floors_served)
print("Electrical Box Brand:", repr(data.electrical_box_brand))
print("Motor Power", repr(data.motor_power))
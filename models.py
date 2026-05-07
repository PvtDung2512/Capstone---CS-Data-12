from dataclasses import dataclass


@dataclass
class ContractData:
    source_file: str = ""
    contract_no: str = ""
    customer_name: str = ""
    customer_address: str = ""
    installation_address: str = ""

    product_days: int = 0
    installation_days: int = 0

    load_capacity_kg: int = 0
    speed_mpm: int = 0
    motor_brand: str = ""
    motor_power: str = ""
    control_system_brand: str = ""
    number_of_stops: int = 0
    number_of_floors_served: int = 0
    electrical_box_brand: str = ""

    pit_depth_mm: int = 0
    overhead_mm: int = 0
    steel_frame: bool = False
    shaft_wall_material: str = ""
    cabin_door_material: str = ""
    landing_door_material: str = ""

    raw_text: str = ""

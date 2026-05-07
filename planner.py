from dataclasses import dataclass
from datetime import datetime, timedelta


@dataclass
class PlanResult:
    start_date: datetime
    product_start: datetime
    product_end: datetime
    installation_start: datetime
    installation_end: datetime


def generate_plan(product_days: int, installation_days: int) -> PlanResult:
    start_date = datetime.today()

    product_start = start_date
    product_end = product_start + timedelta(days=product_days)

    installation_start = product_end
    installation_end = installation_start + timedelta(days=installation_days)

    return PlanResult(
        start_date=start_date,
        product_start=product_start,
        product_end=product_end,
        installation_start=installation_start,
        installation_end=installation_end,
    )
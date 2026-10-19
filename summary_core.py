from collections import defaultdict
from datetime import date, datetime

import openpyxl


def read_data(filepath):
    workbook = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    try:
        data = []
        for row in workbook.active.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            redeemed_at = row[0]
            product_type = str(row[1]) if row[1] else ""
            if isinstance(redeemed_at, datetime):
                day = redeemed_at.date()
            elif isinstance(redeemed_at, date):
                day = redeemed_at
            else:
                day = datetime.strptime(str(redeemed_at)[:10], "%Y-%m-%d").date()
            data.append((
                day,
                "代金券" in product_type,
                float(row[2] or 0),
                float(row[3] or 0),
                float(row[4] or 0),
            ))
        return data
    finally:
        workbook.close()


def aggregate(data):
    result = defaultdict(lambda: [0, 0, 0])
    for day, is_coupon, received, service_fee, merchant_amount in data:
        totals = result[(day, is_coupon)]
        totals[0] += received
        totals[1] += service_fee
        totals[2] += merchant_amount
    return result


def collect_dates(*aggregates):
    days = sorted({key[0] for values in aggregates for key in values})
    if not days:
        raise ValueError("输入文件没有可汇总的有效数据")
    return days

import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook

from summary_core import aggregate, collect_dates, read_data


class SummaryCoreTests(unittest.TestCase):
    def test_reads_and_aggregates_products_and_coupons(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["核销时间", "商品类型", "订单实收", "服务费", "商家应得"])
            sheet.append([datetime(2026, 10, 1, 10), "普通商品", 100, 5, 95])
            sheet.append([date(2026, 10, 1), "代金券", 200, 10, 190])
            workbook.save(path)
            workbook.close()

            result = aggregate(read_data(path))
            self.assertEqual(result[(date(2026, 10, 1), False)], [100, 5, 95])
            self.assertEqual(result[(date(2026, 10, 1), True)], [200, 10, 190])
            self.assertEqual(collect_dates(result), [date(2026, 10, 1)])

    def test_empty_inputs_are_rejected_before_report_generation(self):
        with self.assertRaisesRegex(ValueError, "没有可汇总"):
            collect_dates({}, {})


if __name__ == "__main__":
    unittest.main()

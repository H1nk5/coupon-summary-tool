import argparse
import json

from summary_core import generate_report


def main(argv=None):
    parser = argparse.ArgumentParser(description="按日汇总兰考与民权商品、代金券数据")
    parser.add_argument("--lankao", required=True, help="兰考 Excel 文件")
    parser.add_argument("--minquan", required=True, help="民权 Excel 文件")
    parser.add_argument("--output", required=True, help="输出 XLSX 文件")
    args = parser.parse_args(argv)
    result = generate_report(args.lankao, args.minquan, args.output)
    print(json.dumps({
        **result,
        "first_date": str(result["first_date"]),
        "last_date": str(result["last_date"]),
    }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

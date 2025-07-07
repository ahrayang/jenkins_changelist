
import util
import argparse
import logging
from datetime import datetime, timedelta


def main():
    logging.basicConfig(level=logging.debug, format="%(asctime)s - %(levelname)s - %(message)s")
    parser = argparse.ArgumentParser()
    parser.add_argument("--since", type=str, required=False)
    parser.add_argument("--until", type=str, required=False)
    parser.add_argument("--depot", type=str, default='//Sol/Dev1/...') #dev1과 dev1next를 시트 분리하여 보여줘야 할지?
    args = parser.parse_args()
    if not args.since or not args.until:
        now = datetime.now()
        daterange = now - timedelta(days=10)
        args.since = daterange.strftime("%Y/%m/%d:%H:%M:%S")
        args.until = now.strftime("%Y/%m/%d:%H:%M:%S")
    changes = util.get_changes(args.depot, args.since, args.until)
    collected_data = []
    for change in changes:
        rows = util.parse_describe_grouped(change)
        collected_data.extend(rows)
    if collected_data:
        util.append_to_excel_with_hyperlink(collected_data)
    else:
        logging.debug("추가할 데이터가 없습니다.")

if __name__ == "__main__":
    main()
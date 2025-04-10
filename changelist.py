import argparse
import subprocess
import re
import pandas as pd
import os
import logging
from datetime import datetime, timedelta

def utc_to_kst(utc_str):
    try:
        dt_utc = datetime.strptime(utc_str, "%Y/%m/%d %H:%M:%S")
        dt_kst = dt_utc + timedelta(hours=9)
        return dt_kst.strftime("%Y/%m/%d %H:%M:%S")
    except Exception as e:
        logging.error(f"utc -> kst 변환 실패: {e}")
        return utc_str

def run_command(cmd):
    logging.info(f"[CMD 실행]: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True, check=True)
        logging.info(f"[CMD 결과]:\n{result.stdout}")
        return result.stdout
    except subprocess.CalledProcessError as e:
        logging.error(f"[CMD 실행 실패]: {e.stderr}")
        return ""

def get_changes(depot, since, until):
    cmd = f'p4 changes {depot}@{since},{until}'
    logging.info(f"cmd 전송: {cmd}")
    output = run_command(cmd)
    changes = []
    pattern = re.compile(r"^Change\s+(\d+)\s+on\s+(\d{4}/\d{2}/\d{2}(?:\s+\d{2}:\d{2}:\d{2})?)\s+by\s+(\S+)")
    for line in output.splitlines():
        line = line.strip()
        m = re.search(pattern, line)
        if m:
            change_num = m.group(1)
            changes.append(change_num)
    return changes

def parse_describe(change_num):
    cmd = f"p4 describe -s {change_num}"
    output = run_command(cmd)
    if not output:
        return None
    info = {"change": change_num, "submit_time": "", "user": "", "description": "", "jira_url": "", "diff_summary": ""}
    lines = output.splitlines()
    header_pattern = re.compile(r"^Change\s+\d+\s+by\s+(\S+)@.*\s+on\s+(\d{4}/\d{2}/\d{2}(?:\s+\d{2}:\d{2}:\d{2})?)")
    header_found = False
    description_lines = []
    affected_files = []
    in_affected = False
    for line in lines:
        if not header_found:
            m = re.search(header_pattern, line.strip())
            if m:
                info["user"] = m.group(1)
                info["submit_time"] = m.group(2)
                header_found = True
                if " " not in info["submit_time"]:
                    info["submit_time"] += " 00:00:00"
            continue
        if not in_affected:
            if "Affected files" in line:
                in_affected = True
                continue
            else:
                description_lines.append(line.strip())
        else:
            if line.strip() == "":
                continue
            m = re.match(r"^\s*\.\.\.\s+(//\S+)#\d+\s+(\w+)", line)
            if m:
                depot_path = m.group(1)
                action = m.group(2)
                folder, file_name = os.path.split(depot_path)
                _, ext = os.path.splitext(file_name)
                file_info = f"File: {file_name}, Action: {action}, Type: {ext if ext else 'N/A'}, In Folder: {folder}"
                affected_files.append(file_info)
    info["description"] = "\n".join(description_lines).strip()
    jira_match = re.search(r'(https?://\S+(?:atlassian\.net\S+))', info["description"])
    info["jira_url"] = jira_match.group(1) if jira_match else ""
    info["diff_summary"] = "; ".join(affected_files)
    return info

def append_to_excel(data_list, excel_file="build_history.xlsx"):
    columns = ["Change 번호", "서밋 시점 (KST)", "작업자", "설명", "Jira URL", "Diff 내용"]
    df_new = pd.DataFrame(data_list, columns=columns)
    if os.path.exists(excel_file):
        try:
            df_existing = pd.read_excel(excel_file)
            df_total = pd.concat([df_existing, df_new], ignore_index=True)
        except Exception as e:
            logging.error("Excel 파일 읽기 오류: " + str(e))
            df_total = df_new
    else:
        df_total = df_new
    try:
        df_total.to_excel(excel_file, index=False)
        logging.info(f"Data saved to {excel_file}")
    except Exception as e:
        logging.error("Excel 쓰기 오류: " + str(e))

def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
    parser = argparse.ArgumentParser()
    parser.add_argument("--since", type=str, required=False)
    parser.add_argument("--until", type=str, required=False)
    parser.add_argument("--depot", type=str, default='//Sol/Dev1Next/...') # 저장소 경로로
    args = parser.parse_args()
    if not args.since or not args.until:
        now = datetime.now()
        daterange = now - timedelta(days=7)
        args.since = daterange.strftime("%Y/%m/%d:%H:%M:%S")
        args.until = now.strftime("%Y/%m/%d:%H:%M:%S")
    changes = get_changes(args.depot, args.since, args.until)
    collected_data = []
    for change in changes:
        info = parse_describe(change)
        if info:
            info["submit_time"] = utc_to_kst(info["submit_time"])
            collected_data.append({
                "Change 번호": info["change"],
                "서밋 시점 (KST)": info["submit_time"],
                "작업자": info["user"],
                "설명": info["description"],
                "Jira URL": info["jira_url"],
                "Diff 내용": info["diff_summary"]
            })
    if collected_data:
        append_to_excel(collected_data)
    else:
        logging.info("추가할 데이터가 없습니다.")

if __name__ == "__main__":
    main()

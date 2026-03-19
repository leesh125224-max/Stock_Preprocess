import os
import glob
import re
import win32com.client as win32
import csv
import time
from datetime import datetime
from collections import defaultdict

# --- 설정 (Configuration) ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# DB 저장 대신 CSV 저장 폴더 (사용자 요청 "경로 똑같이")
DB_DIR = r"c:\Users\leesh\.gemini\뉴스보고서\특징주\data\db"
DATA_DIR = r"c:\Users\leesh\.gemini\뉴스보고서\특징주\data\특징주"
CSV_DIR = r"c:\Users\leesh\.gemini\뉴스보고서\특징주\data\특징주_csv"

# 최종 결과 파일
RESULT_CSV_PATH = os.path.join(DB_DIR, "특징주.csv")

def parse_date_from_filename(file_path):
    basename = os.path.basename(file_path)
    # path components to find year folder
    path_parts = file_path.split(os.sep)
    
    year = datetime.now().year # Default
    
    if "25년" in path_parts:
        year = 2025
    elif "26년" in path_parts:
        year = 2026
        
    # 예: 1.20특징주.hwp -> 1, 20
    match = re.search(r"(\d+)\.(\d+)", basename)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        return f"{year}-{month:02d}-{day:02d}"
        
    return None

def convert_hwp_to_csv(hwp, hwp_path):
    basename = os.path.basename(hwp_path)
    # Get parent folder name to handle "25년", "26년" distinction
    parent_dir = os.path.basename(os.path.dirname(hwp_path))
    
    # Prefix csv name with parent dir to avoid collision (e.g. 25년_1.20.csv vs 26년_1.20.csv)
    csv_filename = f"{parent_dir}_{basename.replace('.hwp', '.csv')}"
    csv_path = os.path.join(CSV_DIR, csv_filename)
    
    try:
        if not hwp.Open(hwp_path, "HWP", "forceopen:true"):
            print(f"Failed to open {basename}")
            return None
        
        if hwp.SaveAs(csv_path, "CSV"):
            return csv_path
        else:
            print(f"SaveAs CSV failed for {basename}")
            return None
    except Exception as e:
        print(f"Error converting {basename}: {e}")
        return None

# ... (parse_csv_and_extract_grouped_data is mostly same, but depends on date_str passing)

def parse_csv_items(csv_path):
    """
    CSV 파일에서 유효한 (종목명, 등락률, 이유, 날짜) 데이터를 추출하여 리스트로 반환.
    """
    raw_items = []
    
    try:
        with open(csv_path, 'r', encoding='cp949') as f:
            reader = csv.reader(f)
            rows = list(reader)
    except Exception:
        try:
             with open(csv_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                rows = list(reader)
        except Exception as e:
            print(f"  CSV Read Error {os.path.basename(csv_path)}: {e}")
            return []

    state = 0
    last_valid_fluc = 0.0 # 직전 등락률 저장용 (초기값 0.0 또는 None)

    for row in rows:
        clean_row = [c.strip() for c in row if c.strip()]
        if not clean_row: continue
        row_str = " ".join(clean_row)
        
        # 1. 섹션 구분
        if "종목명" in row_str:
            if state == 1: 
                break
            state = 1
            # 헤더가 나오면 등락률 초기화 (새로운 표 시작)
            last_valid_fluc = 0.0 
            continue 

        # 2. 종료 조건
        if "시간외" in row_str and "종목명" not in row_str:
            break
            
        # 3. 데이터 추출
        if state == 1:
            if len(row) < 4: continue
            
            stock_name = row[0].strip()
            fluc_str = row[1].strip()
            reason = row[-1].strip()
            
            # 등락률 처리 (빈 경우 직전 값 사용)
            # 등락률 처리 (빈 경우 직전 값 사용)
            if not fluc_str:
                fluc_val = last_valid_fluc
            else:
                try:
                    fluc_val = float(fluc_str.replace('%', ''))
                    
                    # 사용자 요청: 30 초과 시 값 보정 (자릿수 오류 수정)
                    # 2982 -> 29.82 (4자리), 113 -> 11.3 (3자리)
                    if fluc_val > 30:
                        if fluc_val >= 1000: # 4자리 이상 (예: 2982 -> 29.82)
                            fluc_val /= 100
                        elif fluc_val >= 100: # 3자리 (예: 113 -> 11.3)
                            fluc_val /= 10
                            
                    last_valid_fluc = fluc_val # 유효값 갱신
                except:
                    # 파싱 실패 시 직전 값 사용 or 0.0
                    fluc_val = last_valid_fluc
            
            # 유효 범위 필터링 (0 < x <= 30)
            # 보정 후에도 범위 밖이면 제외 (혹은 그냥 저장? '0초과 30이하여야돼'는 필터링 조건으로 해석)
            if fluc_val <= 0 or fluc_val > 30:
                continue
            
            if not stock_name or "종목명" in stock_name: continue
            if len(stock_name) > 50: continue

            # 사용자 요청: "상장 첫날" 포함 시 제외
            if "상장 첫날" in reason: continue
            
            # 사용자 요청: 이유가 비어있으면 제외
            if not reason: continue

            raw_items.append({
                "stock_name": stock_name,
                "fluctuation": fluc_val,
                "reason": reason
            })
            
    return raw_items

def main():
    # 1. 준비
    os.makedirs(DB_DIR, exist_ok=True)
    os.makedirs(CSV_DIR, exist_ok=True)
    
    hwp_files = glob.glob(os.path.join(DATA_DIR, "**", "*.hwp"), recursive=True)
    print(f"Found {len(hwp_files)} HWP files.")
    
    if not hwp_files:
        print("No HWP files found.")
        return

    # 2. HWP -> CSV 변환
    print("--- Converting HWP to CSV ---")
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = True
        
        converted_csvs = []
        for hwp_file in hwp_files:
            csv_path = convert_hwp_to_csv(hwp, hwp_file)
            if csv_path:
                converted_csvs.append((csv_path, hwp_file))
                
        hwp.Quit()
    except Exception as e:
        print(f"HWP Automation Error: {e}")
        return

    print(f"Converted {len(converted_csvs)} files. Extracting and Aggregating Globally...")
    
    # 3. 데이터 추출 및 전역 통합 (이유 기준)
    global_reason_stats = defaultdict(list)
    
    for csv_path, hwp_file in converted_csvs:
        # 날짜 파싱은 파일명 처리를 위해 호출만 하고, 실제 데이터에는 포함하지 않음
        _ = parse_date_from_filename(hwp_file)
        
        items = parse_csv_items(csv_path)
        for item in items:
            reason = item['reason']
            # 전체 아이템 저장 (종목명 포함)
            global_reason_stats[reason].append(item)
            
        print(f"  {os.path.basename(hwp_file)} -> {len(items)} items processed.")

    # 4. 최종 데이터 가공 (Max/Min 계산)
    final_data = []
    
    for reason, items in global_reason_stats.items():
        if not items: continue
        
        flucs = [i['fluctuation'] for i in items]
        max_fluc = max(flucs)
        # 데이터가 2개 이상일 때만 min 표시 (하나면 공백)
        min_fluc = min(flucs) if len(flucs) > 1 else None
        
        # 종목명 통합 (중복 제거 후 정렬)
        unique_stocks = sorted(list(set([i['stock_name'] for i in items])))
        stock_names_str = ", ".join(unique_stocks)
        
        final_data.append({
            "종목명": stock_names_str,
            "등락률_최대": f"{max_fluc:.2f}",
            "등락률_최소": f"{min_fluc:.2f}" if min_fluc is not None else "",
            "이유": reason
        })
    
    # 정렬: 등락률 높은 순
    final_data.sort(key=lambda x: float(x["등락률_최대"]), reverse=True)

    # 저장 (utf-8-sig for Excel)
    if final_data:
        try:
            with open(RESULT_CSV_PATH, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=["종목명", "등락률_최대", "등락률_최소", "이유"])
                writer.writeheader()
                writer.writerows(final_data)
            print(f"\nAll Done. Total {len(final_data)} unique reasons saved to {RESULT_CSV_PATH}")
        except Exception as e:
            print(f"Error saving merged CSV: {e}")
    else:
        print("No data extracted.")

if __name__ == "__main__":
    main()

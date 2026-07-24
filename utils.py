import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
import os
import glob
from datetime import datetime, timedelta
import re
import copy

# ----------------------------------------
# 1. 변수설정  및 환경설정 관리
# - 공통으로 사용할 고정값(변수)들을 정의하는 구역입니다.

# 탐지할 파란색 RGB 코드 정의 (엑셀 표준 파란색 및 테마색 대응)
TARGET_BLUE_CODES = ["FF00B0F0", "0000B0F0"]
TARGET_BLUE_SET = frozenset[str](TARGET_BLUE_CODES)

# 데이터 변경사항 구분을 위해 H열(인덱스)에 칠할 색상 정의
COLOR_ADD = '66FF33'    # 연두색: 부서에서 새로 '추가'한 행 표시
COLOR_DELETE = 'FF6666' # 연한 빨간색: 부서에서 '삭제'한 행 표시

# 엑셀 시트 내의 주요 열(Column) 위치 번호 (A=1, B=2, ..., H=8)
INDEX_COL = 8   # H열: 데이터의 고유 번호(인덱스)가 적히는 기준 열
NAME_COL = 9    # I열: 사업명 또는 부서명이 적히는 열
J_COL = 10      # J열: 수입/지출/잔액 구분(세로 병합 대상)
K_COL = 11      # K열: 수입/지출/잔액 라벨(행별)
L_COL = 12      # L열: 항목명(용역비 등) — DEPT DIFF 시작 열

CATEGORY_LABELS = frozenset({'수입', '지출', '잔액', '발생이자'})

# I~L열: EXCEL LAYOUT (병합·서식·템플릿 값)
LAYOUT_COL_START = 9
LAYOUT_COL_END = L_COL

# L열(12)~: DEPT DIFF (항목명·숫자·파란 수정)
DEPT_DATA_COL_START = L_COL

# L열~: 부서 파란색 수정 감지 (H열 인덱스는 제외)
BLUE_SCAN_COL_START = DEPT_DATA_COL_START

# 서식(테두리, 배경색)을 모니터링할 최대 열 제한 (1번째 열부터 60번째 열까지)
STYLE_SNAPSHOT_MAX_COL = 60

# (구) I열 병합 확장 상한 — extend_section_i_name_merge_to_end(max_gap=None) 기본 무제한
MAX_I_MERGE_EXTEND_GAP = 5

# ----------------------------------------
# 2. 날짜 및 파일 매칭 
# - 현재 컴퓨터 날짜를 읽어서 자동으로 지난달(전월) 기준 파일과 시트 이름을 계산합니다.

# 1단계: 오늘 날짜 기준으로 이번 달의 1일을 구합니다.
first_day_of_this_month = datetime.today().replace(day=1)
# 2단계: 이번 달 1일에서 하루를 빼서 '지난달의 마지막 날'을 구합니다. (자동으로 전월 계산 완료)
last_day_of_prev_month = first_day_of_this_month - timedelta(days=1)

# 3단계: 지난달의 년도(뒤의 2자리, 예: '26')와 월(2자리 숫자, 예: '05')을 추출합니다.
prev_year_short = last_day_of_prev_month.strftime('%y')
prev_month = last_day_of_prev_month.month

# 4단계: 자동으로 취합 대상 시트 이름을 목록으로 만듭니다.
numeric_sheets = [
    f'취합({prev_year_short}{prev_month:02d})-원안',
    f"취합({prev_year_short}{prev_month:02d})-수정1안",
    f"취합({prev_year_short}{prev_month:02d})-수정2안",
]

text_sheets = ["전월대비 비교", f'{prev_year_short}-{int(prev_year_short)+1} 지출계획']

# 고정된 폴더 및 파일 경로 설정
base_path = r'C:\Users\user\Desktop\인턴\3)현금흐름'    ####=====!!!!!!!! 추후 수정해야할 필요있음!!!!!!!======####
file_name = f'도시개발본부 현금흐름 분석자료({prev_year_short}.{prev_month:02d}.기준)_부서명.xlsx'
file_path = os.path.join(base_path, '취합용파일')
master_file = os.path.join(file_path, file_name)
target_folder = os.path.join(base_path, '부서별파일')

# 각 부서 코드(H열 입력값)와 취합용 파일상의 부서명(I열) 매칭 맵
dept_index_codes = {
    "도시개발본부": "UDH", "도시기획부": "UP", "개발사업1부": "DB1", "개발사업2부": "DB2",
    "관광사업": "TB", "도시개발1부": "UD1", "도시개발2부": "UD2", "도시사업1부": "UB1",
    "도시사업2부": "UB2", "법인세외": "CT"
}

PREFIX_I_LABELS = {
    "UDH": "도시개발본부 행정관리비", "UP": "도시기획부 행정관리비", "DB1": "개발사업1부 행정관리비",
    "DB2": "개발사업2부 행정관리비", "TB": "관광사업부 행정관리비", "UD1": "도시개발1부 행정관리비",
    "UD2": "도시개발2부 행정관리비", "UB1": "도시사업1부 행정관리비", "UB2": "도시사업2부 행정관리비",
    "CT": "법인세 외",
}

def dept_file_matches_prev_month(filepath, year_short, month, numeric_sheet_names):
    """
    [함수 설명] 부서에서 제출한 파일이 '지난달' 파일이 맞는지 검사합니다.
    [원리] 파일 이름에 '26.05' 같은 날짜가 있거나, 파일 내부에 '취합(2605)-원안' 같은 시트가 있는지 확인합니다.
    """
    base = os.path.basename(filepath) # 전체 경로에서 파일 이름만 추출. 
    # 대소문자 구분 없이 다양한 날짜 형식(26.05, 2605, 26-05)의 패턴을 정의합니다.
    month_patterns = [
        rf'{re.escape(year_short)}\.{month:02d}',
        rf'{re.escape(year_short)}{month:02d}',
        rf'{re.escape(year_short)}-{month:02d}',
        rf'{re.escape(year_short)}\.{month}(?!\d)',
    ]
    # 패턴 중 하나라도 파일 이름에 매칭되면 올바른 파일로 판명합니다.
    if any(re.search(p, base, re.IGNORECASE) for p in month_patterns):
        return True
    try:
        # 파일 이름에 날짜가 없다면, 엑셀을 살짝 열어서 내부에 해당 전월 취합 시트가 있는지 확인합니다.
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=False)
        names = set(wb.sheetnames)
        wb.close()
        return any(sn in names for sn in numeric_sheet_names)
    except Exception:
        return False

# ----------------------------------------
# 3. 엑셀 셀/서식/병합 조작 도구
# - 엑셀의 테두리, 배경색, 병합 상태 등을 복사하고 판별하는 도구들입니다.
##   : 행 추가/삭제 시, 취합용 파일의 셀 서식이 깨지기 때문에, 취합 전 셀 서식을 복사하고 저장 -> 부서별 파일들을 취합할 때, 저장된 셀 서식과 비교하여 셀 서식을 최종 취합본에 붙여넣기 합니다.

def has_any_border(cell):
    """셀에 상, 하, 좌, 우 테두리 선이 하나라도 그려져 있는지 여부를 참/거짓으로 반환합니다."""
    if not cell.border:
        return False
    for side in ('left', 'right', 'top', 'bottom'):
        s = getattr(cell.border, side)
        if s and s.style: # 선의 스타일(실선, 점선 등)이 존재한다면 테두리가 있는 것입니다.
            return True
    return False

def copy_border_deep(border):
    """엑셀의 테두리 서식은 내부적으로 얽혀있어 완전히 독립된 사본(Deep Copy)을 만들어 복사해야 깨지지 않습니다."""
    if not border:
        return None
    sides = {}
    for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
        s = getattr(border, side)
        sides[side] = copy.copy(s) if s else None # 각 방향의 선을 하나씩 똑같이 복제합니다.
    return openpyxl.styles.Border(
        **sides,
        diagonal_direction=border.diagonal_direction,
        outline=border.outline,
        vertical=copy.copy(border.vertical) if border.vertical else None,
        horizontal=copy.copy(border.horizontal) if border.horizontal else None,
    )

def is_blue_cell(cell):
    """셀의 배경색이 미리 지정해 둔 '취합 대상 파란색(TARGET_BLUE_CODES)'인지 판별합니다."""
    fill = cell.fill
    if not fill or fill.fill_type is None:
        return False
    rgb = fill.start_color.rgb
    if rgb is None:
        return False
    return str(rgb).upper() in TARGET_BLUE_SET


def row_has_blue(ws, row_num, max_col, col_start=None):
    """L열~에서 파란색 수정 셀이 있는지 확인 (H열 인덱스는 제외)."""
    start = BLUE_SCAN_COL_START if col_start is None else col_start
    for c in range(start, max_col + 1):
        if is_blue_cell(ws.cell(row=row_num, column=c)):
            return True
    return False

def get_real_cell(ws, cell, merge_lookup=None):
    """
    병합셀일 경우 병합범위 내 가장 상단 셀 획득.
    merge_lookup: {(row, col): (병합_좌상단_행, 병합_좌상단_열)} 사전이 있으면 O(1) 조회.
    """
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        key = (cell.row, cell.column)
        if merge_lookup is not None and key in merge_lookup:
            ar, ac = merge_lookup[key]
            return ws.cell(row=ar, column=ac)
        for merged_range in ws.merged_cells.ranges:
            if cell.coordinate in merged_range:
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell


def build_merge_lookup(ws):
    '''병합 껍데기 셀 → 병합 좌상단(값·서식이 있는 셀) 좌표 lookup (get_real_cell 가속).'''
    lookup = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row == mr.max_row and mr.min_col == mr.max_col:
            continue
        anchor = (mr.min_row, mr.min_col)
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != anchor:
                    lookup[(r, c)] = anchor
    return lookup

def load_workbooks(file_list):
    """여러 개의 파일 경로 목록을 받아서, 엑셀 파일들을 한 번에 다 열어 하나의 리스트 주머니로 반환합니다."""
    workbooks = []
    for f in file_list:
        try:
            wb = openpyxl.load_workbook(
                f, data_only=False, keep_links=False, keep_vba=False
            )
            workbooks.append(wb)
        except Exception as e:
            print(f'파일 로드 실패: {f} / {e}')
    return workbooks


def load_dept_workbooks(dept_files):
    '''부서 파일 경로 → workbook dict (시트별 반복 로드 방지).'''
    dept_wbs = {}
    for f in dept_files:
        try:
            dept_wbs[f] = openpyxl.load_workbook(
                f, data_only=False, keep_links=False, keep_vba=False
            )
        except Exception as e:
            print(f'부서별 파일 로드 실패: {f} / {e}')
    return dept_wbs

def copy_cell_style_safe(src_cell, tgt_cell):
    """원본 셀(src_cell)의 글자체, 테두리, 배경색, 정렬 등을 대상 셀(tgt_cell)로 안전하게 에러 없이 복사합니다."""
    if src_cell.has_style:
        try:
            tgt_cell.font = copy.copy(src_cell.font)
            tgt_cell.border = copy.copy(src_cell.border)
            tgt_cell.fill = copy.copy(src_cell.fill)
            tgt_cell.number_format = src_cell.number_format
            tgt_cell.protection = copy.copy(src_cell.protection)
            tgt_cell.alignment = copy.copy(src_cell.alignment)
        except Exception:
            pass

def is_layout_col(col):
    """I~L EXCEL LAYOUT 구간 (병합·서식·템플릿 값)."""
    return LAYOUT_COL_START <= col <= LAYOUT_COL_END


def is_dept_data_col(col):
    """L열~ DEPT DIFF 구간 (항목명·숫자·파란 수정)."""
    return col >= DEPT_DATA_COL_START


def is_layout_only_col(col):
    """I~K — layout 전용 (DEPT DIFF L~ 구간과 불겹)."""
    return is_layout_col(col) and not is_dept_data_col(col)

def should_capture_layout_cell(cell):
    """셀에 값, 테두리, 배경색 중 하나라도 존재해서 백업(스냅샷)을 떠야 하는 대상인지 판별합니다."""
    if cell.value is not None: return True
    if has_any_border(cell): return True
    if cell.fill and cell.fill.fill_type: return True
    return False

def capture_cell_style(cell):
    """나중에 행이 삽입되면서 서식이 깨질 것을 대비해, 셀의 서식(디자인)만 따로 사진을 찍듯 저장(스냅샷)해 둡니다."""
    if not cell.has_style and not has_any_border(cell):
        if not (cell.fill and cell.fill.fill_type):
            return None
    return {
        'fill': copy.copy(cell.fill) if cell.fill else None,
        'font': copy.copy(cell.font) if cell.font else None,
        'alignment': copy.copy(cell.alignment) if cell.alignment else None,
        'border': copy_border_deep(cell.border),
        'number_format': cell.number_format,
    }

def apply_cell_style(cell, style, apply_value=True, value=None, apply_fill=True, ws=None):
    """백업해 두었던 서식(style) 사진을 바탕으로 특정 셀에 디자인과 값을 다시 그대로 입혀줍니다."""
    if ws is not None:
        cell = get_real_cell(ws, cell)
    elif isinstance(cell, openpyxl.cell.cell.MergedCell):
        return
    if apply_value and value is not None:
        cell.value = value
    if not style:
        return
    if apply_fill and style.get('fill'): cell.fill = style['fill']
    if style.get('font'): cell.font = style['font']
    if style.get('alignment'): cell.alignment = style['alignment']
    if style.get('border'): cell.border = style['border']
    if style.get('number_format'): cell.number_format = style['number_format']

def safe_unmerge_cells(ws, min_row, min_col, max_row, max_col):
    """엑셀 행을 밀고 당길 때 프로그램이 충돌하여 다운되는 현상(KeyError)을 방지하면서 안전하게 셀 병합을 해제합니다."""
    anchor = ws.cell(row=min_row, column=min_col)
    saved_value = anchor.value # 병합이 깨지면서 글자가 날아가지 않도록 임시 저장

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c)
    try:
        ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
    except KeyError:
        for merged in list(ws.merged_cells.ranges):
            if (merged.min_row == min_row and merged.min_col == min_col and
                    merged.max_row == max_row and merged.max_col == max_col):
                ws.merged_cells.remove(merged)
                break
    if saved_value is not None:
        ws.cell(row=min_row, column=min_col).value = saved_value

# -------------------------------------
# 4. 인덱스 및 섹션 매칭 도구
# - H열의 고유 코드(예: DB2-1)를 분석하여 데이터의 소속과 순서를 제어합니다.

def get_valid_prefixes():
    """환경설정에 등록된 유효한 부서 코드 집합(예: 'UDH', 'DB1' 등)을 반환."""
    return set(dept_index_codes.values())


def is_allowed_index_prefix(prefix, master_map, valid_prefixes):
    '''
    신규/수정 인덱스 처리 허용 prefix.
    dept_index_codes(DB2 등) 또는 취합용에 이미 존재하는 prefix(U, Q 등) 허용.
    '''
    if not prefix:
        return False
    if prefix in valid_prefixes:
        return True
    needle = f'{prefix}-'
    return any(idx_str.startswith(needle) for idx_str in master_map)

def extract_prefix(idx_str, pattern):
    """'DB2-5' 라는 인덱스 글자에서 부서 기호인 'DB2'만 쏙 뽑아냅니다 => prefix. 실패하면 None."""
    match = pattern.match(idx_str)
    return match.group(1) if match else None

def get_prefix_i_label(prefix):
    """'DB2' 같은 코드를 주면 '개발사업2부 행정관리비' 같은 엑셀용 정식 명칭 텍스트를 매칭해 줍니다."""
    if prefix in PREFIX_I_LABELS:
        return PREFIX_I_LABELS[prefix]
    for dept_name, code in dept_index_codes.items():
        if code == prefix:
            return f'{dept_name} 행정관리비'
    return None

def is_i_label_value(value):
    """I열에 부서명 텍스트가 와야 하는데, 실수로 수식(=Q40)이나 숫자가 들어가 오염되었는지 검사합니다."""
    if value is None: return False
    if isinstance(value, (int, float)): return True
    if isinstance(value, str):
        s = value.strip()
        if s.startswith('=') or s.isdigit(): return True
        try:
            float(s); return True
        except ValueError:
            return False
    return False

def apply_i_section_label(ws, start_row, prefix, merge_lookup=None):
    """각 부서별 데이터가 시작되는 첫 줄의 I열 대표 칸에 정식 부서명 글자를 적어줍니다."""
    label = get_prefix_i_label(prefix)
    if not label:
        return
    cell = get_real_cell(ws, ws.cell(start_row, NAME_COL), merge_lookup)
    if isinstance(cell, openpyxl.cell.cell.MergedCell):
        return
    cell.value = label

def restore_all_i_section_labels(ws, valid_prefixes, merge_lookup=None):
    """프로그램이 행을 추가/변경하는 과정에서 지워지거나 꼬인 모든 부서명 글자(I열)들을 원본 기준으로 일괄 복구합니다."""
    for start_row, prefix in get_section_header_rows(ws, valid_prefixes):
        apply_i_section_label(ws, start_row, prefix, merge_lookup)

def get_section_header_rows(ws, valid_prefixes=None):
    """
    H열 단독 코드(하이픈 없음) = 섹션 헤더 행.
    DB2, U, UD, Q 등 valid_prefixes 등록 여부와 무관하게 모두 포함.
    """
    headers = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=INDEX_COL).value
        if v is None:
            continue
        s = str(v).strip()
        if not s or '-' in s:
            continue
        headers.append((r, s))
    return headers

def find_section_content_end_row(ws, start_row):
    """하나의 부서 섹션 내부에서 '실제 내용 데이터가 들어있는 진짜 마지막 줄'이 몇 번째 행인지 추적합니다."""
    last = start_row
    for r in range(start_row, ws.max_row + 1):
        h = ws.cell(r, INDEX_COL).value
        if h is not None and str(h).strip():
            last = r
            continue
        for c in range(LAYOUT_COL_START, LAYOUT_COL_END + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != '':
                last = r
                break
    return last

def get_section_end_row(ws, prefix, valid_prefixes, fallback=None):
    """특정 부서 섹션이 끝나고 다음 부서 섹션이 시작되기 바로 전의 행 번호를 계산합니다."""
    headers = get_section_header_rows(ws, valid_prefixes)
    for i, (row, p) in enumerate(headers):
        if p == prefix:
            if i + 1 < len(headers):
                return headers[i + 1][0] - 1 # 다음 부서 시작 행의 바로 윗줄이 끝 줄입니다.
            return find_section_content_end_row(ws, row)
    return fallback if fallback is not None else ws.max_row

def get_sheet_index_map(ws, merge_lookup=None, include_name=False):
    """
    시트 전체 인덱스(H열) → {row_num, name?} 주소록.
    include_name=False면 I열 조회 생략 (취합 diff용, 속도↑).
    """
    index_map = {}
    for r in range(1, ws.max_row + 1):
        idx_value = ws.cell(row=r, column=INDEX_COL).value
        if idx_value is None:
            continue
        idx_str = str(idx_value).strip()
        if not idx_str:
            continue
        entry = {'row_num': r}
        if include_name:
            name_cell = get_real_cell(ws, ws.cell(row=r, column=NAME_COL), merge_lookup)
            entry['name'] = name_cell.value
        index_map[idx_str] = entry
    return index_map


def shift_index_map_rows(index_map, from_row, delta=1):
    '''insert_rows() 후 from_row 이상 인덱스 행번호를 delta만큼 밀어 갱신 (전체 재스캔 방지).'''
    for info in index_map.values():
        if info['row_num'] >= from_row:
            info['row_num'] += delta


def shift_inserted_rows_set(inserted_rows, from_row, delta=1):
    '''
    [버그 수정용 헬퍼] insert_rows() 호출 직후, 이미 기록해 둔 '이전 삽입행' 번호들도
    이번 삽입으로 인해 같이 밀려난 만큼 보정합니다.

    shift_index_map_rows와 동일한 원리이지만 대상이 set(삽입행 모음)이라는 점만 다릅니다.
    한 섹션에 행을 여러 번 연속 추가할 때, 직전에 추가한 삽입행 번호가 그 다음 insert_rows()
    호출로 다시 한 칸 밀려나는 것을 정확히 추적하기 위해 필요합니다.
    '''
    return {r + delta if r >= from_row else r for r in inserted_rows}


def shift_merged_cells_after_insert(ws, insert_row, num_rows=1):
    """
    행이 삽입되었을 때, 삽입 지점 아래에 있는 실제 엑셀 병합 셀들의 주소를 num_rows만큼 밀어주는 함수
    """
    new_merged_cells = []
    for mcr in list(ws.merged_cells.ranges):
        # 1. 병합 구역 전체가 삽입 지점보다 아래에 있는 경우 -> 행 번호를 통째로 밀어줌
        if mcr.min_row >= insert_row:
            mcr.shift(row_shift=num_rows)
            new_merged_cells.append(mcr)
        # 2. 병합 구역 중간에 행이 삽입된 경우 -> 병합 영역의 max_row 높이를 늘려줌
        elif mcr.max_row >= insert_row:
            mcr.max_row += num_rows
            new_merged_cells.append(mcr)
        # 3. 삽입 지점보다 위에 있는 병합 구역은 그대로 유지
        else:
            new_merged_cells.append(mcr)

    # 기존 시트의 병합 정보를 깨끗이 비운 후, 리프레시된 병합 범위를 다시 등록
    ws.merged_cells.ranges = [] # 초기화 시킴
    for mcr in new_merged_cells:
        ws.merged_cells.add(mcr)


def master_insert_rows(ws, insert_at, amount, index_store, inserted_rows):
    '''insert_rows + 병합 보정 + IndexStore·inserted_rows 갱신을 한 번에 처리.'''
    if amount <= 0:
        return inserted_rows
    ws.insert_rows(insert_at, amount=amount)
    shift_merged_cells_after_insert(ws, insert_at, amount)
    shift_index_store_rows(index_store, insert_at, amount)
    inserted_rows = shift_inserted_rows_set(inserted_rows, insert_at, amount)
    for i in range(amount):
        inserted_rows.add(insert_at + i)
    return inserted_rows


def capture_row_snapshot(ws, row, max_col, merge_lookup=None, ws_values=None):
    '''인덱스 행의 열 값 스냅샷 (H열 제외). L열~ 수식은 ws_values에서 계산값 저장.'''
    values = {}
    for c in range(1, max_col + 1):
        if c == INDEX_COL:
            continue
        raw = ws.cell(row=row, column=c)
        cell = get_real_cell(ws, raw, merge_lookup)
        val = cell.value
        if is_dept_data_col(c) and ws_values is not None:
            cached = ws_values.cell(row, c).value
            if isinstance(val, str) and val.startswith('='):
                pass  # L열~ 수식은 문자열 유지 (insert 후 Translator 보정용)
            elif cached is not None:
                val = cached
        values[c] = val
    return values


# ---------------------------------------------------------------------------
# 통합 인덱스맵 (IndexStore)
#
# 취합 시작 시 {인덱스: {row_num, values}} 한 곳에 행번호+데이터를 저장합니다.
# master_map(인덱스→행번호)은 index_store_as_map()으로 얻는 뷰입니다.
#
# 복원 정책 (restore_index_store_data mode):
#   'data'   — L열~ (insert 직후 baseline 데이터 보호)
#   'layout' — I~L (레이아웃 템플릿 적용 직전)
#   'all'    — H 제외 전열 (특수 경우)
#
# H열 없는 행(잔액 헤더 등)은 section_row_baseline에 prefix별 rel_row로 별도 저장.
# ---------------------------------------------------------------------------

def build_index_store(ws, max_col, merge_lookup=None, valid_prefixes=None, ws_values=None):
    '''취합 시작 시점 통합 인덱스맵 + 섹션 비인덱스 행 baseline.'''
    index_map = get_sheet_index_map(ws, merge_lookup, include_name=False)
    store = {}
    for idx, info in index_map.items():
        row = info['row_num']
        store[idx] = {
            'row_num': row,
            'capture_row': row,
            'values': capture_row_snapshot(ws, row, max_col, merge_lookup, ws_values),
        }
    section_baseline = {}
    if valid_prefixes:
        section_baseline = build_section_row_baseline(
            ws, valid_prefixes, max_col, merge_lookup, ws_values=None,
        )
    return store, section_baseline


def build_baseline_map(ws, index_map, max_col, merge_lookup=None):
    '''하위 호환 — IndexStore의 인덱스→{row_num, values} 부분만 반환.'''
    baseline = {}
    for idx, info in index_map.items():
        row = info['row_num']
        baseline[idx] = {
            'row_num': row,
            'values': capture_row_snapshot(ws, row, max_col, merge_lookup),
        }
    return baseline


def index_store_as_map(index_store):
    '''IndexStore → master_map 뷰 {인덱스: {row_num}}.'''
    return {idx: {'row_num': rec['row_num']} for idx, rec in index_store.items()}


def shift_index_store_rows(index_store, from_row, delta=1):
    '''insert_rows() 후 IndexStore 내 row_num 갱신.'''
    for rec in index_store.values():
        if rec['row_num'] >= from_row:
            rec['row_num'] += delta


def build_section_row_baseline(ws, valid_prefixes, max_col, merge_lookup=None,
                               prefix_pattern=None, ws_values=None):
    '''섹션 내 H열 번호형 인덱스가 없는 행(잔액 헤더 등) rel_row→values baseline.'''
    if prefix_pattern is None:
        prefix_pattern = re.compile(r'^([A-Za-z0-9]+)-(\d+)$')
    result = {}
    headers = get_section_header_rows(ws, valid_prefixes)
    for i, (start_row, prefix) in enumerate(headers):
        if i + 1 < len(headers):
            end_row = headers[i + 1][0] - 1
        else:
            end_row = find_section_content_end_row(ws, start_row)
        rows = {}
        for r in range(start_row + 1, end_row + 1):
            h = ws.cell(row=r, column=INDEX_COL).value
            if h is not None and str(h).strip():
                if prefix_pattern.match(str(h).strip()):
                    continue
            rel = r - start_row
            rows[rel] = capture_row_snapshot(ws, r, max_col, merge_lookup, ws_values)
        if rows:
            result[prefix] = {'start_row': start_row, 'rel_end': end_row - start_row, 'rows': rows}
    return result


def restore_section_row_baseline(ws, section_baseline, valid_prefixes, merge_lookup,
                                   inserted_rows=None, data_cols_only=True):
    '''insert 후 섹션 비인덱스 행을 rel_row 기준으로 values 복원.'''
    if not section_baseline:
        return
    inserted_rows = inserted_rows or set()
    header_map = {p: row for row, p in get_section_header_rows(ws, valid_prefixes)}
    for prefix, base in section_baseline.items():
        if prefix not in header_map:
            continue
        start_row = header_map[prefix]
        rel_end = base['rel_end']
        end_row = get_section_end_row(ws, prefix, valid_prefixes, fallback=start_row + rel_end)
        end_row = max(start_row + rel_end, end_row)
        section_inserted = {r for r in inserted_rows if start_row <= r <= end_row}
        rel_to_abs = build_rel_row_to_abs_row_map(
            start_row, end_row, section_inserted, rel_end,
        ) if section_inserted else None
        for rel, values in base['rows'].items():
            if rel_to_abs is not None:
                abs_row = rel_to_abs.get(rel)
            else:
                abs_row = start_row + rel
            if abs_row is None:
                continue
            apply_baseline_row_data(
                ws, abs_row, values, merge_lookup, data_cols_only,
                origin_row=base['start_row'] + rel,
            )


def _is_formula(val):
    return isinstance(val, str) and val.startswith('=')


def _write_snapshot_value(cell, val, allow_formula=False):
    '''baseline 복원: insert 전 수식 문자열은 행 참조가 깨지므로 data 복원 시 제외.'''
    if not allow_formula and _is_formula(val):
        return
    cell.value = val


def restore_index_store_data(master_ws, index_store, merge_lookup, skip_indexes=None,
                             mode='data'):
    '''
    IndexStore values를 현재 row_num 위치에 복원.

    mode='data': L열~ (insert/blue copy 후 데이터 보호)
    mode='layout': I~L (레이아웃 템플릿 적용 직전)
    mode='all': H 제외 전열
    '''
    skip = skip_indexes or set()
    for idx, rec in index_store.items():
        if idx in skip:
            continue
        row = rec['row_num']
        for c, val in rec['values'].items():
            if mode == 'data' and not is_dept_data_col(c):
                continue
            if mode == 'layout' and not is_layout_col(c):
                continue
            raw = master_ws.cell(row=row, column=c)
            cell = get_real_cell(master_ws, raw, merge_lookup)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue
            if mode == 'data' and _is_formula(val):
                _write_snapshot_value(cell, val, allow_formula=True)
            else:
                allow_formula = mode != 'data'
                _write_snapshot_value(cell, val, allow_formula=allow_formula)


def apply_baseline_row_data(ws, row, values, merge_lookup=None, data_cols_only=True,
                            origin_row=None):
    '''baseline values를 지정 행에 복원 (기본: L열~ 데이터만, I~K layout 제외).'''
    for c, val in values.items():
        if data_cols_only and not is_dept_data_col(c):
            continue
        raw = ws.cell(row=row, column=c)
        cell = get_real_cell(ws, raw, merge_lookup)
        if isinstance(cell, openpyxl.cell.cell.MergedCell):
            continue
        allow_formula = _is_formula(val)
        _write_snapshot_value(cell, val, allow_formula=allow_formula)


def restore_baseline_data(master_ws, master_map, baseline_map, merge_lookup,
                          skip_indexes=None, data_cols_only=True):
    '''하위 호환 — restore_index_store_data(mode="data") 래퍼.'''
    restore_index_store_data(
        master_ws, baseline_map, merge_lookup, skip_indexes,
        mode='data' if data_cols_only else 'all',
    )


# ---------------------------------------------------------------------------
# 취합 3층 아키텍처
#
#   1. MAP        — IndexStore: insert·행번호 갱신·밀린 행 baseline 복원
#   2. DEPT DIFF  — 부서 대비: 파란셀·remap·신규행·삭제 표시
#   3. EXCEL LAYOUT — 병합/서식: 템플릿·J열 구분·연두색 H
#
# main.py는 PLAN(부서 스캔) → MAP → DEPT DIFF → EXCEL LAYOUT → DEPT 삭제 순으로 호출.
# ---------------------------------------------------------------------------

def map_execute_row_inserts(master_ws, index_store, new_rows_buffer, master_map,
                            prefix_first_rows, max_col, inserted_rows):
    '''
    MAP: 버퍼된 신규 인덱스 행을 insert_rows하고 IndexStore·inserted_rows 갱신.
    삽입 위치 = 바로 위 인덱스(예: U-4)의 row_num + 1.
    Returns (inserted_rows, added_index_values, master_map, master_merge_lookup).
    '''
    new_rows_buffer.sort(key=lambda x: x['dept_row'])
    added_index_values = []
    master_merge_lookup = build_merge_lookup(master_ws)

    for item in new_rows_buffer:
        idx_str = item['idx_str']
        prefix = item['prefix']
        anchor_idx_str = item['anchor_idx_str']

        insert_at = None
        if anchor_idx_str is not None and anchor_idx_str in master_map:
            insert_at = master_map[anchor_idx_str]['row_num'] + 1
        elif prefix and prefix in prefix_first_rows:
            insert_at = prefix_first_rows[prefix]
        if insert_at is None:
            insert_at = master_ws.max_row + 1

        ref_row = insert_at - 1 if insert_at > 1 else insert_at + 1
        apply_row_style_from_neighbor(master_ws, insert_at, ref_row, max_col)

        inserted_rows = master_insert_rows(
            master_ws, insert_at, 1, index_store, inserted_rows,
        )
        index_store[idx_str] = {'row_num': insert_at, 'values': {}}
        master_map = index_store_as_map(index_store)
        added_index_values.append(idx_str)
        master_merge_lookup = build_merge_lookup(master_ws)
        print(
            f'  -> [행 추가] 인덱스 {idx_str} -> 취합용 {insert_at}행 '
            f'(데이터·서식은 레이아웃 복원 후 반영)'
        )

    return inserted_rows, added_index_values, master_map, master_merge_lookup


def map_sync_prefix_alignment(master_ws, index_store, dept_cache, prefix_pattern,
                              skip_align, max_col, merge_lookup, inserted_rows):
    '''MAP: 부서 기준 구조적 행번호 불일치 시 insert·baseline 데이터 복원.'''
    if not dept_cache:
        return inserted_rows, merge_lookup, []
    inserted_rows, merge_lookup, align_applied = sync_all_prefix_alignments(
        master_ws, index_store, dept_cache, prefix_pattern,
        skip_align, max_col, merge_lookup, inserted_rows,
    )
    for prefix, delta, insert_at, count in align_applied:
        print(
            f'  -> [구조 동기화] {prefix} prefix {delta}행 삽입 @ {insert_at}행 '
            f'({count}개 인덱스 행번호 정렬)'
        )
    if align_applied:
        merge_lookup = build_merge_lookup(master_ws)
    return inserted_rows, merge_lookup, align_applied


def map_restore_after_structural(master_ws, index_store, merge_lookup, skip_indexes):
    '''MAP: insert 직후 IndexStore L열~ 데이터 복원 (밀린 행 보호).'''
    restore_index_store_data(
        master_ws, index_store, merge_lookup,
        skip_indexes=skip_indexes, mode='data',
    )


def map_restore_before_layout(master_ws, index_store, section_baseline, valid_prefixes,
                              merge_lookup, skip_indexes, inserted_rows):
    '''MAP: 레이아웃 템플릿 적용 직전 I~L·비인덱스 행 values 복원.'''
    restore_index_store_data(
        master_ws, index_store, merge_lookup,
        skip_indexes=skip_indexes, mode='layout',
    )
    restore_section_row_baseline(
        master_ws, section_baseline, valid_prefixes, merge_lookup,
        inserted_rows=inserted_rows, data_cols_only=True,
    )


def dept_apply_blue_cell_updates(master_ws, dept_cache, master_map, max_col,
                                 master_merge_lookup, all_cascade_rows, dept_dup_blocked):
    '''DEPT DIFF: 동일 인덱스 행의 파란색(L열~) 수정 반영.'''
    updated = 0
    for df, ctx in dept_cache.items():
        dept_ws = ctx['ws']
        dept_merge_lookup = ctx['merge_lookup']
        dept_map = ctx['index_map']
        dup_blocked = dept_dup_blocked.get(df, set())

        for idx_str, info in dept_map.items():
            if idx_str in dup_blocked:
                continue
            dept_row = info['row_num']
            if dept_row in all_cascade_rows:
                continue
            if idx_str not in master_map:
                continue
            target_row = master_map[idx_str]['row_num']
            if not row_has_blue(dept_ws, dept_row, max_col):
                continue
            copy_row_values_and_color(
                master_ws, dept_ws, target_row, dept_row, max_col,
                master_merge_lookup=master_merge_lookup,
                dept_merge_lookup=dept_merge_lookup,
                data_cols_only=True,
            )
            updated += 1
            print(
                f'  -> [수정 완료] 인덱스 {idx_str} -> 취합용 파일 {target_row}행 데이터 동기화'
            )
    return updated


def dept_apply_remap_and_new_rows(master_ws, master_map, fix_remap_queue, new_rows_buffer,
                                  max_col, master_merge_lookup, fill_add=None):
    '''DEPT DIFF: 재번호 remap + 신규 인덱스 행 데이터(H·L열~) 반영.'''
    updated = 0
    for remap in fix_remap_queue:
        master_idx = remap['master_idx']
        if master_idx not in master_map:
            continue
        master_row = master_map[master_idx]['row_num']
        copy_row_values_and_color(
            master_ws, remap['dept_ws'], master_row, remap['dept_row'], max_col,
            index_value=master_idx,
            master_merge_lookup=master_merge_lookup,
            dept_merge_lookup=remap['dept_merge_lookup'],
            data_cols_only=True,
        )
        updated += 1

    for item in new_rows_buffer:
        idx_str = item['idx_str']
        if idx_str not in master_map:
            continue
        master_row = master_map[idx_str]['row_num']
        copy_row_values_and_color(
            master_ws, item['dept_ws'], master_row, item['dept_row'], max_col,
            fill=fill_add, index_value=idx_str,
            master_merge_lookup=master_merge_lookup,
            dept_merge_lookup=item.get('dept_merge_lookup'),
            data_cols_only=True,
        )
    return updated


def dept_apply_layout_text_overlays(master_ws, master_map, fix_remap_queue, new_rows_buffer,
                                    max_col, master_merge_lookup):
    '''DEPT DIFF: remap/신규행 I~K layout 텍스트 + L열~ dept_apply_data_overlays에서 L~ 재적용.'''
    if not fix_remap_queue and not new_rows_buffer:
        return
    copy_dept_layout_values_for_rows(
        master_ws, fix_remap_queue, new_rows_buffer, master_map,
        max_col, master_merge_lookup,
    )


def _dept_remap_new_row_pairs(fix_remap_queue, new_rows_buffer, master_map):
    pairs = []
    for remap in fix_remap_queue:
        master_idx = remap['master_idx']
        if master_idx not in master_map:
            continue
        pairs.append((
            master_map[master_idx]['row_num'],
            remap['dept_row'],
            remap['dept_ws'],
            remap.get('dept_merge_lookup'),
        ))
    for item in new_rows_buffer:
        idx_str = item['idx_str']
        if idx_str not in master_map:
            continue
        pairs.append((
            master_map[idx_str]['row_num'],
            item['dept_row'],
            item['dept_ws'],
            item.get('dept_merge_lookup'),
        ))
    return pairs


def dept_apply_data_overlays(master_ws, master_map, fix_remap_queue, new_rows_buffer,
                             max_col, master_merge_lookup=None):
    '''DEPT DIFF: remap/신규행 L열~ 값·파란색 재적용 (layout 복원 후 부서 수정 최종 반영).'''
    if not fix_remap_queue and not new_rows_buffer:
        return
    blue_fill = PatternFill(
        start_color=TARGET_BLUE_CODES[1],
        end_color=TARGET_BLUE_CODES[1],
        fill_type='solid',
    )
    for master_row, dept_row, dept_ws, dept_merge_lookup in _dept_remap_new_row_pairs(
        fix_remap_queue, new_rows_buffer, master_map,
    ):
        for c in range(DEPT_DATA_COL_START, max_col + 1):
            d_cell_raw = dept_ws.cell(dept_row, c)
            d_cell = get_real_cell(dept_ws, d_cell_raw, dept_merge_lookup)
            is_blue = is_blue_cell(d_cell_raw) or is_blue_cell(d_cell)
            m_raw = master_ws.cell(master_row, c)
            m_cell = get_real_cell(master_ws, m_raw, master_merge_lookup)
            if isinstance(m_cell, openpyxl.cell.cell.MergedCell):
                continue
            if d_cell.value is not None or is_blue:
                m_cell.value = translate_dept_formula(
                    d_cell.value, d_cell_raw.coordinate, master_row, dept_row,
                )
            if is_blue:
                m_cell.fill = blue_fill

import re

_EXT_LINK_PATTERN = re.compile(r'\[\d+\](?=[^!\]]+!)')
_EXT_FILE_LINK_PATTERN = re.compile(
    r"'[^']*\[(?:[^\]]+\.(?:xlsx|xlsm|xlsb|xls))\]([^']+)'!",
    re.IGNORECASE,
)


def _localize_ext_formula_text(formula: str) -> str | None:
    '''외부 통합문서 참조 수식 → 내부 시트 참조. 변경 없으면 None.'''
    if not (isinstance(formula, str) and formula.startswith('=')):
        return None
    new_v = _EXT_FILE_LINK_PATTERN.sub(r"'\1'!", formula)
    new_v = _EXT_LINK_PATTERN.sub('', new_v)
    return new_v if new_v != formula else None


def strip_external_link_refs(wb, sheet_names=None):
    '''
    수식 안의 외부 통합문서 참조를 같은 파일 내부 참조로 변환합니다.
    예: =MATCH([1]DashBoard2!F30,A3:A5,0) -> =MATCH(DashBoard2!F30,A3:A5,0)
        ='C:\\path\\[file.xlsx]취합(2606)'!A1 -> ='취합(2606)'!A1
    '''
    fixed = []
    sheets = sheet_names or wb.sheetnames
    for name in sheets:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                new_v = _localize_ext_formula_text(v) if isinstance(v, str) else None
                if new_v is None:
                    continue
                fixed.append((name, cell.coordinate, v, new_v))
                cell.value = new_v
    return fixed


def repair_master_external_link_refs(path, sheet_names=('Parameter',)):
    '''
    취합용 원본 xlsx의 깨진 외부 참조([n]Sheet!)를 내부 참조로 고친 뒤 저장.

    openpyxl로 저장하면 DashBoard 도면이 깨질 수 있어 Excel COM으로만 저장한다.
    기본 대상: Parameter 시트 (A6:E6의 MATCH([1]DashBoard2!...) 등).
    반환: [(sheet, coord, old, new), ...]
    '''
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        return []

    import win32com.client as win32

    excel = None
    wb = None
    fixed = []
    try:
        excel = win32.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            excel.AskToUpdateLinks = False
        except Exception:
            pass

        # UpdateLinks=0: 깨진 외부 링크 갱신 시도 없이 연다
        wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=False)

        if sheet_names:
            targets = list(sheet_names)
        else:
            targets = [
                wb.Worksheets(i).Name
                for i in range(1, wb.Worksheets.Count + 1)
            ]

        for name in targets:
            try:
                ws = wb.Worksheets(name)
            except Exception:
                continue
            try:
                used = ws.UsedRange
                row0, col0 = int(used.Row), int(used.Column)
                n_rows, n_cols = int(used.Rows.Count), int(used.Columns.Count)
            except Exception:
                row0, col0, n_rows, n_cols = 1, 1, 20, 10

            for r in range(row0, row0 + n_rows):
                for c in range(col0, col0 + n_cols):
                    cell = ws.Cells(r, c)
                    try:
                        formula = cell.Formula
                    except Exception:
                        continue
                    new_f = _localize_ext_formula_text(formula) if isinstance(formula, str) else None
                    if new_f is None:
                        continue
                    cell.Formula = new_f
                    coord = cell.Address.replace('$', '')
                    fixed.append((name, coord, formula, new_f))

        if fixed:
            wb.Save()
        return fixed
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

def fix_formulas_after_inserts(ws, inserted_rows, max_col, merge_lookup=None):
    '''
    insert_rows()는 수식 행 참조를 갱신하지 않음.
    삽입된 행 수만큼 Translator로 L열~ 수식을 보정 (잔액·SUMIF 행 포함).
    '''
    if not inserted_rows:
        return 0
    fixed = 0
    for r in range(1, ws.max_row + 1):
        row_delta = sum(1 for ir in inserted_rows if ir < r)
        if row_delta == 0:
            continue
        origin_row = r - row_delta
        for c in range(DEPT_DATA_COL_START, max_col + 1):
            raw = ws.cell(r, c)
            cell = get_real_cell(ws, raw, merge_lookup)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue
            val = cell.value
            if not _is_formula(val):
                continue
            new_val = translate_dept_formula(
                val, f'{get_column_letter(c)}{origin_row}', r, origin_row,
            )
            if new_val != val:
                cell.value = new_val
                fixed += 1
    return fixed


def find_dept_sheet_for_prefix(dept_cache, prefix, prefix_pattern):
    '''prefix(DB1, DB2…)를 담당하는 부서 시트·merge_lookup 반환.'''
    for ctx in dept_cache.values():
        for idx in ctx['index_map']:
            m = prefix_pattern.match(idx)
            if m and m.group(1) == prefix:
                return ctx['ws'], ctx['merge_lookup']
    return None, None


def copy_dept_row_data_to_master(master_ws, dept_ws, master_row, dept_row, max_col,
                                 master_merge_lookup=None, dept_merge_lookup=None):
    '''부서 한 행 L열~ → 마스터 (수식은 row_delta=master_row-dept_row 로 Translator 보정).'''
    blue_fill = PatternFill(
        start_color=TARGET_BLUE_CODES[1],
        end_color=TARGET_BLUE_CODES[1],
        fill_type='solid',
    )
    for c in range(DEPT_DATA_COL_START, max_col + 1):
        d_raw = dept_ws.cell(dept_row, c)
        d_cell = get_real_cell(dept_ws, d_raw, dept_merge_lookup)
        is_blue = is_blue_cell(d_raw) or is_blue_cell(d_cell)
        m_raw = master_ws.cell(master_row, c)
        m_cell = get_real_cell(master_ws, m_raw, master_merge_lookup)
        if isinstance(m_cell, openpyxl.cell.cell.MergedCell):
            continue
        if _is_formula(d_cell.value):
            continue
        if _is_formula(m_cell.value):
            continue
        if d_cell.value is None and not is_blue:
            continue
        m_cell.value = d_cell.value
        if is_blue:
            m_cell.fill = blue_fill


def dept_sync_sections_from_dept(master_ws, dept_cache, valid_prefixes, prefix_pattern,
                                 max_col, master_merge_lookup):
    '''
    insert·layout 후 prefix 섹션 전체(잔액·SUMIF 등 비인덱스 행 포함) L열~를 부서에서 동기화.
    섹션 헤더 row_offset = master_start - dept_start 으로 행 매핑, 수식 행 참조 보정.
    '''
    updated = 0
    for master_start, prefix in get_section_header_rows(master_ws, valid_prefixes):
        dept_ws, dept_ml = find_dept_sheet_for_prefix(dept_cache, prefix, prefix_pattern)
        if dept_ws is None:
            continue
        dept_header_map = {p: r for r, p in get_section_header_rows(dept_ws, valid_prefixes)}
        if prefix not in dept_header_map:
            continue
        dept_start = dept_header_map[prefix]
        row_offset = master_start - dept_start
        dept_end = get_section_end_row(dept_ws, prefix, valid_prefixes)
        master_end = get_section_end_row(
            master_ws, prefix, valid_prefixes,
            fallback=dept_end + row_offset,
        )
        for dr in range(dept_start, dept_end + 1):
            mr = dr + row_offset
            if mr > master_end:
                break
            copy_dept_row_data_to_master(
                master_ws, dept_ws, mr, dr, max_col,
                master_merge_lookup, dept_ml,
            )
            updated += 1
    return updated


def dept_sync_all_index_rows(master_ws, dept_cache, master_map, max_col,
                             master_merge_lookup, skip_indexes, all_cascade_rows,
                             dept_dup_blocked):
    '''하위 호환 — 섹션 전체 동기화로 대체 (인덱스-only는 잔액 SUMIF 행 누락).'''
    prefix_pattern = re.compile(r'^([A-Za-z0-9]+)-(\d+)$')
    valid_prefixes = get_valid_prefixes()
    return dept_sync_sections_from_dept(
        master_ws, dept_cache, valid_prefixes, prefix_pattern,
        max_col, master_merge_lookup,
    )


def dept_mark_deleted_indexes(master_ws, index_store, dept_cache, master_map,
                              prefix_pattern, valid_prefixes, fill_del, skip_indexes=None):
    '''DEPT DIFF: prefix 담당 부서에 없는 인덱스 H열 빨간색. Returns deleted_count.'''
    confirmed = collect_deletions_per_dept(
        index_store, dept_cache, prefix_pattern, master_map, valid_prefixes,
        skip_indexes=skip_indexes,
    )
    count = 0
    for del_idx in confirmed:
        del_row = master_map[del_idx]['row_num']
        master_ws.cell(row=del_row, column=INDEX_COL).fill = fill_del
        count += 1
        print(f'  -> [삭제] 인덱스 {del_idx} (취합용 {del_row}행) H열 빨간색')
    return count


def excel_layout_restore_after_inserts(master_ws, layout_snapshot, valid_prefixes,
                                       prefix_pattern, index_store, section_baseline,
                                       master_map, fix_remap_queue, new_rows_buffer,
                                       skip_align, inserted_rows, added_index_values,
                                       max_col, fill_add):
    '''
    EXCEL LAYOUT: insert 후 병합·서식·J열 구분·신규 H 연두색 복원.
    Returns master_merge_lookup.
    '''
    merge_lookup = build_merge_lookup(master_ws)

    map_restore_before_layout(
        master_ws, index_store, section_baseline, valid_prefixes,
        merge_lookup, skip_align, inserted_rows,
    )

    restore_master_layouts_after_inserts(
        master_ws, layout_snapshot, valid_prefixes, inserted_rows=inserted_rows,
    )
    merge_lookup = build_merge_lookup(master_ws)

    dept_apply_layout_text_overlays(
        master_ws, master_map, fix_remap_queue, new_rows_buffer,
        max_col, merge_lookup,
    )

    dept_apply_data_overlays(
        master_ws, master_map, fix_remap_queue, new_rows_buffer,
        max_col, merge_lookup,
    )

    rebuild_all_j_category_merges(
        master_ws, valid_prefixes, prefix_pattern, merge_lookup,
    )
    merge_lookup = build_merge_lookup(master_ws)

    ensure_all_section_i_name_merges(master_ws, valid_prefixes, merge_lookup)

    if added_index_values:
        reapply_index_marker_fills(
            master_ws, added_index_values, fill_add, index_map=master_map,
        )

    return merge_lookup


def collect_deletions_per_dept(index_store, dept_cache, prefix_pattern, master_map, valid_prefixes,
                               skip_indexes=None):
    '''
    prefix를 담당하는 부서 파일에서 해당 인덱스가 없으면 삭제로 확정.
    (합집합 all_dept_indexes만 쓰면 다른 부서 파일 때문에 누락되는 문제 방지)
    '''
    confirmed = set()
    skip = skip_indexes or set()
    numbered = [k for k in index_store if prefix_pattern.match(k)]
    for ctx in dept_cache.values():
        dept_map = ctx['index_map']
        dept_prefixes = collect_index_prefixes(dept_map, prefix_pattern)
        for idx in numbered:
            if idx in skip:
                continue
            match = prefix_pattern.match(idx)
            if not match:
                continue
            prefix = match.group(1)
            if prefix not in dept_prefixes:
                continue
            if not is_allowed_index_prefix(prefix, master_map, valid_prefixes):
                continue
            if idx not in dept_map:
                confirmed.add(idx)
    return confirmed


def _prefix_row_mismatches(index_store, dept_map, master_map, prefix, prefix_pattern, skip_indexes):
    mismatches = []
    for idx, rec in index_store.items():
        match = prefix_pattern.match(idx)
        if not match or match.group(1) != prefix:
            continue
        if idx in skip_indexes:
            continue
        if idx not in dept_map or idx not in master_map:
            continue
        master_row = master_map[idx]['row_num']
        dept_row = dept_map[idx]['row_num']
        if master_row != dept_row:
            mismatches.append((idx, master_row, dept_row))
    return mismatches


def sync_all_prefix_alignments(master_ws, index_store, dept_cache, prefix_pattern,
                               skip_indexes, max_col, merge_lookup, inserted_rows):
    '''
    부서 파일 기준으로 master 행번호를 맞춤.
    같은 prefix 내에서 master↔dept 행 차이가 일정하면 insert_rows로 구조 동기화 후
    baseline 데이터를 해당 행에 복원.
    '''
    prefix_to_dept = {}
    for ctx in dept_cache.values():
        dept_map = ctx['index_map']
        for prefix in collect_index_prefixes(dept_map, prefix_pattern):
            if prefix not in prefix_to_dept:
                prefix_to_dept[prefix] = dept_map

    master_map = index_store_as_map(index_store)
    planned = []
    for prefix, dept_map in prefix_to_dept.items():
        mismatches = _prefix_row_mismatches(
            index_store, dept_map, master_map, prefix, prefix_pattern, skip_indexes,
        )
        if not mismatches:
            continue
        delta = mismatches[0][2] - mismatches[0][1]
        if delta <= 0:
            continue
        if not all(dr - mr == delta for _, mr, dr in mismatches):
            continue
        insert_at = min(mr for _, mr, _ in mismatches)
        planned.append((insert_at, delta, prefix, len(mismatches)))

    planned.sort(key=lambda x: x[0], reverse=True)
    applied = []
    for insert_at, delta, prefix, count in planned:
        inserted_rows = master_insert_rows(
            master_ws, insert_at, delta, index_store, inserted_rows,
        )
        merge_lookup = build_merge_lookup(master_ws)
        master_map = index_store_as_map(index_store)
        for idx, rec in index_store.items():
            match = prefix_pattern.match(idx)
            if not match or match.group(1) != prefix:
                continue
            if idx in skip_indexes or idx not in master_map:
                continue
            apply_baseline_row_data(
                master_ws, master_map[idx]['row_num'], rec['values'],
                merge_lookup, data_cols_only=True,
            )
        applied.append((prefix, delta, insert_at, count))

    return inserted_rows, merge_lookup, applied

        
def build_prefix_first_rows(index_map, prefix_pattern):
    '''prefix별 첫 행 번호 (신규행 삽입 위치 계산용).'''
    prefix_first = {}
    for idx_str, info in index_map.items():
        prefix = extract_prefix(idx_str, prefix_pattern)
        if not prefix:
            continue
        row = info['row_num']
        if prefix not in prefix_first or row < prefix_first[prefix]:
            prefix_first[prefix] = row
    return prefix_first


def build_dept_sheet_cache(dept_wbs, sheet_name):
    '''
    부서 workbook×시트별 index_map·merge_lookup·rows_sorted를 1회만 계산.
    numeric 3시트 루프에서 재사용 → 시트당 8회 스캔을 1회로 축소.
    '''
    cache = {}
    for path, wb in dept_wbs.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        merge_lookup = build_merge_lookup(ws)
        index_map = get_sheet_index_map(ws, merge_lookup, include_name=False)
        rows_sorted = sorted((info['row_num'], idx) for idx, info in index_map.items())
        cache[path] = {
            'ws': ws,
            'merge_lookup': merge_lookup,
            'index_map': index_map,
            'rows_sorted': rows_sorted,
        }
    return cache

def get_next_index_number(master_map, prefix):
    """부서에서 새로운 데이터 행을 만들었을 때, 기존 취합용의 번호(예: DB2-5) 뒤를 이어받을 '다음 차례 순번(DB2-6)'을 계산합니다."""
    max_num = 0
    pattern = re.compile(rf"^{re.escape(prefix)}-(\d+)$")
    for idx_str in master_map.keys():
        match = pattern.match(idx_str)
        if match:
            num = int(match.group(1))
            if num > max_num: max_num = num
    return max_num + 1


def get_suggested_new_index(master_map, prefix, pending_add_indexes=None):

    all_indexes = set(master_map.keys())

    if pending_add_indexes:
        all_indexes |= pending_add_indexes

    max_num = 0

    pattern = re.compile(rf"^{re.escape(prefix)}-(\d+)$")

    for idx in all_indexes:
        m = pattern.match(idx)
        if m:
            max_num = max(max_num, int(m.group(1)))

    return f"{prefix}-{max_num + 1}"


def find_duplicate_index_rows(ws):
    '''H열에 같은 인덱스가 2행 이상이면 [(idx, first_row, dup_row), ...]'''
    first_row = {}
    dups = []
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=INDEX_COL).value
        if v is None or not str(v).strip():
            continue
        idx = str(v).strip()
        if idx in first_row:
            dups.append((idx, first_row[idx], r))
        else:
            first_row[idx] = r
    return dups


def validate_new_index(idx_str, prefix, master_map, pending_add_indexes, prefix_pattern):
    '''
    신규 추가 인덱스 검증: 취합용 해당 prefix 최대번호+1만 허용.
    (중간 삽입 시 U-5 재사용·아래 행 U-6~ 재번호 방지)
    Returns (ok, message, suggested_idx)
    '''
    if not prefix_pattern.match(idx_str):
        return False, f'번호형 인덱스 형식이 아닙니다 ({idx_str})', None
    if idx_str in master_map:
        return False, f'{idx_str}는 취합용에 이미 있습니다 (값 변경은 파란색 수정)', None
    suggested = get_suggested_new_index(master_map, prefix, pending_add_indexes)
    if idx_str != suggested:
        return False, (
            f'신규행 인덱스 {idx_str} 불가 — 중간 삽입 시 기존 번호 재사용·재번호 금지. '
            f'→ {suggested} 를 입력하세요'
        ), suggested
    return True, '', suggested


def get_neighbor_indexes(dept_map, dept_row):
    '''dept_row 바로 위·아래(행번호 기준)에 있는 인덱스 문자열'''
    prev_idx, next_idx = None, None
    prev_row, next_row = -1, 10**9
    for idx, info in dept_map.items():
        r = info['row_num']
        if r < dept_row and r > prev_row:
            prev_row = r
            prev_idx = idx
        elif r > dept_row and r < next_row:
            next_row = r
            next_idx = idx
    return prev_idx, next_idx


def get_prefix_master_max(master_map, prefix, prefix_pattern):
    '''취합용 index_map에서 prefix별 최대 번호 (예: U-8 → 8). 없으면 0.'''
    max_num = 0
    for idx_str in master_map:
        match = prefix_pattern.match(idx_str)
        if match and match.group(1) == prefix:
            max_num = max(max_num, int(match.group(2)))
    return max_num


def is_bottom_append_only(dept_map, master_map, prefix, master_max, prefix_pattern):
    '''
    U-1~U-max 행 위치가 취합용과 동일하고 U-(max+1)만 바로 아래 붙은 경우
    → 정상 하단 추가로 보고 재번호 자동 보정 제외.
    '''
    if master_max == 0:
        return False
    for k in range(1, master_max + 1):
        idx = f'{prefix}-{k}'
        if idx not in dept_map or idx not in master_map:
            return False
        if dept_map[idx]['row_num'] != master_map[idx]['row_num']:
            return False
    extra = f'{prefix}-{master_max + 1}'
    tail = f'{prefix}-{master_max}'
    if extra not in dept_map:
        return False
    return dept_map[extra]['row_num'] == master_map[tail]['row_num'] + 1


def dept_has_deleted_master_indices(dept_map, master_map, prefix, master_max, prefix_pattern):
    '''취합용에는 있는데 부서에서 빠진 번호(DB2-2 삭제 등)가 있으면 True.'''
    for k in range(1, master_max + 1):
        idx = f'{prefix}-{k}'
        if idx in master_map and idx not in dept_map:
            return True
    return False


def is_bottom_append_with_gaps(dept_map, master_map, prefix, master_max, prefix_pattern):
    '''
    중간 번호가 삭제된 경우에도, 존재하는 번호는 취합용과 같은 행이고
    max+1만 맨 아래 붙었으면 정상 하단 추가.
    '''
    extra = f'{prefix}-{master_max + 1}'
    if extra not in dept_map:
        return False
    for k in range(1, master_max + 1):
        idx = f'{prefix}-{k}'
        if idx in master_map and idx in dept_map:
            if dept_map[idx]['row_num'] != master_map[idx]['row_num']:
                return False
    present = [
        f'{prefix}-{k}' for k in range(1, master_max + 1)
        if f'{prefix}-{k}' in dept_map
    ]
    if not present:
        return False
    last_idx = max(present, key=lambda i: dept_map[i]['row_num'])
    return dept_map[extra]['row_num'] == dept_map[last_idx]['row_num'] + 1


def row_content_signature(ws, row, max_col, merge_lookup=None):
    '''H열 제외 행 내용 비교용 (재번호 연쇄 vs 하단 추가 구분).'''
    parts = []
    for c in range(1, max_col + 1):
        if c == INDEX_COL:
            continue
        cell = get_real_cell(ws, ws.cell(row=row, column=c), merge_lookup)
        v = cell.value
        parts.append('' if v is None else str(v).strip())
    return tuple(parts)


def confirm_renumber_shift(dept_ws, master_ws, plan, master_map, dept_map, max_col,
                           prefix_pattern, dept_merge_lookup=None,
                           master_merge_lookup=None):
    '''
    재번호 연쇄 확정:
    - 삽입행(U-n) 내용 ≠ 취합용 U-n
    - 취합용 U-k 내용 = 부서 U-(k+1) (k = n .. master_max-1)
    '''
    insert_idx = plan['insert_idx']
    insert_row = plan['insert_row']
    prefix = plan['prefix']
    if insert_idx not in master_map:
        return False
    match = prefix_pattern.match(insert_idx)
    if not match:
        return False
    n = int(match.group(2))
    master_max = get_prefix_master_max(master_map, prefix, prefix_pattern)

    master_row = master_map[insert_idx]['row_num']
    if row_content_signature(master_ws, master_row, max_col, master_merge_lookup) == \
       row_content_signature(dept_ws, insert_row, max_col, dept_merge_lookup):
        return False

    for k in range(n, master_max):
        m_idx = f'{prefix}-{k}'
        d_idx = f'{prefix}-{k + 1}'
        if m_idx not in master_map or d_idx not in dept_map:
            return False
        m_row = master_map[m_idx]['row_num']
        d_row = dept_map[d_idx]['row_num']
        if row_content_signature(master_ws, m_row, max_col, master_merge_lookup) != \
           row_content_signature(dept_ws, d_row, max_col, dept_merge_lookup):
            return False
    return True


def confirm_renumber_shift_light(dept_ws, master_ws, plan, master_map, dept_map, max_col,
                                 prefix_pattern, dept_merge_lookup=None,
                                 master_merge_lookup=None):
    '''
    내용 전체 비교가 #REF! 등으로 실패할 때 사용하는 완화 검증.
    삽입행만 다르고, 바로 아래 행이 취합용 U-n 내용을 담고 있으면 재번호로 판단.
    '''
    insert_idx = plan['insert_idx']
    insert_row = plan['insert_row']
    prefix = plan['prefix']
    if insert_idx not in master_map:
        return False
    match = prefix_pattern.match(insert_idx)
    if not match:
        return False
    n = int(match.group(2))
    next_idx = f'{prefix}-{n + 1}'
    if next_idx not in dept_map or dept_map[next_idx]['row_num'] != insert_row + 1:
        return False
    master_row = master_map[insert_idx]['row_num']
    if row_content_signature(master_ws, master_row, max_col, master_merge_lookup) == \
       row_content_signature(dept_ws, insert_row, max_col, dept_merge_lookup):
        return False
    return row_content_signature(master_ws, master_row, max_col, master_merge_lookup) == \
           row_content_signature(
               dept_ws, dept_map[next_idx]['row_num'], max_col, dept_merge_lookup
           )


def confirm_renumber_shift_structural(plan, master_map, dept_map, prefix_pattern,
                                      dept_ws=None, master_ws=None, max_col=None,
                                      dept_merge_lookup=None, master_merge_lookup=None):
    '''
    내용 비교(#REF 등)가 불가할 때: 삽입행만 다르고, 아래 번호가 한 칸씩 밀린 배치면 재번호로 확정.
    U-5 삽입 시 U-6@371+1, …, U-9@371+(max-n+1) 위치를 검사해 U-3 등 오탐을 차단.
    '''
    insert_idx = plan['insert_idx']
    insert_row = plan['insert_row']
    prefix = plan['prefix']
    new_index = plan['new_index']
    if insert_idx not in master_map or new_index not in dept_map:
        return False
    match = prefix_pattern.match(insert_idx)
    if not match or match.group(1) != prefix:
        return False
    n = int(match.group(2))
    master_max = get_prefix_master_max(master_map, prefix, prefix_pattern)
    if master_max == 0:
        return False

    for k in range(1, n):
        idx = f'{prefix}-{k}'
        if idx not in dept_map or idx not in master_map:
            return False
        if dept_map[idx]['row_num'] != master_map[idx]['row_num']:
            return False

    if dept_ws is not None and master_ws is not None and max_col is not None:
        master_row = master_map[insert_idx]['row_num']
        if row_content_signature(master_ws, master_row, max_col, master_merge_lookup) == \
           row_content_signature(dept_ws, insert_row, max_col, dept_merge_lookup):
            return False

    expected_last = insert_row + (master_max - n + 1)
    if dept_map[new_index]['row_num'] != expected_last:
        return False

    for k in range(n + 1, master_max + 2):
        d_idx = f'{prefix}-{k}'
        expected_row = insert_row + (k - n)
        if d_idx not in dept_map or dept_map[d_idx]['row_num'] != expected_row:
            return False
    return True


def confirm_renumber_shift_minimal(dept_ws, master_ws, plan, master_map, max_col,
                                   dept_merge_lookup=None, master_merge_lookup=None):
    '''삽입행(U-n) 내용만 취합용과 다르면 재번호 삽입으로 판단 (완화).'''
    insert_idx = plan['insert_idx']
    insert_row = plan['insert_row']
    if insert_idx not in master_map:
        return False
    master_row = master_map[insert_idx]['row_num']
    return row_content_signature(master_ws, master_row, max_col, master_merge_lookup) != \
           row_content_signature(dept_ws, insert_row, max_col, dept_merge_lookup)


def detect_same_row_renumber_insert(prefix, dept_map, master_map, prefix_pattern,
                                    dept_ws=None, master_ws=None, max_col=None,
                                    dept_merge_lookup=None, master_merge_lookup=None):
    '''
    U-1~U-max 행번호가 취합용과 같고 U-(max+1)만 맨 아래 추가된 재번호 실수.
    U-1~U-(max-1) 내용이 같고 U-max부터 첫 불일치 = 삽입 시작(U-5 등).
    '''
    master_max = get_prefix_master_max(master_map, prefix, prefix_pattern)
    if master_max == 0:
        return None
    if dept_has_deleted_master_indices(dept_map, master_map, prefix, master_max, prefix_pattern):
        return None

    new_index = f'{prefix}-{master_max + 1}'
    if new_index not in dept_map:
        return None

    for k in range(1, master_max + 1):
        idx = f'{prefix}-{k}'
        if idx not in dept_map or idx not in master_map:
            return None
        if dept_map[idx]['row_num'] != master_map[idx]['row_num']:
            return None

    last_idx = f'{prefix}-{master_max}'
    if dept_map[new_index]['row_num'] != master_map[last_idx]['row_num'] + 1:
        return None

    if dept_ws is None or master_ws is None or max_col is None:
        return None

    insert_k = None
    insert_row = None
    for k in range(1, master_max + 1):
        idx = f'{prefix}-{k}'
        m_row = master_map[idx]['row_num']
        d_row = dept_map[idx]['row_num']
        if row_content_signature(master_ws, m_row, max_col, master_merge_lookup) == \
           row_content_signature(dept_ws, d_row, max_col, dept_merge_lookup):
            continue
        if k >= master_max:
            break
        below_idx = f'{prefix}-{k + 1}'
        if below_idx not in dept_map or below_idx not in master_map:
            break
        if dept_map[below_idx]['row_num'] != master_map[below_idx]['row_num']:
            break
        m_below = master_map[below_idx]['row_num']
        d_below = dept_map[below_idx]['row_num']
        if row_content_signature(master_ws, m_below, max_col, master_merge_lookup) == \
           row_content_signature(dept_ws, d_below, max_col, dept_merge_lookup):
            continue
        insert_k = k
        insert_row = d_row
        break

    if insert_k is None:
        return None

    insert_idx = f'{prefix}-{insert_k}'
    remaps = []
    cascade_rows = {insert_row}
    cascade_dept_indexes = {insert_idx}
    for k in range(insert_k + 1, master_max + 2):
        dept_idx = f'{prefix}-{k}'
        master_idx = f'{prefix}-{k - 1}'
        d_row = dept_map[dept_idx]['row_num']
        remaps.append((dept_idx, d_row, master_idx))
        cascade_rows.add(d_row)
        cascade_dept_indexes.add(dept_idx)

    # 신규행(U-9)은 부서에서 삽입한 물리 위치(예: U-4 바로 아래 371행)에 들어가야 함.
    # 삽입 기준 = 바로 위 인덱스(U-4). (예전에 U-8을 기준으로 쓰면 맨 아래에 넣어져 U-5~U-8·잔액 구역이 어긋남)
    # 바로 위 인덱스 — 신규행을 넣을 때 그 아래 row_num을 구하는 기준 (예: U-4)
    anchor_idx = f'{prefix}-{insert_k - 1}' if insert_k > 1 else None
    return {
        'prefix': prefix,
        'insert_row': insert_row,
        'insert_idx': insert_idx,
        'anchor_idx': anchor_idx,
        'remaps': remaps,
        'new_index': new_index,
        'cascade_rows': cascade_rows,
        'cascade_dept_indexes': cascade_dept_indexes,
    }


def detect_renumber_cascade_candidates(prefix, dept_map, master_map, prefix_pattern):
    '''
    중간 삽입 + 재번호 연쇄 후보 목록 (U-4→U-5(신규)→U-6…→U-9).
    여러 시작점이 잡히면 build 단계에서 내용 검증 후 가장 뒤쪽(n 큰) 후보 채택.
    '''
    master_max = get_prefix_master_max(master_map, prefix, prefix_pattern)
    if master_max == 0:
        return []

    if dept_has_deleted_master_indices(dept_map, master_map, prefix, master_max, prefix_pattern):
        return []

    cascade_end_num = master_max + 1
    new_index = f'{prefix}-{cascade_end_num}'
    candidates = []

    for idx_str, info in dept_map.items():
        if idx_str not in master_map:
            continue
        match = prefix_pattern.match(idx_str)
        if not match or match.group(1) != prefix:
            continue
        n = int(match.group(2))
        if n >= master_max:
            continue

        dept_row = info['row_num']
        hijack, _, _ = detect_index_hijack(
            idx_str, dept_row, dept_map, master_map, prefix_pattern
        )
        if not hijack:
            continue

        cascade_len = master_max - n + 2
        remaps = []
        cascade_rows = set()
        cascade_dept_indexes = set()
        valid = True
        for k in range(cascade_len):
            expected_num = n + k
            expected_idx = f'{prefix}-{expected_num}'
            expected_row = dept_row + k
            cascade_dept_indexes.add(expected_idx)
            if expected_idx not in dept_map:
                valid = False
                break
            if dept_map[expected_idx]['row_num'] != expected_row:
                valid = False
                break
            cascade_rows.add(expected_row)
            if k >= 1:
                master_idx = f'{prefix}-{expected_num - 1}'
                remaps.append((expected_idx, expected_row, master_idx))

        if not valid:
            continue

        # 바로 위 인덱스 (예: U-5 삽입 → U-4)
        anchor_idx = f'{prefix}-{n - 1}' if n > 1 else None
        candidates.append({
            'prefix': prefix,
            'insert_row': dept_row,
            'insert_idx': idx_str,
            'anchor_idx': anchor_idx,
            'remaps': remaps,
            'new_index': new_index,
            'cascade_rows': cascade_rows,
            'cascade_dept_indexes': cascade_dept_indexes,
        })

    return candidates


def detect_renumber_cascade(prefix, dept_map, master_map, prefix_pattern):
    '''단일 후보만 필요할 때 (테스트용).'''
    cands = detect_renumber_cascade_candidates(prefix, dept_map, master_map, prefix_pattern)
    return cands[-1] if cands else None


def collect_index_prefixes(index_map, prefix_pattern):
    '''index_map 키에서 prefix 집합 추출 (U, DB2 등).'''
    prefixes = set()
    for idx_str in index_map:
        match = prefix_pattern.match(idx_str)
        if match:
            prefixes.add(match.group(1))
    return prefixes


def pick_best_renumber_plan(candidates, confirm_fn, dept_ws, master_ws, master_map,
                            dept_map, max_col, prefix_pattern,
                            dept_merge_lookup=None, master_merge_lookup=None):
    '''full/light 검증 통과 후보 중 insert 번호(n)가 가장 큰 plan = 실제 삽입점(U-5 등).'''
    passing = []
    for plan in candidates:
        if confirm_fn(
            dept_ws, master_ws, plan, master_map, dept_map, max_col,
            prefix_pattern, dept_merge_lookup, master_merge_lookup,
        ):
            passing.append(plan)
    if not passing:
        return None
    return max(
        passing,
        key=lambda p: int(prefix_pattern.match(p['insert_idx']).group(2)),
    )


def pick_best_renumber_plan_structural(candidates, dept_ws, master_ws, master_map, dept_map,
                                       max_col, prefix_pattern,
                                       dept_merge_lookup=None, master_merge_lookup=None):
    passing = []
    for plan in candidates:
        if confirm_renumber_shift_structural(
            plan, master_map, dept_map, prefix_pattern,
            dept_ws, master_ws, max_col, dept_merge_lookup, master_merge_lookup,
        ):
            passing.append(plan)
    if not passing:
        return None
    return min(
        passing,
        key=lambda p: int(prefix_pattern.match(p['insert_idx']).group(2)),
    )


def build_dept_index_fixup(dept_map, master_map, prefix_pattern, valid_prefixes,
                           dept_ws=None, master_ws=None, max_col=None,
                           dept_merge_lookup=None, master_merge_lookup=None):
    '''부서 시트 prefix별 재번호 자동 보정 계획. {prefix: plan}'''
    plans = {}
    can_confirm = all(x is not None for x in (
        dept_ws, master_ws, max_col,
    ))
    prefixes = collect_index_prefixes(dept_map, prefix_pattern) | \
                 collect_index_prefixes(master_map, prefix_pattern)
    for prefix in prefixes:
        same_row = detect_same_row_renumber_insert(
            prefix, dept_map, master_map, prefix_pattern,
            dept_ws, master_ws, max_col, dept_merge_lookup, master_merge_lookup,
        )
        if same_row:
            plans[prefix] = same_row
            continue

        candidates = detect_renumber_cascade_candidates(
            prefix, dept_map, master_map, prefix_pattern
        )
        if not candidates:
            continue
        chosen = pick_best_renumber_plan(
            candidates, confirm_renumber_shift,
            dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
            dept_merge_lookup, master_merge_lookup,
        )
        if not chosen:
            chosen = pick_best_renumber_plan(
                candidates, confirm_renumber_shift_light,
                dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
                dept_merge_lookup, master_merge_lookup,
            )
        if not chosen:
            chosen = pick_best_renumber_plan_structural(
                candidates, dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
                dept_merge_lookup, master_merge_lookup,
            )
        if chosen:
            plans[prefix] = chosen
    return plans


def should_skip_dept_new_index_add(idx_str, prefix, dept_map, master_map, prefix_pattern,
                                   fix_plans, dept_ws=None, master_ws=None, max_col=None,
                                   dept_merge_lookup=None, master_merge_lookup=None):
    '''
    재번호 실수로 부서 맨 아래 U-9가 있을 때, 하단 ADD 대신 자동보정에 맡김.
    fix_plans에 없어도 cascade+삽입행 내용 차이면 skip.
    '''
    if prefix in fix_plans:
        return True
    suggested = get_suggested_new_index(master_map, prefix)
    if idx_str != suggested or idx_str not in dept_map:
        return False
    if detect_same_row_renumber_insert(
        prefix, dept_map, master_map, prefix_pattern,
        dept_ws, master_ws, max_col, dept_merge_lookup, master_merge_lookup,
    ):
        return True
    cands = detect_renumber_cascade_candidates(prefix, dept_map, master_map, prefix_pattern)
    if not cands:
        return False
    if dept_ws is None or master_ws is None or max_col is None:
        return True
    if pick_best_renumber_plan(
        cands, confirm_renumber_shift_light,
        dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
        dept_merge_lookup, master_merge_lookup,
    ):
        return True
    if pick_best_renumber_plan(
        cands, confirm_renumber_shift,
        dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
        dept_merge_lookup, master_merge_lookup,
    ):
        return True
    return pick_best_renumber_plan_structural(
        cands, dept_ws, master_ws, master_map, dept_map, max_col, prefix_pattern,
        dept_merge_lookup, master_merge_lookup,
    ) is not None


def fix_plan_virtual_indexes(plan):
    '''보정 plan이 취합용에 유지한다고 볼 인덱스 집합 (삭제 오판 방지).'''
    indexes = {plan['new_index']}
    for _, _, master_idx in plan['remaps']:
        indexes.add(master_idx)
    return indexes


def detect_index_hijack(idx_str, dept_row, dept_map, master_map, prefix_pattern):
    '''
    신규행에 U-5를 넣고 바로 아래 기존 U-5를 U-6으로 밀어 올린 경우 감지.

    U-4 → U-5(신규) → U-6(구 U-5) 처럼 번호가 끊김 없이 이어지면 재번호 연쇄.
    U-4 → U-9(신규) → U-5 처럼 중간에 max+1이 끼면 정상(UPDATE 허용).
    Returns (is_hijack, message, suggested_idx)
    '''
    match = prefix_pattern.match(idx_str)
    if not match or idx_str not in master_map:
        return False, '', None
    prefix = match.group(1)
    num = int(match.group(2))
    prev_idx, next_idx = get_neighbor_indexes(dept_map, dept_row)
    expected_prev = f'{prefix}-{num - 1}' if num > 1 else None
    expected_next = f'{prefix}-{num + 1}'
    if next_idx != expected_next:
        return False, '', None
    if dept_map[next_idx]['row_num'] != dept_row + 1:
        return False, '', None
    if expected_prev is not None and prev_idx != expected_prev:
        return False, '', None
    suggested = get_suggested_new_index(master_map, prefix)
    master_row = master_map[idx_str]['row_num']
    return True, (
        f'{idx_str}는 취합용 {master_row}행의 기존 번호입니다. '
        f'중간 신규행(부서 {dept_row}행)에는 {suggested} 를 입력하고, '
        f'아래 행을 {expected_next}~ 으로 재번호하지 마세요'
    ), suggested


# ------------------------------------------
# 5. 고급 레이아웃 복원 도구
# - 행이 추가되거나 깨졌을 때 취합용 파일의 원래 디자인 양식을 복구하는 함수들입니다.

def extract_section_layout_templates(ws, valid_prefixes, col_end=None):
    """행이 추가되어 밀리기 전, 순수한 상태의 마스터 엑셀 디자인틀(병합 구조, 색상)을 기억장치에 저장해 둡니다."""
    if col_end is None: col_end = min(ws.max_column, STYLE_SNAPSHOT_MAX_COL)
    headers = get_section_header_rows(ws, valid_prefixes)
    result = {}
    for i, (start_row, prefix) in enumerate(headers):
        if i + 1 < len(headers): end_row = headers[i + 1][0] - 1
        else: end_row = find_section_content_end_row(ws, start_row)
        merges = []
        for mr in ws.merged_cells.ranges:
            if mr.max_col < LAYOUT_COL_START or mr.min_col > LAYOUT_COL_END: continue
            if mr.min_row < start_row or mr.max_row > end_row: continue
            anchor = ws.cell(mr.min_row, mr.min_col)
            merges.append({
                'rel_min_row': mr.min_row - start_row,
                'rel_max_row': mr.max_row - start_row,
                'min_col': mr.min_col, 'max_col': mr.max_col,
                'value': anchor.value, 'style': capture_cell_style(anchor),
            })
        row_cells = []
        for r in range(start_row, end_row + 1):
            for c in range(LAYOUT_COL_START, col_end + 1):
                cell = ws.cell(r, c)
                if isinstance(cell, openpyxl.cell.cell.MergedCell): continue
                if not should_capture_layout_cell(cell): continue
                row_cells.append({'rel_row': r - start_row, 'col': c, 'value': cell.value, 'style': capture_cell_style(cell)})
        result[prefix] = {'rel_end_row': end_row - start_row, 'merges': sorted(merges, key=lambda x: (x['rel_min_row'], x['min_col'])), 'row_cells': row_cells}
    return result


def _extend_layout_rows_from_template(ws, tmpl, start_row, template_end, end_row, valid_prefixes):
    '''섹션이 템플릿보다 길어진 경우, 템플릿 마지막 데이터행의 I~L 서식을 추가 행에 복제.'''
    rel_last = template_end - start_row
    if rel_last < 1:
        return
    last_cells = [rc for rc in tmpl['row_cells'] if rc['rel_row'] == rel_last]
    if not last_cells:
        return
    for abs_row in range(template_end + 1, end_row + 1):
        for rc in last_cells:
            cell = ws.cell(abs_row, rc['col'])
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue
            if is_layout_col(rc['col']):
                apply_cell_style(cell, rc.get('style'), apply_value=False)


def upsert_section_layout_template(templates, prefix, tmpl):
    '''같은 prefix는 rel_end_row가 더 긴(행 추가 반영) 템플릿만 유지.'''
    cur = templates.get(prefix)
    if cur is None or tmpl['rel_end_row'] >= cur['rel_end_row']:
        templates[prefix] = tmpl


def upsert_layout_templates_for_row_add(templates, dept_ws, item_prefix, valid_prefixes, col_end):
    '''행 추가 prefix와 그 다음 섹션 헤더 템플릿만 부서 파일에서 반영.'''
    extracted = extract_section_layout_templates(dept_ws, valid_prefixes, col_end=col_end)
    if item_prefix in extracted:
        upsert_section_layout_template(templates, item_prefix, extracted[item_prefix])
    headers = [p for _, p in get_section_header_rows(dept_ws, valid_prefixes)]
    if item_prefix in headers:
        idx = headers.index(item_prefix)
        if idx + 1 < len(headers):
            next_prefix = headers[idx + 1]
            if next_prefix in extracted:
                upsert_section_layout_template(templates, next_prefix, extracted[next_prefix])


def reapply_pristine_section_layouts(ws, original_templates, skip_prefixes, valid_prefixes,
                                      inserted_rows=None):
    '''행 삽입으로 깨진 미변경 섹션(UP, DB1 등)을 취합용 원본 레이아웃으로 복원.

    inserted_rows: master_ws에 새로 삽입된 절대 행 번호 집합. apply_section_layout_templates로
        그대로 전달되어, 삽입 지점 아래 행들이 어긋나 덮어써지는 문제를 방지합니다.
    '''
    skip = skip_prefixes or set()
    pristine = {p: t for p, t in original_templates.items() if p not in skip}
    if pristine:
        apply_section_layout_templates(
            ws, pristine, valid_prefixes, finalize_i_label=False, inserted_rows=inserted_rows,
        )


def restore_master_layouts_after_inserts(ws, original_templates, valid_prefixes, inserted_rows=None):
    '''
    행 삽입 후 I~L 레이아웃을 취합 시작 시점 master 스냅샷으로 복원합니다.

    부서 파일 템플릿은 행 배치(삽입 위치)가 master와 달라 병합이 어긋날 수 있으므로
    master 원본만 사용하고, 섹션이 길어진 경우 I열 세로 병합을 end_row까지 확장합니다.
    '''
    if not original_templates:
        return
    inserted_rows = inserted_rows or set()
    apply_section_layout_templates(
        ws, original_templates, valid_prefixes,
        finalize_i_label=False, inserted_rows=inserted_rows,
    )
    apply_border_overlay_from_snapshot(
        ws, original_templates, valid_prefixes, inserted_rows=inserted_rows,
    )
    for start_row, prefix in get_section_header_rows(ws, valid_prefixes):
        if prefix not in original_templates:
            continue
        end_row = get_section_end_row(ws, prefix, valid_prefixes)
        extend_section_i_name_merge_to_end(
            ws, start_row, end_row, prefix=prefix, max_gap=None,
        )
        apply_i_section_label(ws, start_row, prefix)


def build_rel_row_to_abs_row_map(start_row, end_row, inserted_rows, rel_end_row):
    """
    [버그 수정용 헬퍼] 템플릿의 상대 위치(rel_row)를 실제 시트의 절대 행 번호로 정확히 매핑합니다.

    [문제 배경] 기존 코드는 '절대행 = start_row + rel_row' 로 단순 계산했습니다.
    이는 섹션 안에 새 행이 삽입되지 않았을 때만 맞는 계산입니다.
    행이 중간에 삽입되면 그 아래 실제 행들은 한 칸씩 밀려나지만, rel_row는
    삽입 전 기준으로 고정되어 있어 결과적으로 엉뚱한 행에 옛 값이 덮어씌워집니다.

    [해결 원리] '삽입되지 않은(원본부터 있던) 행'만 순서대로 세어서,
    rel_row번째에 해당하는 '원본 행'을 찾으면 그것이 곧 정답 절대행입니다.
    즉 삽입된 행은 세지 않고 건너뛰면서(skip) 카운트합니다.

    반환값: {rel_row: 절대행} 매핑 딕셔너리. rel_row가 가리킬 원본 행이
    삽입행 때문에 더 이상 존재하지 않는 경우(섹션이 비정상적으로 짧아진 경우)는
    매핑에서 제외됩니다 — 호출부에서 안전하게 .get()으로 처리해야 합니다.
    """
    mapping = {}
    rel = 0
    r = start_row
    while r <= end_row and rel <= rel_end_row:
        if r not in inserted_rows:
            mapping[rel] = r
            rel += 1
        r += 1
    return mapping


def apply_section_layout_templates(ws, templates, valid_prefixes, finalize_i_label=True,
                                    inserted_rows=None):
    """새 행 삽입 때문에 뒤틀리고 찢어진 엑셀 양식틀(J/K열 글자, 병합, 수입/지출 색상)을 백업본을 바탕으로 완벽히 재조립합니다.

    inserted_rows: 이번 취합 작업 중 master_ws에 새로 insert_rows()로 삽입된
        절대 행 번호의 집합(set). 전달하면 rel_row → 절대행 계산 시
        삽입된 행을 건너뛰고 매핑하여, 삽입 지점 아래 행들이 한 칸씩
        어긋나 덮어써지는 문제를 방지합니다. None이면(하위호환) 기존처럼
        'start_row + rel_row' 단순 계산으로 동작합니다.
    """
    inserted_rows = inserted_rows or set()
    header_map = {prefix: row for row, prefix in get_section_header_rows(ws, valid_prefixes)}
    for prefix, tmpl in templates.items():
        if prefix not in header_map: continue
        start_row = header_map[prefix]
        template_end = start_row + tmpl['rel_end_row']
        section_end = get_section_end_row(ws, prefix, valid_prefixes, fallback=template_end)
        headers = get_section_header_rows(ws, valid_prefixes)
        header_idx = next((i for i, (_, p) in enumerate(headers) if p == prefix), None)
        end_row = max(template_end, section_end) if (header_idx is not None and header_idx + 1 < len(headers)) else section_end

        # 이 섹션 범위 안에서 실제로 삽입된 행만 추려서 rel_row 보정에 사용
        section_inserted = {r for r in inserted_rows if start_row <= r <= end_row}
        rel_to_abs = build_rel_row_to_abs_row_map(
            start_row, end_row, section_inserted, tmpl['rel_end_row']
        )

        unmerge_layout_merges_in_range(ws, start_row, end_row)

        for m in tmpl['merges']:
            if section_inserted:
                min_row = rel_to_abs.get(m['rel_min_row'])
                max_row = rel_to_abs.get(m['rel_max_row'])
                if min_row is None or max_row is None:
                    continue
            else:
                min_row = start_row + m['rel_min_row']
                max_row = start_row + m['rel_max_row']
            ws.merge_cells(start_row=min_row, start_column=m['min_col'], end_row=max_row, end_column=m['max_col'])
            anchor = ws.cell(min_row, m['min_col'])
            if m['rel_min_row'] == 0 and m['min_col'] == NAME_COL:
                merge_value = get_prefix_i_label(prefix) or m['value']
            else:
                merge_value = m['value']
            apply_cell_style(anchor, m.get('style'), apply_value=True, value=merge_value)

        for rc in tmpl['row_cells']:
            rel_row, col = rc['rel_row'], rc['col']
            if section_inserted:
                abs_row = rel_to_abs.get(rel_row)
                if abs_row is None:
                    # 이 템플릿 위치에 해당하는 원본 행이 삽입으로 인해 없어진 경우
                    # (정상적으로는 발생하지 않지만 방어적으로 스킵)
                    continue
            else:
                abs_row = start_row + rel_row
            if abs_row > end_row:
                continue
            cell = ws.cell(abs_row, col)
            if isinstance(cell, openpyxl.cell.cell.MergedCell): continue
            if is_layout_col(col):
                apply_cell_style(cell, rc.get('style'), apply_value=(col != NAME_COL), value=rc['value'])
            else:
                apply_cell_style(cell, rc.get('style'), apply_value=False, apply_fill=False)
        # 삽입으로 섹션이 길어진 행: 템플릿 마지막 데이터행 레이아웃을 아래로 확장
        if end_row > template_end:
            _extend_layout_rows_from_template(
                ws, tmpl, start_row, template_end, end_row, valid_prefixes,
            )
        extend_section_i_name_merge_to_end(
            ws, start_row, end_row, prefix=prefix, max_gap=None,
        )
        if finalize_i_label:
            apply_i_section_label(ws, start_row, prefix)

def apply_border_overlay_from_snapshot(ws, master_snapshot, valid_prefixes,
                                       only_prefixes=None, skip_prefixes=None,
                                       inserted_rows=None):
    """행 삽입으로 인해 끊어지거나 지워진 데이터 구역의 격자 테두리 선들을 촘촘히 덧씌워 복원합니다.

    inserted_rows: master_ws에 새로 삽입된 절대 행 번호 집합. 전달 시 rel_row → 절대행
        계산을 삽입행을 건너뛰는 방식으로 보정하여, 테두리가 엉뚱한 행에 덧씌워지는 것을 방지합니다.
    """
    only = set(only_prefixes) if only_prefixes is not None else None
    skip = skip_prefixes or set()
    inserted_rows = inserted_rows or set()
    header_map = {prefix: row for row, prefix in get_section_header_rows(ws, valid_prefixes)}
    for prefix, tmpl in master_snapshot.items():
        if prefix not in header_map:
            continue
        if only is not None and prefix not in only:
            continue
        if prefix in skip:
            continue
        start_row = header_map[prefix]
        template_end = start_row + tmpl['rel_end_row']
        section_end = get_section_end_row(ws, prefix, valid_prefixes, fallback=template_end)
        end_row = max(template_end, section_end)

        section_inserted = {r for r in inserted_rows if start_row <= r <= end_row}
        rel_to_abs = build_rel_row_to_abs_row_map(
            start_row, end_row, section_inserted, tmpl['rel_end_row']
        ) if section_inserted else None

        for rc in tmpl['row_cells']:
            style = rc.get('style')
            if not style or not style.get('border'): continue
            rel_row, col = rc['rel_row'], rc['col']
            if rel_to_abs is not None:
                abs_row = rel_to_abs.get(rel_row)
                if abs_row is None:
                    continue
            else:
                abs_row = start_row + rel_row
            if abs_row > end_row: continue
            cell = ws.cell(abs_row, col)
            if isinstance(cell, openpyxl.cell.cell.MergedCell) or col == INDEX_COL: continue
            if not has_any_border(cell): cell.border = style['border']

def unmerge_layout_merges_in_range(ws, start_row, end_row):
    """레이아웃을 깔끔하게 새로 깔기 위해, 지정한 범위 내의 I~L열 병합을 일단 전부 강제로 풉니다."""
    targets = []
    for mr in ws.merged_cells.ranges:
        if mr.max_col < LAYOUT_COL_START or mr.min_col > LAYOUT_COL_END: continue
        if mr.min_row <= end_row and mr.max_row >= start_row:
            targets.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
    for coords in targets:
        safe_unmerge_cells(ws, *coords)

def extend_section_i_name_merge_to_end(ws, start_row, end_row, prefix=None, max_gap=None):
    """
    I열 사업명 세로 병합을 섹션 end_row까지 확장.
    max_gap=None 이면 행 추가 개수 제한 없음.
    기존 병합이 없으면 start_row~end_row 새로 생성.
    """
    if end_row <= start_row:
        return

    best = None
    for mr in ws.merged_cells.ranges:
        if mr.min_col != NAME_COL or mr.max_col != NAME_COL:
            continue
        if mr.max_row <= mr.min_row:
            continue
        if mr.min_row < start_row or mr.min_row > end_row:
            continue
        if mr.max_row < start_row:
            continue
        priority = (
            1 if mr.min_row <= start_row <= mr.max_row else 0,
            mr.max_row - mr.min_row,
        )
        if best is None or priority > best[4]:
            best = (mr.min_row, mr.min_col, mr.max_row, mr.max_col, priority)

    label = get_prefix_i_label(prefix) if prefix else None

    if best is None:
        anchor = ws.cell(start_row, NAME_COL)
        saved_style = capture_cell_style(anchor)
        saved_value = label or anchor.value
        ws.merge_cells(
            start_row=start_row, start_column=NAME_COL,
            end_row=end_row, end_column=NAME_COL,
        )
        anchor = ws.cell(start_row, NAME_COL)
        if saved_value is not None:
            anchor.value = saved_value
        apply_cell_style(anchor, saved_style, apply_value=False)
        return

    min_row, min_col, max_row, max_col = best[:4]
    anchor = ws.cell(min_row, min_col)
    saved_value = label or anchor.value
    saved_style = capture_cell_style(anchor)

    if max_row >= end_row:
        if saved_value is not None and anchor.value != saved_value:
            anchor.value = saved_value
        return

    gap = end_row - max_row
    if max_gap is not None and gap > max_gap:
        return

    safe_unmerge_cells(ws, min_row, min_col, max_row, max_col)
    ws.merge_cells(
        start_row=min_row, start_column=min_col,
        end_row=end_row, end_column=max_col,
    )
    anchor = ws.cell(min_row, min_col)
    if saved_value is not None:
        anchor.value = saved_value
    apply_cell_style(anchor, saved_style, apply_value=False)


def ensure_all_section_i_name_merges(ws, valid_prefixes, merge_lookup=None):
    '''모든 섹션 I열 병합 확장 + PREFIX_I_LABELS 사업명 재적용 (행 추가 후 호출).'''
    for start_row, prefix in get_section_header_rows(ws, valid_prefixes):
        end_row = get_section_end_row(ws, prefix, valid_prefixes)
        extend_section_i_name_merge_to_end(
            ws, start_row, end_row, prefix=prefix, max_gap=None,
        )
        apply_i_section_label(ws, start_row, prefix, merge_lookup)


def get_row_category_label(ws, row, merge_lookup=None):
    '''J열(병합 포함) 또는 K열에서 섹션 구분 라벨(수입/지출/잔액/발생이자)을 읽습니다.'''
    for col in (J_COL, K_COL):
        raw = ws.cell(row=row, column=col)
        cell = get_real_cell(ws, raw, merge_lookup)
        if cell.value is None:
            continue
        label = str(cell.value).strip()
        if label in CATEGORY_LABELS:
            return label
    return None


def collect_indexed_row_category_labels(ws, indexed_rows, merge_lookup):
    '''
    J 세로 병합을 풀기 전에 각 인덱스 행의 구분 라벨을 수집합니다.
    병합 범위에서 좌상단(값이 들어 있는 셀)이 아닌 행은 J가 비어 보이므로, 반드시 unmerge 전에 호출해야 합니다.
    '''
    labels = {}
    for r in indexed_rows:
        lab = get_row_category_label(ws, r, merge_lookup)
        if lab:
            labels[r] = lab
    return labels


def unmerge_j_vertical_merges_in_range(ws, start_row, end_row):
    '''섹션 범위 내 J열 단독 세로 병합만 해제 (J:K 가로 병합은 유지).'''
    targets = []
    for mr in ws.merged_cells.ranges:
        if mr.min_col != J_COL or mr.max_col != J_COL:
            continue
        if mr.min_row <= end_row and mr.max_row >= start_row:
            targets.append((mr.min_row, mr.min_col, mr.max_row, mr.max_col))
    for coords in targets:
        safe_unmerge_cells(ws, *coords)


def rebuild_j_category_merges_in_section(ws, start_row, end_row, prefix, prefix_pattern,
                                         merge_lookup=None):
    '''
    같은 수입/지출/잔액/발생이자 블록의 2번째 행~마지막 행에 J열 세로 병합을 재구성합니다.

    원본 양식: 블록 첫 행은 J:K 가로 병합, 이어지는 행들은 J열만 세로 병합(예: J242:J244=수입).
    행 추가 시 새 행이 같은 카테고리면 J 병합 범위를 자동으로 늘립니다.
    '''
    if merge_lookup is None:
        merge_lookup = build_merge_lookup(ws)

    indexed_rows = []
    for r in range(start_row + 1, end_row + 1):
        h = ws.cell(row=r, column=INDEX_COL).value
        if h is None:
            continue
        idx_str = str(h).strip()
        match = prefix_pattern.match(idx_str)
        if match and match.group(1) == prefix:
            indexed_rows.append(r)

    if len(indexed_rows) < 2:
        return

    row_labels = collect_indexed_row_category_labels(ws, indexed_rows, merge_lookup)

    unmerge_j_vertical_merges_in_range(ws, start_row, end_row)

    blocks = []
    for r in indexed_rows:
        label = row_labels.get(r)
        if not label:
            continue
        if blocks and blocks[-1][0] == label:
            blocks[-1][1].append(r)
        else:
            blocks.append([label, [r]])

    for label, rows in blocks:
        if len(rows) < 2:
            continue
        rest = rows[1:]
        merge_start = rest[0]
        merge_end = rows[-1]
        if merge_end <= merge_start:
            continue
        style_src = get_real_cell(ws, ws.cell(merge_start, J_COL), merge_lookup)
        saved_style = capture_cell_style(style_src)
        ws.merge_cells(
            start_row=merge_start, start_column=J_COL,
            end_row=merge_end, end_column=J_COL,
        )
        anchor = ws.cell(merge_start, J_COL)
        apply_cell_style(anchor, saved_style, apply_value=True, value=label)


def rebuild_all_j_category_merges(ws, valid_prefixes, prefix_pattern, merge_lookup=None):
    '''모든 부서 섹션에서 J열 수입/지출/잔액 세로 병합을 재구성합니다.'''
    if merge_lookup is None:
        merge_lookup = build_merge_lookup(ws)
    headers = get_section_header_rows(ws, valid_prefixes)
    for i, (start_row, prefix) in enumerate(headers):
        if i + 1 < len(headers):
            end_row = headers[i + 1][0] - 1
        else:
            end_row = find_section_content_end_row(ws, start_row)
        rebuild_j_category_merges_in_section(
            ws, start_row, end_row, prefix, prefix_pattern, merge_lookup,
        )


def apply_row_style_from_neighbor(ws, target_row, source_row, max_col):
    """새로 생성된 하얀 빈 행(target_row)에 바로 위/아래 이웃 행(source_row)의 서식(글꼴, 테두리)을 그대로 복제해 자연스럽게 만듭니다."""
    for c in range(1, max_col + 1):
        if is_layout_col(c): continue
        tgt_raw = ws.cell(row=target_row, column=c)
        if isinstance(tgt_raw, openpyxl.cell.cell.MergedCell): continue
        src_cell = ws.cell(row=source_row, column=c)
        copy_cell_style_safe(src_cell, tgt_raw)


def copy_dept_layout_values_for_rows(master_ws, fix_remap_queue, new_rows_buffer, master_map,
                                     max_col, master_merge_lookup=None):
    '''레이아웃 병합·색상 복원 후, 부서 I~K열 사용자 입력값 덮어씀 (L열~는 dept_apply_data_overlays).'''
    pairs = _dept_remap_new_row_pairs(fix_remap_queue, new_rows_buffer, master_map)
    for master_row, dept_row, dept_ws, dept_merge_lookup in pairs:
        for c in range(LAYOUT_COL_START, min(max_col, K_COL) + 1):
            d_cell = get_real_cell(dept_ws, dept_ws.cell(dept_row, c), dept_merge_lookup)
            if d_cell.value is None:
                continue
            m_raw = master_ws.cell(master_row, c)
            m_cell = get_real_cell(master_ws, m_raw, master_merge_lookup)
            if isinstance(m_cell, openpyxl.cell.cell.MergedCell):
                continue
            m_cell.value = d_cell.value


def reapply_index_marker_fills(ws, index_values, fill, index_map=None):
    """레이아웃 복원 후 H열 연두색 마커 재적용 (index_map 재사용 가능)."""
    if index_map is None:
        index_map = get_sheet_index_map(ws, include_name=False)
    for idx_str in index_values:
        info = index_map.get(idx_str)
        if info:
            ws.cell(info['row_num'], INDEX_COL).fill = fill

# -----------------------------------------
# 6. 3-Way Diff 예외상황 판단 도구
# - 

def is_section_header_index(idx_str, valid_prefixes):
    """'DB2'처럼 뒤에 숫자가 안 붙은 대분류 제목용 단독 인덱스인지 확인합니다."""
    return idx_str in valid_prefixes

def find_index_below_row(ws, row_num):
    """현재 행 아래쪽을 훑어 내려가며 가장 가깝게 만나는 H열 인덱스 글자를 찾습니다."""
    for br in range(row_num + 1, ws.max_row + 1):
        v = ws.cell(row=br, column=INDEX_COL).value
        if v is not None and str(v).strip(): return str(v).strip()
    return None

def master_has_missing_index_in_gap(master_map, anchor_idx_str, below_idx_str, prefix_pattern):
    """부서 파일의 빈 줄이 새로 만든 신규 행인지, 아니면 기존에 있던 행을 삭제해서 빈 것인지 번호 차이(Gap)를 통해 정확히 추론합니다."""
    if not anchor_idx_str or not below_idx_str: return False
    a_match = prefix_pattern.match(anchor_idx_str)
    b_match = prefix_pattern.match(below_idx_str)
    if not a_match or not b_match or a_match.group(1) != b_match.group(1): return False
    a_num, b_num = int(a_match.group(2)), int(b_match.group(2))
    if b_num <= a_num + 1: return False
    prefix = a_match.group(1)
    for n in range(a_num + 1, b_num):
        if f'{prefix}-{n}' in master_map: return True # 취합용 파일에만 인덱스가 존재한다면 부서에서 '삭제'한 행입니다.
    return False

def dept_confirms_index_deletion(idx_str, dept_map, master_map, prefix_pattern, valid_prefixes):
    '''부서 파일에서 H만 지우고 위·아래 번호가 이어지면 삭제 확정 (DB2-2 등).'''
    if idx_str in dept_map:
        return False
    match = prefix_pattern.match(idx_str)
    if not match:
        return False
    prefix = match.group(1)
    if not is_allowed_index_prefix(prefix, master_map, valid_prefixes):
        return False
    dept_indexes = set(dept_map.keys())
    num = int(match.group(2))
    prev_idx = f'{prefix}-{num - 1}' if num > 1 else None
    next_idx = f'{prefix}-{num + 1}'
    has_prev = prev_idx in dept_indexes if prev_idx else prefix in dept_indexes
    has_next = next_idx in dept_indexes
    if not (has_prev and has_next):
        return False
    anchor = prev_idx if prev_idx else prefix
    return master_has_missing_index_in_gap(
        master_map, anchor, next_idx, prefix_pattern
    )


def collect_confirmed_deletions(master_index_list, all_dept_indexes):
    """
    마스터 인덱스 리스트와 부서 파일 수집 인덱스 세트를 비교하여
    부서 파일에서 사라진 모든 인덱스를 정확하게 찾아 반환합니다.
    """
    # 1. 마스터 리스트에서 None을 제외한 유효한 인덱스들만 세트로 변환
    master_indexes_set = set(idx for idx in master_index_list if idx is not None)
    
    # 2. 차집합을 통해 마스터에는 있지만 부서에는 없는 인덱스를 추출 (마지막 행 여부 상관없이 포착)
    confirmed_deletions = master_indexes_set - all_dept_indexes
    
    return confirmed_deletions

def find_blank_index_rows(dept_ws, prefix_pattern, master_map=None):
    """
    H열이 비어 있고 다른 열에 값/수정(파란색)이 있는 행을 신규 후보로 찾는다.

    삭제행(H만 지우고 나머지 데이터 유지)과 구분:
    - 위·아래 인덱스 사이 gap에 마스터 전용 번호(DB2-2 등)가 있으면 → 삭제, 신규 제외
    - gap 없음 + 데이터 있음 → 신규
    """
    blank_rows = []
    valid_prefixes = get_valid_prefixes()
    for r in range(1, dept_ws.max_row + 1):
        idx_val = dept_ws.cell(row=r, column=INDEX_COL).value
        if idx_val is not None and str(idx_val).strip():
            continue

        has_data = False
        for c in range(1, dept_ws.max_column + 1):
            if c == INDEX_COL:
                continue
            cell = dept_ws.cell(row=r, column=c)
            if cell.value is not None or is_blue_cell(cell):
                has_data = True
                break
        if not has_data:
            continue

        section_prefix = None
        anchor_idx_str = None
        for br in range(r - 1, 0, -1):
            v = dept_ws.cell(row=br, column=INDEX_COL).value
            if v is not None and str(v).strip():
                v_str = str(v).strip()
                m = prefix_pattern.match(v_str)
                if m:
                    section_prefix = m.group(1)
                    anchor_idx_str = v_str
                break
        if section_prefix is None:
            for br in range(r + 1, dept_ws.max_row + 1):
                v = dept_ws.cell(row=br, column=INDEX_COL).value
                if v is not None and str(v).strip():
                    v_str = str(v).strip()
                    m = prefix_pattern.match(v_str)
                    if m:
                        section_prefix = m.group(1)
                    break
        if section_prefix not in valid_prefixes:
            continue

        if master_map is not None and anchor_idx_str:
            below_idx_str = find_index_below_row(dept_ws, r)
            if master_has_missing_index_in_gap(
                master_map, anchor_idx_str, below_idx_str, prefix_pattern
            ):
                continue

        blank_rows.append((r, section_prefix, anchor_idx_str))
    return blank_rows

def is_blank_temp_index(idx_str):
    """임시로 생성된 가짜 블랭크 식별용 텍스트 키인지 검사합니다."""
    return '-BLANK-' in idx_str

def alloc_temp_index(prefix, temp_counters):
    """H열이 비어있던 신규행에 프로그램 내부 처리를 위해 가상의 임시 번호(예: DB2-T1)를 부여합니다."""
    temp_counters[prefix] = temp_counters.get(prefix, 0) + 1
    return f'{prefix}-T{temp_counters[prefix]}'

def translate_dept_formula(value, origin_coord, master_row, dept_row):
    '''부서 행→마스터 행 복사 시 수식 내 행 참조를 row_delta만큼 보정.'''
    if not isinstance(value, str) or not value.startswith('='):
        return value
    row_delta = master_row - dept_row
    if row_delta == 0:
        return value
    try:
        return Translator(value, origin=origin_coord).translate_formula(
            row_delta=row_delta, col_delta=0,
        )
    except Exception:
        return value


def copy_row_values_and_color(master_ws, dept_ws, master_row, dept_row, max_col, fill=None,
                              index_value=None, master_merge_lookup=None, dept_merge_lookup=None,
                              force_copy_row=False, data_cols_only=False):
    """부서 파일 행 → 마스터 행 값·파란색 배경 복사.

    data_cols_only=True 이면 I~K layout 전용 열은 건너뛰고 H(인덱스)와 L열~(항목·숫자·파란 수정)을 복사.
    L열은 layout 복원 후 dept_apply_data_overlays / layout 후 blue 재적용으로 최종 반영.
    """
    blue_fill = PatternFill(
        start_color=TARGET_BLUE_CODES[1],
        end_color=TARGET_BLUE_CODES[1],
        fill_type='solid',
    )
    copy_layout = (force_copy_row or fill is not None) and not data_cols_only
    for c in range(1, max_col + 1):
        m_cell_raw = master_ws.cell(row=master_row, column=c)
        d_cell_raw = dept_ws.cell(row=dept_row, column=c)
        d_cell = get_real_cell(dept_ws, d_cell_raw, dept_merge_lookup)
        is_blue = is_blue_cell(d_cell_raw) or is_blue_cell(d_cell)

        if c == INDEX_COL:
            if not isinstance(m_cell_raw, openpyxl.cell.cell.MergedCell):
                if index_value is not None:
                    m_cell_raw.value = index_value
                if fill is not None:
                    m_cell_raw.fill = fill
            continue

        m_cell = get_real_cell(master_ws, m_cell_raw, master_merge_lookup)
        if isinstance(m_cell, openpyxl.cell.cell.MergedCell):
            continue

        if is_layout_only_col(c):
            if data_cols_only:
                continue
            if is_blue or copy_layout:
                if d_cell.value is not None or copy_layout:
                    m_cell.value = d_cell.value
                if is_blue:
                    m_cell.fill = blue_fill
                elif copy_layout and d_cell.fill and d_cell.fill.fill_type:
                    m_cell.fill = copy.copy(d_cell.fill)
            continue

        if not (d_cell.value is not None or is_blue or force_copy_row):
            if data_cols_only:
                continue
        m_cell.value = translate_dept_formula(
            d_cell.value, d_cell_raw.coordinate, master_row, dept_row,
        )
        if is_blue:
            m_cell.fill = blue_fill
        elif force_copy_row and d_cell.fill and d_cell.fill.fill_type:
            m_cell.fill = copy.copy(d_cell.fill)

"""
DashBoard 시트 — 차트 위 텍스트 상자 추출 · 전처리 · 부서 diff · 취합본 저장.
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from contextlib import contextmanager

import pandas as pd
import win32com.client as win32
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------

DASHBOARD_SHEET_PREFIXES = ('DashBoard', 'Dashboard')
# openpyxl 저장 전 임시 제외 → 저장 후 COM으로 원본에서 복원
STRIP_DASHBOARD_SHEET_NAMES = (
    'DashBoard1', 'Dashboard1', 'DashBoard2', 'Dashboard2',
)
# COM Copy 후 외부 통합문서 참조 → 같은 파일 내부 시트로
#  1) ='C:\path\[file.xlsx]Sheet'!A1  → ='Sheet'!A1
#  2) ='[1]Sheet'!A1 / =[1]Sheet!A1   → ='Sheet'!A1 / =Sheet!A1
_COM_FILE_EXT_RE = re.compile(
    r"'[^']*\[(?:[^\]]+\.(?:xlsx|xlsm|xlsb|xls))\]([^']+)'!",
    re.IGNORECASE,
)
_COM_INDEX_EXT_RE = re.compile(r'\[\d+\](?=[^!\]]+!)')
_PARAMETER_DASHBOARD_MATCH = {
    # 따옴표로 로컬 시트 강제 — LinkSources에 'DashBoard2' 잔존 시
    # DashBoard2!F30 이 외부링크로 묶여 #REF!가 나는 것을 막음
    'A6': "=MATCH('DashBoard2'!F30,A3:A5,0)",
    'B6': "=MATCH('DashBoard2'!F31,B3:B5,0)",
    'C6': "=MATCH('DashBoard2'!F32,C3:C5,0)",
    'D6': "=MATCH('DashBoard2'!F33,D3:D5,0)",
    'E6': "=MATCH('DashBoard2'!F34,E3:E5,0)",
}
UNIT_LABEL = '(단위 : 백만원)'
# '(단위 : 백만원)', '(단위:백만원)' 등 표기 변형
UNIT_LABEL_RE = re.compile(r'단위\s*[:：]?\s*백만')
STATUS_EXISTING = '기존'
STATUS_DEPT_ADD = '부서추가'
FILL_DEPT_ADD = PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
HEADER_FONT = Font(bold=True)

OUTPUT_COLUMNS = [
    'status', 'source_file', 'chart_title', 'chart_name', 'text',
]

# 차트 위 섹션 제목 상자 식별용 (예: "□ 운영자본금", "ㅁ 사업준비금")
SECTION_HEADER_MARKERS = '□■▢▣◻◼◽◾▪▫●○◎◇◆❑❒ㅁ※·-'
# 섹션 제목은 날짜·금액 숫자가 없는 짧은 텍스트 (주석은 숫자 포함)
SECTION_TITLE_MAX_LEN = 30
# 대시보드 큰 섹션명 (셀/도형에서 우선 매칭)
SECTION_NAME_KEYWORDS = ('운영자본금', '사업준비금', '부채')

# ---------------------------------------------------------------------------
# pywin32 — 텍스트 상자 추출
# ---------------------------------------------------------------------------


def _shape_text(shape) -> str:
    for attr in ('TextFrame2', 'TextFrame'):
        frame = getattr(shape, attr, None)
        if frame is None:
            continue
        try:
            if hasattr(frame, 'HasText') and not frame.HasText:
                continue
            text = frame.TextRange.Text
            if text:
                return text.replace('\r\n', '\n').replace('\r', '\n').strip()
        except Exception:
            continue
    return ''


def _iter_shapes(container):
    try:
        count = container.Shapes.Count
    except Exception:
        return
    for i in range(1, count + 1):
        yield container.Shapes(i)


def _chart_title(chart_obj) -> str:
    try:
        chart = chart_obj.Chart
        if chart.HasTitle:
            return str(chart.ChartTitle.Text or '').strip()
    except Exception:
        pass
    return chart_obj.Name


def _is_unit_label(text: str) -> bool:
    """단위 표기('(단위 : 백만원)' 등) 여부."""
    return bool(UNIT_LABEL_RE.search(str(text or '')))


def _cell_section_title_above_chart(ws, chart_obj) -> str:
    """
    차트 바로 위 셀에서 섹션명(□ 운영자본금 등)을 읽음.
    대시보드 섹션 제목이 도형이 아니라 셀에 있는 경우가 많음.
    """
    try:
        tl = chart_obj.TopLeftCell
        start_row = max(1, int(tl.Row) - 4)
        end_row = int(tl.Row)
        start_col = max(1, int(tl.Column) - 1)
        end_col = int(tl.Column) + 8
    except Exception:
        return ''

    candidates = []
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            try:
                val = ws.Cells(r, c).Value
            except Exception:
                continue
            if val is None:
                continue
            text = str(val).strip()
            if not text or _is_unit_label(text):
                continue
            if _looks_like_section_header(text) or _match_section_keyword(text):
                # 차트에 더 가까운 행을 우선
                candidates.append((end_row - r, c, text))
    if not candidates:
        return ''
    candidates.sort()
    return _clean_section_title(candidates[0][2])


def extract_dashboard_data(excel_path, sheet_name, excel_app=None, include_chart_shapes=True):
    excel_path = os.path.abspath(excel_path)
    if not os.path.isfile(excel_path):
        raise FileNotFoundError(excel_path)

    own_excel = excel_app is None
    if own_excel:
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

    wb = None
    charts = []
    text_boxes = []

    try:
        wb = excel_app.Workbooks.Open(excel_path, ReadOnly=True)
        ws = wb.Worksheets(sheet_name)

        for chart_obj in ws.ChartObjects():
            charts.append({
                'name': chart_obj.Name,
                'title': _chart_title(chart_obj),
                'cell_title': _cell_section_title_above_chart(ws, chart_obj),
                'top': float(chart_obj.Top),
                'left': float(chart_obj.Left),
                'width': float(chart_obj.Width),
                'height': float(chart_obj.Height),
            })

        for shape in _iter_shapes(ws):
            text = _shape_text(shape)
            if not text:
                continue
            text_boxes.append({
                'name': shape.Name,
                'text': text,
                'top': float(shape.Top),
                'left': float(shape.Left),
                'width': float(shape.Width),
                'height': float(shape.Height),
                'source': 'sheet',
            })

        if include_chart_shapes:
            for chart_obj in ws.ChartObjects():
                chart = chart_obj.Chart
                for shape in _iter_shapes(chart):
                    text = _shape_text(shape)
                    if not text:
                        continue
                    text_boxes.append({
                        'name': shape.Name,
                        'text': text,
                        'top': float(chart_obj.Top + shape.Top),
                        'left': float(chart_obj.Left + shape.Left),
                        'width': float(shape.Width),
                        'height': float(shape.Height),
                        'source': 'chart',
                        'chart_object': chart_obj.Name,
                    })

        text_boxes.sort(key=lambda x: (x['top'], x['left']))
        charts.sort(key=lambda x: (x['top'], x['left']))
        return {'charts': charts, 'text_boxes': text_boxes}
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if own_excel:
            excel_app.Quit()


@contextmanager
def excel_session():
    app = win32.DispatchEx('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    try:
        yield app
    finally:
        app.Quit()


def discover_dashboard_sheets(excel_path, excel_app=None):
    """워크북에서 DashBoard/Dashboard 로 시작하는 시트명 목록."""
    excel_path = os.path.abspath(excel_path)
    own_excel = excel_app is None
    if own_excel:
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
    wb = None
    try:
        wb = excel_app.Workbooks.Open(excel_path, ReadOnly=True)
        names = []
        for i in range(1, wb.Worksheets.Count + 1):
            name = wb.Worksheets(i).Name
            if name.startswith(DASHBOARD_SHEET_PREFIXES):
                names.append(name)
        return sorted(names)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if own_excel:
            excel_app.Quit()


# ---------------------------------------------------------------------------
# 전처리 · 정렬
# ---------------------------------------------------------------------------


def normalize_text(text: str) -> str:
    if not isinstance(text, str):
        return ''
    t = re.sub(r'\s+', ' ', text.strip())
    return t


def extract_sort_key(text: str) -> str:
    cleaned = re.sub(r'[()`]', '', str(text))
    match = re.match(r'^([\d\.]+년?[\s\d\w]*)', cleaned)
    if not match:
        return '99.99'
    date_part = match.group(1).strip().replace('월', '').replace('년', '')
    parts = date_part.split('.')
    if len(parts) == 2 and len(parts[1]) == 1 and parts[1].isdigit():
        return f'{parts[0]}.0{parts[1]}'
    return date_part


def _flatten_shape_text(text: str) -> str:
    """도형 하나의 텍스트에서 줄바꿈을 없애고 한 줄로 합침 (단위행 제외)."""
    lines = [ln.strip() for ln in str(text).split('\n')]
    lines = [ln for ln in lines if ln and not _is_unit_label(ln) and ln != UNIT_LABEL]
    return ' '.join(lines)


def process_textboxes_df(df: pd.DataFrame) -> pd.DataFrame:
    """도형별 줄바꿈 제거(한 셀) · 단위행 제거 · chart+날짜 정렬."""
    if df.empty:
        return df.copy()

    out = df.copy()
    out['text'] = out['text'].astype(str).map(_flatten_shape_text)
    out = out[out['text'] != ''].reset_index(drop=True)
    out['text_norm'] = out['text'].map(normalize_text)
    out['sort_key'] = out['text'].map(extract_sort_key)

    sort_cols = ['chart_name', 'sort_key', 'text']
    for col in sort_cols:
        if col not in out.columns:
            out[col] = ''
    out = out.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)
    return out


def textboxes_to_dataframe(text_boxes: list[dict]) -> pd.DataFrame:
    if not text_boxes:
        return pd.DataFrame(columns=OUTPUT_COLUMNS + ['text_norm', 'sort_key', 'name'])
    return pd.DataFrame(text_boxes)


def extract_and_process(excel_path, sheet_name, excel_app=None) -> pd.DataFrame:
    data = extract_dashboard_data(excel_path, sheet_name, excel_app=excel_app)
    boxes = group_text_boxes_by_chart(data['text_boxes'], data['charts'])
    df = textboxes_to_dataframe(boxes)
    return process_textboxes_df(df)


# ---------------------------------------------------------------------------
# 차트 그룹 · diff · 취합본
# ---------------------------------------------------------------------------


def _box_center(item):
    return (
        item['top'] + item['height'] / 2,
        item['left'] + item['width'] / 2,
    )


def _in_bounds(cy, cx, region, margin=12):
    return (
        region['top'] - margin <= cy <= region['top'] + region['height'] + margin
        and region['left'] - margin <= cx <= region['left'] + region['width'] + margin
    )


def _clean_section_title(text: str) -> str:
    """섹션 제목 상자/셀 텍스트에서 머리기호를 떼고 실제 제목만 반환."""
    first = str(text).split('\n')[0].strip()
    return first.lstrip(SECTION_HEADER_MARKERS + ' \t').strip()


def _match_section_keyword(text: str) -> str:
    """알려진 섹션 키워드가 포함되면 정규화된 섹션명 반환."""
    core = _clean_section_title(text)
    if not core or _is_unit_label(core):
        return ''
    for kw in SECTION_NAME_KEYWORDS:
        if kw in core:
            # '부채(공사채) 발행잔고'처럼 키워드를 포함한 전체 제목 유지
            return core
    return ''


def _looks_like_section_header(text: str) -> bool:
    """
    '□ 운영자본금'·'ㅁ 사업준비금'·'부채(공사채) 발행잔고' 등 섹션 제목 판별.
    '(단위 : 백만원)' 및 날짜/금액 주석은 제외.
    """
    s = str(text).strip()
    if not s or _is_unit_label(s):
        return False
    if _match_section_keyword(s):
        return True
    core = _clean_section_title(s)
    if not core or _is_unit_label(core):
        return False
    if core.startswith('('):
        return False
    if any(ch.isdigit() for ch in core):
        return False
    return len(core) <= SECTION_TITLE_MAX_LEN


def _pick_chart_section_title(chart, shape_title: str) -> str:
    """셀 제목 → 도형 제목 → (단위 제외) 차트 제목 순으로 섹션명 선택."""
    for cand in (
        chart.get('cell_title', ''),
        shape_title,
        chart.get('title', ''),
    ):
        if not cand or _is_unit_label(cand):
            continue
        kw = _match_section_keyword(cand)
        if kw:
            return kw
        if _looks_like_section_header(cand):
            return _clean_section_title(cand)
    # 차트 객체명(Chart 1 등)은 쓰지 않음 — 단위/빈 값 방지
    return ''


def _assign_section_titles(text_boxes, charts):
    """
    차트별 실제 섹션 제목 결정.
    우선순위: 차트 위 셀 제목 → 근처 섹션 도형 → 차트 제목(단위 제외).
    """
    headers = [it for it in text_boxes if _looks_like_section_header(it['text'])]
    best = {}  # ci -> (우선순위점수, 거리, 제목) 점수 낮을수록 좋음
    for h in headers:
        ci = min(range(len(charts)), key=lambda i: abs(charts[i]['top'] - h['top']))
        dist = abs(charts[ci]['top'] - h['top'])
        title = _match_section_keyword(h['text']) or _clean_section_title(h['text'])
        # 알려진 섹션 키워드가 있으면 우선
        score = 0 if _match_section_keyword(h['text']) else 1
        prev = best.get(ci)
        if prev is None or (score, dist) < (prev[0], prev[1]):
            best[ci] = (score, dist, title)

    titles = {}
    for ci, ch in enumerate(charts):
        shape_title = best[ci][2] if ci in best else ''
        titles[ci] = _pick_chart_section_title(ch, shape_title)
    return titles


def group_text_boxes_by_chart(text_boxes, charts):
    if not charts:
        for item in text_boxes:
            item['chart_title'] = ''
            item['chart_name'] = ''
        return text_boxes

    for item in text_boxes:
        cy, cx = _box_center(item)
        matched = [i for i, ch in enumerate(charts) if _in_bounds(cy, cx, ch)]
        if len(matched) == 1:
            ci = matched[0]
        elif matched:
            ci = min(matched, key=lambda i: abs(cy - (charts[i]['top'] + charts[i]['height'] / 2)))
        else:
            ci = min(range(len(charts)), key=lambda i: abs(cy - (charts[i]['top'] + charts[i]['height'] / 2)))
        item['_chart_idx'] = ci

    chart_titles = _assign_section_titles(text_boxes, charts)

    for item in text_boxes:
        ci = item.pop('_chart_idx')
        title = chart_titles.get(ci, '')
        item['chart_title'] = title
        item['chart_name'] = title
    return text_boxes


def _row_key(chart_name, text_norm) -> tuple:
    return (str(chart_name or ''), str(text_norm or ''))


def find_dept_additions(master_df: pd.DataFrame, dept_df: pd.DataFrame) -> pd.DataFrame:
    """취합용에 없고 부서에만 있는 (chart_name, text) 조합."""
    if dept_df.empty:
        return dept_df.copy()
    master_keys = {
        _row_key(r.chart_name, r.text_norm)
        for r in master_df.itertuples(index=False)
    } if not master_df.empty else set()

    mask = []
    for r in dept_df.itertuples(index=False):
        mask.append(_row_key(r.chart_name, r.text_norm) not in master_keys)
    return dept_df.loc[mask].copy()


def build_consolidated(master_df: pd.DataFrame, dept_results: list[tuple[str, pd.DataFrame]]) -> pd.DataFrame:
    """
    취합본 테이블: 기존(취합용) + 부서추가(빨간 표시 대상).
    dept_results: [(source_label, processed_df), ...]
    """
    rows = []

    if not master_df.empty:
        for r in master_df.itertuples(index=False):
            rows.append({
                'status': STATUS_EXISTING,
                'source_file': '취합용',
                'chart_title': getattr(r, 'chart_title', ''),
                'chart_name': getattr(r, 'chart_name', ''),
                'text': r.text,
                'text_norm': r.text_norm,
                'sort_key': r.sort_key,
            })

    seen_add_keys = set()
    for source_label, dept_df in dept_results:
        additions = find_dept_additions(master_df, dept_df)
        for r in additions.itertuples(index=False):
            key = _row_key(r.chart_name, r.text_norm)
            if key in seen_add_keys:
                continue
            seen_add_keys.add(key)
            rows.append({
                'status': STATUS_DEPT_ADD,
                'source_file': source_label,
                'chart_title': getattr(r, 'chart_title', ''),
                'chart_name': getattr(r, 'chart_name', ''),
                'text': r.text,
                'text_norm': r.text_norm,
                'sort_key': r.sort_key,
            })

    if not rows:
        return pd.DataFrame(columns=OUTPUT_COLUMNS + ['text_norm', 'sort_key'])

    out = pd.DataFrame(rows)
    out = out.sort_values(
        by=['chart_name', 'sort_key', 'status', 'text'],
        ascending=True,
    ).reset_index(drop=True)
    return out


def _dept_label(dept_path: str) -> str:
    base = os.path.splitext(os.path.basename(dept_path))[0]
    for token in ('_부서명', '_부서', '분석자료'):
        base = base.replace(token, '')
    return base.strip() or os.path.basename(dept_path)


def _sanitize(val):
    # openpyxl은 스칼라(str/숫자/bool/None)만 허용. 그 외(튜플/리스트/객체 등)는 문자열로 강제.
    if val is None or isinstance(val, (int, float, bool)):
        return val
    return re.sub(ILLEGAL_CHARACTERS_RE, '', str(val))

def save_consolidated_xlsx(df: pd.DataFrame, xlsx_path: str):
    """취합 데이터 + 부서추가 행 빨간 배경."""
    export_cols = [c for c in OUTPUT_COLUMNS if c in df.columns]
    os.makedirs(os.path.dirname(os.path.abspath(xlsx_path)), exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = '취합본'

    for col_idx, col_name in enumerate(export_cols, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT

    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        is_add = getattr(row, 'status', '') == STATUS_DEPT_ADD
        for col_idx, col_name in enumerate(export_cols, 1):
            val = _sanitize(getattr(row, col_name, ''))
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if is_add:
                cell.fill = FILL_DEPT_ADD

    for col_idx, col_name in enumerate(export_cols, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(40, max(12, len(col_name) + 2))

    wb.save(xlsx_path)


def aggregate_dashboard_consolidation(
    master_path,
    dept_paths,
    output_dir,
    sheet_names=None,
    ):
    """
    취합용 + 부서 DashBoard 텍스트 추출·diff·취합본 XLSX 저장.

    Returns:
        dict[sheet_name] -> {'xlsx': path, 'add_count': int}
    """
    master_path = os.path.abspath(master_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    results = {}
    with excel_session() as excel_app:
        if sheet_names is None:
            sheet_names = discover_dashboard_sheets(master_path, excel_app=excel_app)
        if not sheet_names:
            print('  DashBoard 시트가 없습니다 — 건너뜀.')
            return results

        for sheet_name in sheet_names:
            print(f'  [{sheet_name}] DashBoard 텍스트 추출·취합...')
            try:
                master_df = extract_and_process(master_path, sheet_name, excel_app=excel_app)
            except Exception as e:
                print(f'    취합용 추출 실패: {e}')
                continue

            dept_results = []
            for dept_path in dept_paths:
                label = _dept_label(dept_path)
                try:
                    dept_df = extract_and_process(dept_path, sheet_name, excel_app=excel_app)
                    dept_results.append((label, dept_df))
                except Exception as e:
                    print(f'    부서 {label} 스킵 ({e})')

            consolidated = build_consolidated(master_df, dept_results)
            add_count = int((consolidated['status'] == STATUS_DEPT_ADD).sum()) if not consolidated.empty else 0

            safe_sheet = re.sub(r'[\\/:*?"<>|]', '_', sheet_name)
            folder_file = os.path.join(output_dir, '최종취합본')
            os.makedirs(folder_file, exist_ok=True)
            xlsx_path = os.path.join(folder_file, f'dashboard_{safe_sheet}_취합본.xlsx')

            save_consolidated_xlsx(consolidated, xlsx_path)

            print(f'    기존 {len(master_df)}건 / 부서추가 {add_count}건')
            print(f'    XLSX → {xlsx_path} (부서추가=빨간 배경)')

            results[sheet_name] = {
                'xlsx': xlsx_path,
                'add_count': add_count,
                'rows': len(consolidated),
            }

    return results


def strip_dashboard_sheets(workbook, sheet_names=None) -> list[str]:
    """
    openpyxl 워크북에서 DashBoard1/2 시트를 제거한다.

    openpyxl이 차트/도형이 있는 시트를 저장하면 xl/drawings 관계가 깨져
    Excel이 '내용을 복구하시겠습니까?' / Open 오류를 낸다.
    저장 전에 빼고, 저장 후 COM으로 취합용 원본을 붙인 뒤
    Copy가 만든 [n]Sheet! 외부참조를 내부참조로 되돌린다.
    """
    targets = set(sheet_names or STRIP_DASHBOARD_SHEET_NAMES)
    removed = [
        name for name in list(workbook.sheetnames)
        if name in targets
    ]
    for name in removed:
        del workbook[name]
    # 제외한 시트 때문에 남은 시트에 깨진 차트/이미지 참조가 있으면 제거
    for ws in workbook.worksheets:
        if hasattr(ws, '_charts'):
            ws._charts = []
        if hasattr(ws, '_images'):
            ws._images = []
    return removed


def _localize_external_formula(formula: str) -> str | None:
    """외부 통합문서 참조 수식을 내부 시트 참조로 바꾼다. 변경 없으면 None."""
    if not (isinstance(formula, str) and formula.startswith('=')):
        return None
    new_f = _COM_FILE_EXT_RE.sub(r"'\1'!", formula)
    new_f = _COM_INDEX_EXT_RE.sub('', new_f)
    return new_f if new_f != formula else None


def _com_localize_external_sheet_refs(workbook, sheet_names=None) -> int:
    """
    COM 워크북 수식의 외부 통합문서 참조를 같은 파일 내부 시트로 바꾼다.

    DashBoard를 다른 통합문서에서 Copy하면 Excel이 시트 참조를
    'C:\\...\\[취합용.xlsx]취합(2606)'! 형태로 바꿔, 값이 원본 파일을
    따라가거나 링크가 없으면 깨진다.
    """
    if sheet_names is None:
        names = [
            workbook.Worksheets(i).Name
            for i in range(1, workbook.Worksheets.Count + 1)
        ]
    else:
        names = list(sheet_names)

    fixed = 0
    for name in names:
        try:
            ws = workbook.Worksheets(name)
        except Exception:
            continue
        try:
            used = ws.UsedRange
            row0, col0 = int(used.Row), int(used.Column)
        except Exception:
            continue

        try:
            formulas = used.Formula
        except Exception:
            continue
        if formulas is None:
            continue
        if not isinstance(formulas, tuple):
            formulas = ((formulas,),)

        for i, row in enumerate(formulas):
            if not isinstance(row, tuple):
                row = (row,)
            for j, formula in enumerate(row):
                new_f = _localize_external_formula(formula)
                if new_f is None:
                    continue
                ws.Cells(row0 + i, col0 + j).Formula = new_f
                fixed += 1
    return fixed


def _com_ensure_parameter_dashboard_links(workbook) -> int:
    """Parameter!A6:E6이 로컬 DashBoard2 시트를 참조하도록 고정 (제외·복원 후 #REF! 방지)."""
    try:
        ws = workbook.Worksheets('Parameter')
    except Exception:
        return 0
    fixed = 0
    for addr, formula in _PARAMETER_DASHBOARD_MATCH.items():
        cell = ws.Range(addr)
        try:
            cur = cell.Formula
        except Exception:
            cur = None
        if cur != formula:
            cell.Formula = formula
            fixed += 1
    return fixed


def _com_drop_phantom_dashboard_links(workbook) -> int:
    """
    COM Copy 과정에서 생긴 'DashBoard2' 이름의 가짜 외부 링크를 제거한다.

    시트 DashBoard2와 링크 이름 DashBoard2가 동시에 있으면
    Parameter의 MATCH(DashBoard2!F30)이 시트가 아니라 깨진 외부링크로
    묶여 #REF!가 난다. 수식을 로컬화한 뒤 해당 링크만 BreakLink.
    """
    try:
        links = workbook.LinkSources(1)  # xlExcelLinks
    except Exception:
        return 0
    if not links:
        return 0
    if isinstance(links, str):
        links = (links,)

    phantom_names = {
        'DashBoard1', 'DashBoard2', 'Dashboard1', 'Dashboard2',
    }
    dropped = 0
    for link in links:
        name = str(link)
        base = os.path.basename(name)
        if name not in phantom_names and base not in phantom_names:
            continue
        try:
            workbook.BreakLink(Name=name, Type=1)
            dropped += 1
            print(f'  잔존 외부링크 제거: {name}')
        except Exception as e:
            print(f'  외부링크 제거 실패({name}): {e}')
    return dropped


def _com_finalize_dashboard_links(workbook) -> int:
    """복원 직후: [n]/경로 외부참조 제거 + Parameter↔DashBoard2 링크 보장 + 가짜 링크 삭제."""
    n_loc = _com_localize_external_sheet_refs(workbook)
    n_param = _com_ensure_parameter_dashboard_links(workbook)
    n_drop = _com_drop_phantom_dashboard_links(workbook)
    if n_loc:
        print(f'  DashBoard 외부참조 로컬화: 셀 {n_loc}건')
    if n_param:
        print(f'  Parameter↔DashBoard2 내부참조 보정: {n_param}건')
    return n_loc + n_param + n_drop


def purge_xlsx_drawing_parts(xlsx_path: str) -> None:
    """
    openpyxl 저장본 xlsx 안에서 깨진 도면/차트 패키지 조각을 제거한다.

    DashBoard 시트를 지워도 xl/drawings·xl/charts·worksheet의 <drawing> 태그/
    .rels 링크가 남아 있으면 Excel Workbooks.Open이
    'Open 메서드에 오류가 있습니다'(-2146827284)로 실패한다.
    """
    import zipfile
    from io import BytesIO

    xlsx_path = os.path.abspath(xlsx_path)
    buf = BytesIO()

    def _clean_content_types(data: bytes) -> bytes:
        text = data.decode('utf-8', errors='ignore')
        text = re.sub(
            r'<Override[^>]+(drawings/|charts/|chartsheets/)[^/]*/>',
            '',
            text,
        )
        return text.encode('utf-8')

    def _clean_rels(data: bytes) -> bytes:
        text = data.decode('utf-8', errors='ignore')
        text = re.sub(
            r'<Relationship[^>]+Type="[^"]*(drawing|chart|chartsheet)[^"]*"[^/]*/>',
            '',
            text,
            flags=re.IGNORECASE,
        )
        return text.encode('utf-8')

    def _clean_sheet_xml(data: bytes) -> bytes:
        text = data.decode('utf-8', errors='ignore')
        text = re.sub(r'<drawing[^>]*?/>', '', text)
        text = re.sub(r'<drawing[^>]*>.*?</drawing>', '', text, flags=re.DOTALL)
        text = re.sub(r'<legacyDrawing[^>]*?/>', '', text)
        text = re.sub(r'<legacyDrawingHF[^>]*?/>', '', text)
        return text.encode('utf-8')

    skip_prefixes = (
        'xl/drawings/',
        'xl/charts/',
        'xl/chartsheets/',
    )
    with zipfile.ZipFile(xlsx_path, 'r') as zin, \
            zipfile.ZipFile(buf, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            name = info.filename.replace('\\', '/')
            if any(name.startswith(p) for p in skip_prefixes):
                continue
            data = zin.read(info.filename)
            base = name.rsplit('/', 1)[-1]
            if name == '[Content_Types].xml':
                data = _clean_content_types(data)
            elif name.endswith('.rels'):
                data = _clean_rels(data)
            elif name.startswith('xl/worksheets/') and base.endswith('.xml'):
                data = _clean_sheet_xml(data)
            # ZipInfo 재사용 시 암호/플래그 이슈 피하려면 새 이름으로 기록
            zout.writestr(name, data)

    with open(xlsx_path, 'wb') as f:
        f.write(buf.getvalue())


def _excel_open_workbook(excel_app, path, read_only=False, allow_repair=True):
    """
    Workbooks.Open 안전 호출 (위치 인자만).
    allow_repair=True 이면 CorruptLoad(복구 모드)까지 시도.
    """
    import shutil
    import tempfile

    path = os.path.abspath(path)

    def _open(p, corrupt_load=0):
        # Open(Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
        #      IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
        #      Converter, AddToMru, Local, CorruptLoad)
        if corrupt_load:
            return excel_app.Workbooks.Open(
                p, 0, bool(read_only),
                None, None, None, True,
                None, None, None, None, None, None, False,
                int(corrupt_load),
            )
        return excel_app.Workbooks.Open(p, 0, bool(read_only))

    errors = []
    for corrupt_load in (0, 1) if allow_repair else (0,):
        try:
            return _open(path, corrupt_load)
        except Exception as e:
            tag = '직접 열기' if corrupt_load == 0 else 'CorruptLoad 열기'
            errors.append(f'{tag}: {e}')

        safe_name = re.sub(r'[^\w.\-]+', '_', os.path.basename(path))
        tmp = os.path.join(tempfile.gettempdir(), f'cashflow_open_{os.getpid()}_{safe_name}')
        try:
            shutil.copy2(path, tmp)
            return _open(tmp, corrupt_load)
        except Exception as e:
            tag = 'temp 복사 열기' if corrupt_load == 0 else 'temp+CorruptLoad 열기'
            errors.append(f'{tag}: {e}')

    raise RuntimeError(
        'Excel Workbooks.Open 실패\n'
        f'  경로: {path}\n  ' + '\n  '.join(errors)
    )


def _com_dump_openpyxl_sheet(py_ws, com_ws, chunk_rows=250):
    """openpyxl 시트 값(+병합)을 COM 시트로 고속 배열 기록."""
    max_r = py_ws.max_row or 0
    max_c = py_ws.max_column or 0
    if max_r < 1 or max_c < 1:
        return
    for r0 in range(1, max_r + 1, chunk_rows):
        r1 = min(r0 + chunk_rows - 1, max_r)
        block = []
        for r in range(r0, r1 + 1):
            block.append(tuple(py_ws.cell(r, c).value for c in range(1, max_c + 1)))
        com_ws.Range(com_ws.Cells(r0, 1), com_ws.Cells(r1, max_c)).Value = tuple(block)
    for merged in py_ws.merged_cells.ranges:
        try:
            com_ws.Range(str(merged)).Merge()
        except Exception:
            pass


def _com_attach_dashboard_sheets(excel_app, mas_wb, out_wb, sheet_names):
    """취합용 mas_wb의 DashBoard 시트를 out_wb 맨 뒤에 붙인다."""
    restored = []
    mas_names = {
        mas_wb.Worksheets(i).Name for i in range(1, mas_wb.Worksheets.Count + 1)
    }
    for sheet_name in sheet_names:
        if sheet_name not in mas_names:
            print(f'  [{sheet_name}] 취합용에 시트가 없어 복원 스킵.')
            continue
        try:
            # 이미 있으면 제거 후 교체
            for i in range(out_wb.Worksheets.Count, 0, -1):
                if out_wb.Worksheets(i).Name == sheet_name:
                    target = out_wb.Worksheets(i)
                    # Excel은 통합문서에 '표시(Visible)' 상태 시트가 0개가 되는 걸 허용하지 않는다.
                    # (예: 나머지 시트가 전부 숨김이고 이 시트만 표시 상태인 경우 Delete()가 -2146827284로 실패)
                    # 삭제 대상이 마지막 표시 시트라면, 다른 시트를 하나 임시로 표시시켜 제약을 피한다.
                    visible_count = sum(
                        1 for j in range(1, out_wb.Worksheets.Count + 1)
                        if out_wb.Worksheets(j).Visible == -1  # xlSheetVisible
                    )
                    if visible_count <= 1 and target.Visible == -1:
                        for j in range(1, out_wb.Worksheets.Count + 1):
                            other = out_wb.Worksheets(j)
                            if other.Name != sheet_name:
                                other.Visible = -1
                                print(
                                    f'  [{sheet_name}] 삭제 전 임시로 [{other.Name}] 시트를 표시 처리 '
                                    f'(표시 시트 0개 방지)'
                                )
                                break
                    target.Delete()
                    break
            src = mas_wb.Worksheets(sheet_name)
            # Excel은 '숨김(Hidden/VeryHidden)' 시트에 대해 .Copy()를 호출하면
            # "Worksheet 클래스 중 Copy 메서드에 오류가 있습니다"로 실패한다.
            # 취합용(mas_wb)은 어차피 저장하지 않으므로, 복사 직전에만 임시로
            # 표시 상태로 바꿨다가 끝나면 원래 상태로 되돌린다.
            orig_visible = src.Visible
            if orig_visible != -1:  # -1 == xlSheetVisible
                src.Visible = -1
            try:
                # '개수+1 = 새 시트' 같은 인덱스 가정은 위치가 어긋나면 잘못된 인덱스
                # 접근(Invalid index) 오류로 이어질 수 있어, 복사 전/후 시트 '이름 집합'을
                # 비교해 새로 추가된 시트를 정확히 찾는다.
                before_names = {
                    out_wb.Worksheets(j).Name for j in range(1, out_wb.Worksheets.Count + 1)
                }
                # 1차 시도: Copy(Before, After) 위치 인자로 out_wb에 직접 복사.
                # (Copy(After=...) 같은 '이름있는 인자'는 win32com 동적 디스패치에서
                #  무시되고 조용히 새 통합문서가 생겨버리는 경우가 있어 위치 인자를 쓴다.)
                anchor = out_wb.Worksheets(out_wb.Worksheets.Count)
                src.Copy(None, anchor)
                after_names = {
                    out_wb.Worksheets(j).Name for j in range(1, out_wb.Worksheets.Count + 1)
                }
                added_names = after_names - before_names

                if len(added_names) != 1:
                    # 2차 시도(폴백): 인자 없이 복사 → 임시 통합문서 생성.
                    # 이 임시 통합문서의 시트를 Move가 아니라 다시 Copy로 out_wb에 옮긴다.
                    # Move는 원본(임시 통합문서)에 시트가 0개 남아 오류가 나지만,
                    # Copy는 원본을 비우지 않으므로 안전하다.
                    print(f'  [{sheet_name}] 1차 복사 방식 실패 → 대안(임시 통합문서 경유) 시도')
                    src.Copy()
                    temp_wb = excel_app.ActiveWorkbook
                    try:
                        temp_ws = temp_wb.Worksheets(1)
                        temp_ws.Copy(None, out_wb.Worksheets(out_wb.Worksheets.Count))
                    finally:
                        temp_wb.Close(SaveChanges=False)
                    after_names = {
                        out_wb.Worksheets(j).Name for j in range(1, out_wb.Worksheets.Count + 1)
                    }
                    added_names = after_names - before_names
                if len(added_names) != 1:
                    raise RuntimeError(
                        f'복사 후 새로 추가된 시트를 특정할 수 없습니다 '
                        f'(복사 전 {len(before_names)}개 → 복사 후 {len(after_names)}개, '
                        f'추가분: {sorted(added_names)})'
                    )
                new_ws = out_wb.Worksheets(next(iter(added_names)))
            finally:
                if orig_visible != -1:
                    src.Visible = orig_visible
            if new_ws.Name != sheet_name:
                new_ws.Name = sheet_name

            ws = new_ws
            n_shapes = int(ws.Shapes.Count)
            n_charts = int(ws.ChartObjects().Count)
            chart_shapes = 0
            for ci in range(1, n_charts + 1):
                try:
                    chart_shapes += int(ws.ChartObjects(ci).Chart.Shapes.Count)
                except Exception:
                    pass
            restored.append(sheet_name)
            print(
                f'  [{sheet_name}] 복원 완료 '
                f'(시트도형 {n_shapes}, 차트 {n_charts}, 차트내도형 {chart_shapes})'
            )
        except Exception as e:
            import traceback
            print(f'  [{sheet_name}] 대시보드 복원 실패: {e}')
            traceback.print_exc()
    return restored


def _rebuild_output_with_dashboards_via_com(excel_app, master_path, output_path, sheet_names):
    """
    Excel이 openpyxl 저장본을 Open하지 못할 때 사용하는 우회 경로.

    - 취합 데이터: openpyxl로 읽음 (Excel Open 불필요)
    - 새 Excel 통합문서에 값 배열로 기록
    - DashBoard: 취합용(master)에서 Copy/Move
    - 로컬 temp SaveAs 후 최종 경로로 복사
    """
    import shutil
    import tempfile
    from openpyxl import load_workbook

    print('  [우회] Excel이 최종파일을 열 수 없어 재조립합니다...')
    print('         (취합 데이터=openpyxl 읽기, DashBoard=취합용 COM 복사)')

    data_wb = load_workbook(output_path, data_only=False, keep_links=False)
    mas_wb = None
    out_wb = None
    xl_openxml = 51
    clean_path = os.path.join(
        tempfile.gettempdir(),
        f'cashflow_rebuild_{os.getpid()}.xlsx',
    )
    try:
        mas_wb = _excel_open_workbook(
            excel_app, master_path, read_only=True, allow_repair=False,
        )
        if sheet_names is None:
            sheet_names = [
                mas_wb.Worksheets(i).Name
                for i in range(1, mas_wb.Worksheets.Count + 1)
                if mas_wb.Worksheets(i).Name.startswith(DASHBOARD_SHEET_PREFIXES)
            ]

        excel_app.DisplayAlerts = False
        out_wb = excel_app.Workbooks.Add()
        while out_wb.Worksheets.Count > 1:
            out_wb.Worksheets(out_wb.Worksheets.Count).Delete()

        first = True
        for sheet_name in data_wb.sheetnames:
            print(f'  [우회] 데이터 시트 기록: {sheet_name}')
            if first:
                com_ws = out_wb.Worksheets(1)
                com_ws.Name = sheet_name[:31]
                first = False
            else:
                # Add(After=)는 DispatchEx에서 실패할 수 있어 Add 후 Move 사용
                last = out_wb.Worksheets(out_wb.Worksheets.Count)
                com_ws = out_wb.Worksheets.Add()
                com_ws.Move(After=last)
                com_ws.Name = sheet_name[:31]
            _com_dump_openpyxl_sheet(data_wb[sheet_name], com_ws)

        data_wb.close()
        data_wb = None

        restored = _com_attach_dashboard_sheets(excel_app, mas_wb, out_wb, sheet_names)
        _com_finalize_dashboard_links(out_wb)

        if os.path.isfile(clean_path):
            try:
                os.remove(clean_path)
            except OSError:
                pass
        out_wb.SaveAs(clean_path, FileFormat=xl_openxml)
        out_wb.Close(SaveChanges=False)
        out_wb = None
        mas_wb.Close(SaveChanges=False)
        mas_wb = None
        shutil.copy2(clean_path, output_path)
        return restored
    finally:
        if data_wb is not None:
            try:
                data_wb.close()
            except Exception:
                pass
        if out_wb is not None:
            try:
                out_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if mas_wb is not None:
            try:
                mas_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if os.path.isfile(clean_path):
            try:
                os.remove(clean_path)
            except OSError:
                pass


def restore_dashboard_sheets_to_output(master_path, output_path, sheet_names=None, excel_app=None):
    """
    취합용(master) DashBoard 시트를 최종 취합본(output)에 붙인다.

    1순위: Excel이 output을 Open할 수 있으면 그 책에 DashBoard만 붙임
    2순위: Open 실패 시 → openpyxl로 데이터 읽고 새 통합문서 재조립 + DashBoard 복사
           (취합용은 Open 되고, openpyxl 저장본만 Open 안 되는 경우 대응)
    """
    import shutil
    import tempfile

    master_path = os.path.abspath(master_path)
    output_path = os.path.abspath(output_path)
    if not os.path.isfile(master_path) or not os.path.isfile(output_path):
        print('  DashBoard 복원 스킵 — master/output 파일을 찾을 수 없습니다.')
        return []

    own_excel = excel_app is None
    if own_excel:
        excel_app = win32.DispatchEx('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        try:
            excel_app.AskToUpdateLinks = False
            excel_app.AlertBeforeOverwriting = False
        except Exception:
            pass

    restored = []
    out_wb = None
    mas_wb = None
    xl_openxml = 51
    clean_path = os.path.join(
        tempfile.gettempdir(),
        f'cashflow_dashboard_restore_{os.getpid()}.xlsx',
    )
    try:
        print(f'  취합용 Open: {master_path}')
        mas_wb = _excel_open_workbook(
            excel_app, master_path, read_only=True, allow_repair=False,
        )
        if sheet_names is None:
            sheet_names = [
                mas_wb.Worksheets(i).Name
                for i in range(1, mas_wb.Worksheets.Count + 1)
                if mas_wb.Worksheets(i).Name.startswith(DASHBOARD_SHEET_PREFIXES)
            ]

        # 1순위: 최종파일 Open 시도
        try:
            try:
                purge_xlsx_drawing_parts(output_path)
            except Exception as e:
                print(f'  최종파일 purge 경고: {e}')
            print(f'  최종파일 Open: {output_path}')
            out_wb = _excel_open_workbook(
                excel_app, output_path, read_only=False, allow_repair=True,
            )
        except Exception as open_err:
            print(f'  최종파일 Open 실패 → 우회 재조립 사용\n    ({open_err})')
            # master는 이미 열려 있으므로 닫고 우회 함수에서 다시 연다
            try:
                mas_wb.Close(SaveChanges=False)
            except Exception:
                pass
            mas_wb = None
            return _rebuild_output_with_dashboards_via_com(
                excel_app, master_path, output_path, sheet_names,
            )

        excel_app.DisplayAlerts = False
        restored = _com_attach_dashboard_sheets(excel_app, mas_wb, out_wb, sheet_names)
        # Copy가 만든 [1]취합... 외부참조 → 같은 파일 내부 시트로 되돌림
        _com_finalize_dashboard_links(out_wb)

        if os.path.isfile(clean_path):
            try:
                os.remove(clean_path)
            except OSError:
                pass
        out_wb.SaveAs(clean_path, FileFormat=xl_openxml)
        out_wb.Close(SaveChanges=False)
        out_wb = None
        mas_wb.Close(SaveChanges=False)
        mas_wb = None
        shutil.copy2(clean_path, output_path)
    finally:
        if out_wb is not None:
            try:
                out_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if mas_wb is not None:
            try:
                mas_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if os.path.isfile(clean_path):
            try:
                os.remove(clean_path)
            except OSError:
                pass
        if own_excel:
            excel_app.Quit()

    return restored


# ---------------------------------------------------------------------------
# CLI · 하위 호환
# ---------------------------------------------------------------------------


def save_text_boxes_csv(text_boxes, csv_path):
    fields = ['chart_title', 'chart_name', 'name', 'text']
    with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        writer.writeheader()
        for item in text_boxes:
            writer.writerow(item)


def extract_text_boxes(excel_path, sheet_name='DashBoard2', include_chart_shapes=True):
    return extract_dashboard_data(excel_path, sheet_name, include_chart_shapes=include_chart_shapes)['text_boxes']


def main(argv=None):
    parser = argparse.ArgumentParser(description='DashBoard 텍스트 상자 추출·취합')
    sub = parser.add_subparsers(dest='command')

    ext = sub.add_parser('extract', help='단일 파일 추출 CSV')
    ext.add_argument('excel_path')
    ext.add_argument('--sheet', default='DashBoard2')

    con = sub.add_parser('consolidate', help='취합용+부서 diff 취합본')
    con.add_argument('master_path')
    con.add_argument('--dept-dir', help='부서별파일 폴더')
    con.add_argument('--output-dir', help='출력 폴더 (기본: master와 동일)')
    con.add_argument('--sheet', action='append', help='시트명 (반복 가능)')

    args = parser.parse_args(argv)
    if args.command is None:
        parser.print_help()
        return

    if args.command == 'extract':
        df = extract_and_process(args.excel_path, args.sheet)
        out = os.path.splitext(args.excel_path)[0] + f'_{args.sheet}_processed.csv'
        df.to_csv(out, index=False, encoding='utf-8-sig')
        print(f'저장: {out} ({len(df)}건)')
        return

    dept_dir = args.dept_dir or os.path.join(os.path.dirname(args.master_path), '부서별파일')
    output_dir = args.output_dir or os.path.dirname(os.path.abspath(args.master_path))
    dept_paths = sorted(
        os.path.join(dept_dir, f)
        for f in os.listdir(dept_dir)
        if f.endswith('.xlsx') and not f.startswith('~$')
    ) if os.path.isdir(dept_dir) else []

    aggregate_dashboard_consolidation(
        args.master_path,
        dept_paths,
        output_dir,
        sheet_names=args.sheet or None,
    )


if __name__ == '__main__':
    main()


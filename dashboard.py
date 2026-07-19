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


def strip_dashboard_sheets(workbook) -> list[str]:
    """
    openpyxl 워크북에서 DashBoard 시트를 제거한다.

    openpyxl이 차트/도형이 있는 시트를 저장하면 xl/drawings 관계가 깨져
    Excel이 '내용을 복구하시겠습니까?'를 띄운다.
    저장 전에 대시보드 시트를 빼 두고, 저장 후 COM으로 취합용 원본을 붙인다.
    """
    removed = [
        name for name in list(workbook.sheetnames)
        if name.startswith(DASHBOARD_SHEET_PREFIXES)
    ]
    for name in removed:
        del workbook[name]
    return removed


def restore_dashboard_sheets_to_output(master_path, output_path, sheet_names=None, excel_app=None):
    """
    취합용(master) DashBoard 시트를 최종 취합본(output)에 붙인다.

    전제: output은 openpyxl 저장 전에 DashBoard가 제거된 '깨끗한' 파일이어야
    Excel 복구 프롬프트가 뜨지 않는다.

    절차:
      (1) output에 남아 있는 동명 DashBoard가 있으면 삭제
      (2) master 시트를 Copy() → Move로 output에 삽입
      (3) 로컬 temp에 SaveAs 후 최종 경로로 복사 (OneDrive 잠금/동기화 회피)

    Returns: 복원된 시트명 리스트.
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
            # 자동 복구 UI/알림 억제
            excel_app.AskToUpdateLinks = False
            excel_app.AlertBeforeOverwriting = False
        except Exception:
            pass

    restored = []
    out_wb = None
    mas_wb = None
    xl_openxml = 51  # xlOpenXMLWorkbook
    clean_path = os.path.join(
        tempfile.gettempdir(),
        f'cashflow_dashboard_restore_{os.getpid()}.xlsx',
    )
    try:
        # UpdateLinks=0: 링크 갱신 안 함 / ReadOnly=False
        out_wb = excel_app.Workbooks.Open(
            output_path, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True,
        )
        mas_wb = excel_app.Workbooks.Open(
            master_path, UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True,
        )

        out_names = {out_wb.Worksheets(i).Name for i in range(1, out_wb.Worksheets.Count + 1)}
        mas_names = {
            mas_wb.Worksheets(i).Name
            for i in range(1, mas_wb.Worksheets.Count + 1)
        }
        if sheet_names is None:
            # 취합용 시트 순서 유지
            sheet_names = [
                mas_wb.Worksheets(i).Name
                for i in range(1, mas_wb.Worksheets.Count + 1)
                if mas_wb.Worksheets(i).Name.startswith(DASHBOARD_SHEET_PREFIXES)
            ]

        excel_app.DisplayAlerts = False
        for sheet_name in sheet_names:
            if sheet_name not in mas_names:
                print(f'  [{sheet_name}] 취합용에 시트가 없어 복원 스킵.')
                continue
            try:
                insert_after = None
                if sheet_name in out_names:
                    old = out_wb.Worksheets(sheet_name)
                    idx = int(old.Index)
                    if idx > 1:
                        insert_after = out_wb.Worksheets(idx - 1)
                    old.Delete()
                    out_names.discard(sheet_name)

                if out_wb.Worksheets.Count < 1:
                    out_wb.Worksheets.Add()

                # master → 임시 통합문서(Copy 무인자) → output으로 Move
                src = mas_wb.Worksheets(sheet_name)
                src.Copy()
                temp_sheet = excel_app.ActiveWorkbook.Worksheets(1)
                if insert_after is not None:
                    temp_sheet.Move(After=insert_after)
                else:
                    # openpyxl 저장 전 제거된 경우 → 맨 뒤에 추가
                    temp_sheet.Move(After=out_wb.Worksheets(out_wb.Worksheets.Count))

                current_names = {
                    out_wb.Worksheets(i).Name
                    for i in range(1, out_wb.Worksheets.Count + 1)
                }
                if sheet_name not in current_names:
                    out_wb.ActiveSheet.Name = sheet_name

                # 도형 복원 여부 확인
                ws = out_wb.Worksheets(sheet_name)
                n_shapes = int(ws.Shapes.Count)
                n_charts = int(ws.ChartObjects().Count)
                chart_shapes = 0
                for ci in range(1, n_charts + 1):
                    try:
                        chart_shapes += int(ws.ChartObjects(ci).Chart.Shapes.Count)
                    except Exception:
                        pass

                out_names.add(sheet_name)
                restored.append(sheet_name)
                print(
                    f'  [{sheet_name}] 복원 완료 '
                    f'(시트도형 {n_shapes}, 차트 {n_charts}, 차트내도형 {chart_shapes})'
                )
            except Exception as e:
                print(f'  [{sheet_name}] 대시보드 복원 실패: {e}')

        if restored:
            if os.path.isfile(clean_path):
                try:
                    os.remove(clean_path)
                except OSError:
                    pass
            # OneDrive 경로에 직접 SaveAs 하면 동기화/잠금으로 깨질 수 있어
            # 로컬 temp에 저장한 뒤 최종 경로로 복사한다.
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

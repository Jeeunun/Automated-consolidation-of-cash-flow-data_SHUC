"""
DashBoard 시트 — 차트 위 텍스트 상자 추출 · 전처리 · 부서 diff · 취합본 저장.

- 추출: pywin32 (Shape/TextFrame) — openpyxl 불가
- 전처리: 줄바꿈 분리, 단위행 제거, 날짜 정렬
- 취합: 취합용 대비 부서 추가분 → CSV + 빨간 배경 XLSX

Windows + Excel + pywin32 + pandas + openpyxl 필요.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
from contextlib import contextmanager

import pandas as pd
import win32com.client as win32
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 설정
# ---------------------------------------------------------------------------

DASHBOARD_SHEET_PREFIXES = ('DashBoard', 'Dashboard')
UNIT_LABEL = '(단위 : 백만원)'
STATUS_EXISTING = '기존'
STATUS_DEPT_ADD = '부서추가'
FILL_DEPT_ADD = PatternFill(start_color='FF6666', end_color='FF6666', fill_type='solid')
HEADER_FONT = Font(bold=True)

OUTPUT_COLUMNS = [
    'status', 'source_file', 'chart_title', 'chart_name', 'text',
    'top', 'left', 'width', 'height', 'source', 'chart_object',
]

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


def process_textboxes_df(df: pd.DataFrame) -> pd.DataFrame:
    """줄바꿈 분리 · 단위행 제거 · chart+날짜 정렬."""
    if df.empty:
        return df.copy()

    out = df.copy()
    out['text'] = out['text'].astype(str).str.split('\n')
    out = out.explode('text').reset_index(drop=True)
    out['text'] = out['text'].str.strip()
    out = out[out['text'] != UNIT_LABEL]
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


def group_text_boxes_by_chart(text_boxes, charts):
    if not charts:
        for item in text_boxes:
            item['chart_title'] = ''
            item['chart_name'] = ''
        return text_boxes

    for item in text_boxes:
        cy, cx = _box_center(item)
        matched = [ch for ch in charts if _in_bounds(cy, cx, ch)]
        if len(matched) == 1:
            ch = matched[0]
        elif matched:
            ch = min(matched, key=lambda c: abs(cy - (c['top'] + c['height'] / 2)))
        else:
            ch = min(charts, key=lambda c: abs(cy - (c['top'] + c['height'] / 2)))
        item['chart_title'] = ch['title']
        item['chart_name'] = ch['name']
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
                'top': getattr(r, 'top', ''),
                'left': getattr(r, 'left', ''),
                'width': getattr(r, 'width', ''),
                'height': getattr(r, 'height', ''),
                'source': getattr(r, 'source', ''),
                'chart_object': getattr(r, 'chart_object', ''),
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
                'top': getattr(r, 'top', ''),
                'left': getattr(r, 'left', ''),
                'width': getattr(r, 'width', ''),
                'height': getattr(r, 'height', ''),
                'source': getattr(r, 'source', ''),
                'chart_object': getattr(r, 'chart_object', ''),
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


def save_consolidated_csv(df: pd.DataFrame, csv_path: str):
    export_cols = [c for c in OUTPUT_COLUMNS if c in df.columns]
    os.makedirs(os.path.dirname(os.path.abspath(csv_path)), exist_ok=True)
    df[export_cols].to_csv(csv_path, index=False, encoding='utf-8-sig')


def save_consolidated_xlsx(df: pd.DataFrame, xlsx_path: str):
    """CSV와 동일 데이터 + 부서추가 행 빨간 배경."""
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
            val = getattr(row, col_name, '')
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
    취합용 + 부서 DashBoard 텍스트 추출·diff·취합본 CSV/XLSX 저장.

    Returns:
        dict[sheet_name] -> {'csv': path, 'xlsx': path, 'add_count': int}
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
            csv_path = os.path.join(output_dir, f'dashboard_{safe_sheet}_취합본.csv')
            xlsx_path = os.path.join(output_dir, f'dashboard_{safe_sheet}_취합본.xlsx')

            save_consolidated_csv(consolidated, csv_path)
            save_consolidated_xlsx(consolidated, xlsx_path)

            print(f'    기존 {len(master_df)}건 / 부서추가 {add_count}건')
            print(f'    CSV  → {csv_path}')
            print(f'    XLSX → {xlsx_path} (부서추가=빨간 배경)')

            results[sheet_name] = {
                'csv': csv_path,
                'xlsx': xlsx_path,
                'add_count': add_count,
                'rows': len(consolidated),
            }

    return results


# ---------------------------------------------------------------------------
# CLI · 하위 호환
# ---------------------------------------------------------------------------


def save_text_boxes_csv(text_boxes, csv_path):
    fields = [
        'chart_title', 'chart_name', 'name', 'text',
        'top', 'left', 'width', 'height', 'source', 'chart_object',
    ]
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

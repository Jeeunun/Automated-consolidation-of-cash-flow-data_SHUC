import os
import openpyxl
import glob
import time
import re
from openpyxl.styles import PatternFill
import cProfile

# utils.py에서 함수와 변수들을 가져옵니다. (!! 반드시 utils.py, dashboard.py와 main.py는 같은 경로에 있어야 합니다 !!)
from utils import (
    has_any_border, copy_border_deep, is_blue_cell, row_has_blue, get_real_cell,
    load_workbooks, load_dept_workbooks, build_merge_lookup, build_dept_sheet_cache,
    shift_index_map_rows, shift_inserted_rows_set, shift_merged_cells_after_insert,
    build_index_store, index_store_as_map,
    map_execute_row_inserts, map_sync_prefix_alignment, map_restore_after_structural,
    dept_apply_blue_cell_updates,     dept_apply_remap_and_new_rows, dept_sync_all_index_rows,
    dept_mark_deleted_indexes, excel_layout_restore_after_inserts, fix_formulas_after_inserts,
    repair_master_external_link_refs,
    build_prefix_first_rows, copy_cell_style_safe,
    is_layout_col, should_capture_layout_cell,
    capture_cell_style, apply_cell_style, safe_unmerge_cells, get_valid_prefixes,
    extract_prefix, get_prefix_i_label, is_i_label_value, apply_i_section_label,
    is_allowed_index_prefix,
    get_section_header_rows, find_section_content_end_row,
    get_section_end_row, get_sheet_index_map,
    find_duplicate_index_rows, validate_new_index,
    build_dept_index_fixup, fix_plan_virtual_indexes,
    detect_renumber_cascade_candidates, should_skip_dept_new_index_add,
    get_suggested_new_index,
    extract_section_layout_templates,
    apply_row_style_from_neighbor,
    TARGET_BLUE_CODES, COLOR_ADD, COLOR_DELETE, INDEX_COL, NAME_COL, 
    STYLE_SNAPSHOT_MAX_COL, numeric_sheets, text_sheets,
    master_file, target_folder, base_path, prev_year_short, prev_month, dept_file_matches_prev_month
)

def _load_master_values_workbook():
    '''L열~ 수식의 계산값 스냅샷용 (data_only).'''
    if os.path.exists(master_file):
        return openpyxl.load_workbook(
            master_file, data_only=True, keep_links=False,
        )
    return None

# ------------------------------------
# [숫자 시트 취합]
def aggregate_numeric_sheets(master_wb, dept_wbs, numeric_sheets):
    """
    [함수 설명] 부서별 파일에서 숫자 시트들을 취합용 파일로 가져와 취합합니다.

    [3층 아키텍처]
      PLAN       — 부서 스캔: 신규/재번호/수정·삭제 후보 수집
      MAP        — IndexStore: insert·행번호 갱신·밀린 L열~ baseline 복원
      DEPT DIFF  — 파란셀·remap·신규행 데이터·layout 후 전체 인덱스 동기화
      EXCEL LAYOUT — I~L 병합/서식·J열 구분·연두색 H (insert 있을 때만)

    dept_wbs: {파일경로: workbook} — 시트마다 재로드하지 않고 재사용.
    """
    master_wb_values = _load_master_values_workbook()
    # 추가행과 삭제행의 구분을 위한 색상
    fill_add = PatternFill(start_color=COLOR_ADD, end_color=COLOR_ADD, fill_type='solid')
    fill_del = PatternFill(start_color=COLOR_DELETE, end_color=COLOR_DELETE, fill_type='solid')

    # 인덱스 구조를 판별하기 위한 글자 패턴 (예: 영어/숫자 조합 뒤에 하이픈과 숫자 -> UP-1,,,)
    prefix_pattern = re.compile(r"^([A-Za-z0-9]+)-(\d+)$")
    valid_prefixes = get_valid_prefixes()

    # 취합해야 하는 세 가지 숫자 시트(원안, 수정1, 수정2)를 하나씩 돌며 작업합니다.
    for sheet_name in numeric_sheets:
        if sheet_name not in master_wb.sheetnames:
            print(f'취합용 파일에 {sheet_name}이 없습니다.')
            continue

        master_ws = master_wb[sheet_name]
        max_col = master_ws.max_column
        print(f'[{sheet_name}] 인덱스 매칭 및 취합 작업 시작...')

        master_merge_lookup = build_merge_lookup(master_ws)
        ws_values = (
            master_wb_values[sheet_name]
            if master_wb_values and sheet_name in master_wb_values.sheetnames
            else None
        )
        index_store, section_baseline = build_index_store(
            master_ws, max_col, master_merge_lookup, valid_prefixes,
            ws_values=ws_values,
        )
        master_map = index_store_as_map(index_store)
        original_master_layout = extract_section_layout_templates(
            master_ws, valid_prefixes, col_end=max_col
        )
        prefix_first_rows = build_prefix_first_rows(master_map, prefix_pattern)
        unknown_prefix_warnings = set()
        index_validation_warnings = []

        all_dept_indexes = set()
        updated_count = 0
        added_count = 0
        deleted_count = 0 # 삭제 카운트 초기화 선언 보완
        new_rows_buffer = []
        pending_add_indexes = set()
        fix_remap_queue = []
        fix_virtual_indexes = set()
        
        # 이번 시트 처리 중 master_ws에 새로 삽입된 절대 행 번호 모음.
        inserted_rows = set()
        dept_dup_blocked = {}
        all_cascade_rows = set()

        dept_cache = build_dept_sheet_cache(dept_wbs, sheet_name)

        for df, ctx in dept_cache.items():
            dept_ws = ctx['ws']
            dept_merge_lookup = ctx['merge_lookup']
            dept_map = ctx['index_map']
            dept_rows_sorted = ctx['rows_sorted']

            dup_blocked = set()
            for idx, r1, r2 in find_duplicate_index_rows(dept_ws):
                dup_blocked.add(idx)
                msg = f'[인덱스 중복] {idx} — {r1}행·{r2}행에 동일 입력 (부서 파일 수정 필요)'
                index_validation_warnings.append(msg)
                print(f'  -> {msg}')
            dept_dup_blocked[df] = dup_blocked

            def find_index_above(dept_row_num, _sorted=dept_rows_sorted):
                '''부서 시트에서 dept_row_num 바로 위에 있는 H열 인덱스(예: U-4)를 반환.'''
                above_idx = None
                for r_num, idx in _sorted:
                    if r_num < dept_row_num:
                        above_idx = idx
                    else:
                        break
                return above_idx

            fix_plans = build_dept_index_fixup(
                dept_map, master_map, prefix_pattern, valid_prefixes,
                dept_ws=dept_ws, master_ws=master_ws, max_col=max_col,
                dept_merge_lookup=dept_merge_lookup,
                master_merge_lookup=master_merge_lookup,
            )
            cascade_rows = set()
            for plan in fix_plans.values():
                cascade_rows.update(plan['cascade_rows'])
            all_cascade_rows.update(cascade_rows)

            for prefix, plan in fix_plans.items():
                remap_parts = []
                for dept_idx, dept_row, master_idx in plan['remaps']:
                    if master_idx not in master_map:
                        continue
                    fix_remap_queue.append({
                        'master_idx': master_idx,
                        'dept_row': dept_row,
                        'dept_ws': dept_ws,
                        'dept_merge_lookup': dept_merge_lookup,
                    })
                    remap_parts.append(f'{dept_idx}→{master_idx}')
                if remap_parts:
                    print(f'  -> [인덱스 자동보정] {", ".join(remap_parts)}')
                    all_dept_indexes.update(
                        master_idx for master_idx, _, _ in plan['remaps']
                        if master_idx in master_map
                    )

                anchor_idx_str = plan['anchor_idx'] or find_index_above(plan['insert_row'])
                if plan['new_index'] not in pending_add_indexes and plan['new_index'] not in master_map:
                    new_rows_buffer.append({
                        'idx_str': plan['new_index'],
                        'dept_ws': dept_ws,
                        'dept_row': plan['insert_row'],
                        'prefix': prefix,
                        'anchor_idx_str': anchor_idx_str,
                        'dept_merge_lookup': dept_merge_lookup,
                        'auto_fixup': True,
                    })
                    pending_add_indexes.add(plan['new_index'])
                all_dept_indexes.update(fix_plan_virtual_indexes(plan))
                fix_virtual_indexes.update(fix_plan_virtual_indexes(plan))
                print(
                    f'  -> [인덱스 자동보정] 삽입행 {plan["insert_idx"]} '
                    f'(부서 {plan["insert_row"]}행) → {plan["new_index"]}'
                )

            for idx_str, info in dept_map.items():
                if idx_str in dup_blocked:
                    continue
                dept_row = info['row_num']
                if dept_row in cascade_rows:
                    continue

                all_dept_indexes.add(idx_str)

                if idx_str in master_map:
                    # 취합용 파일에 이미 존재하는 인덱스들의 일반 수정셀 취합은 
                    # 신규 행 추가가 완료되어 행 번호가 최종 확정된 이후(하단 루프)에 처리합니다.
                    pass
                else:
                    # [행 추가] 부서 H열에 직접 입력한 신규 인덱스만 추가
                    prefix = extract_prefix(idx_str, prefix_pattern)
                    if prefix is not None and not is_allowed_index_prefix(
                        prefix, master_map, valid_prefixes
                    ):
                        unknown_prefix_warnings.add(idx_str)
                        continue

                    ok, val_msg, suggested = validate_new_index(
                        idx_str, prefix, master_map, pending_add_indexes, prefix_pattern
                    )
                    if not ok:
                        index_validation_warnings.append(val_msg)
                        print(f'  -> [인덱스 오류] {val_msg}')
                        continue

                    if idx_str in pending_add_indexes or idx_str in master_map:
                        continue

                    if should_skip_dept_new_index_add(
                        idx_str, prefix, dept_map, master_map, prefix_pattern,
                        fix_plans, dept_ws=dept_ws, master_ws=master_ws,
                        max_col=max_col, dept_merge_lookup=dept_merge_lookup,
                        master_merge_lookup=master_merge_lookup,
                    ):
                        if prefix not in fix_plans:
                            print(
                                f'  -> [인덱스] {idx_str} 하단 추가 생략 '
                                f'(재번호 자동보정 대상 — 삽입행→{get_suggested_new_index(master_map, prefix)})'
                            )
                        continue

                    anchor_idx_str = find_index_above(dept_row)
                    new_rows_buffer.append({
                        'idx_str': idx_str, 'dept_ws': dept_ws, 'dept_row': dept_row,
                        'prefix': prefix, 'anchor_idx_str': anchor_idx_str,
                        'dept_merge_lookup': dept_merge_lookup,
                    })
                    pending_add_indexes.add(idx_str)
                    print(f'  -> [신규 인덱스] {idx_str} (부서 {dept_row}행)')

        # ---- LAYER 1: MAP — insert·행번호·baseline L열~ ----
        inserted_rows, added_index_values, master_map, master_merge_lookup = (
            map_execute_row_inserts(
                master_ws, index_store, new_rows_buffer, master_map,
                prefix_first_rows, max_col, inserted_rows,
            )
        )
        added_count += len(added_index_values)
        all_dept_indexes.update(added_index_values)

        skip_align = set(added_index_values) | fix_virtual_indexes

        inserted_rows, master_merge_lookup, _align_applied = map_sync_prefix_alignment(
            master_ws, index_store, dept_cache, prefix_pattern,
            skip_align, max_col, master_merge_lookup, inserted_rows,
        )
        master_map = index_store_as_map(index_store)

        map_restore_after_structural(
            master_ws, index_store, master_merge_lookup, skip_align,
        )

        # ---- LAYER 2: DEPT DIFF — 파란셀·remap·신규행 데이터 ----
        updated_count += dept_apply_blue_cell_updates(
            master_ws, dept_cache, master_map, max_col,
            master_merge_lookup, all_cascade_rows, dept_dup_blocked,
        )

        if new_rows_buffer or fix_remap_queue:
            updated_count += dept_apply_remap_and_new_rows(
                master_ws, master_map, fix_remap_queue, new_rows_buffer,
                max_col, master_merge_lookup, fill_add=fill_add,
            )

        # ---- LAYER 3: EXCEL LAYOUT — I~L 병합·J열 구분·연두색 H ----
        if new_rows_buffer or inserted_rows:
            master_merge_lookup = excel_layout_restore_after_inserts(
                master_ws, original_master_layout, valid_prefixes,
                prefix_pattern, index_store, section_baseline,
                master_map, fix_remap_queue, new_rows_buffer,
                skip_align, inserted_rows, added_index_values,
                max_col, fill_add,
            )
            updated_count += fix_formulas_after_inserts(
                master_ws, inserted_rows, max_col, master_merge_lookup,
            )
            updated_count += dept_sync_all_index_rows(
                master_ws, dept_cache, master_map, max_col,
                master_merge_lookup, skip_align, all_cascade_rows, dept_dup_blocked,
            )
            updated_count += dept_apply_blue_cell_updates(
                master_ws, dept_cache, master_map, max_col,
                master_merge_lookup, all_cascade_rows, dept_dup_blocked,
            )

        # ---- DEPT DIFF: prefix 담당 부서에 없는 인덱스 → H열 빨간색 ----
        deleted_count += dept_mark_deleted_indexes(
            master_ws, index_store, dept_cache, master_map,
            prefix_pattern, valid_prefixes, fill_del, skip_indexes=skip_align,
        )

        # 부서별 코드에 없는 부서는 처리되지 않습니다.
        if unknown_prefix_warnings:
            print(
                f'정의되지 않은 부서 코드: {len(unknown_prefix_warnings)}건 건너뜀 '
                f'{sorted(unknown_prefix_warnings)}'
            )
        if index_validation_warnings:
            print(f' 인덱스 검증 경고 {len(index_validation_warnings)}건 — 위 안내대로 부서 파일 수정 후 재취합하세요.')

        print(f'>> {sheet_name} 처리 완료 / 수정: {updated_count}건, 추가: {added_count}건, 삭제: {deleted_count}건\n')
# ------------------------------------
# 텍스트 시트 단순 행 대체
def aggregate_text_sheets(master_wb, dept_wbooks, text_sheets):
    """
    [함수 설명] 전월대비 비교나 지출계획 등 텍스트 중심 시트 내용을 취합합니다.
    [원리] 인덱스 번호 없이, 행 단위로 부서 파일의 파란색 수정 마크가 감지되면 해당 취합용 행 전체를 덮어씌웁니다.
    """
    for sheet_name in text_sheets:
        if sheet_name not in master_wb.sheetnames:
            print(f'{sheet_name} 시트가 취합용 파일에 없습니다.')
            continue

        master_ws = master_wb[sheet_name]
        print(f'[{sheet_name}] 텍스트 중심 서술형 시트 취합 시작.')

        max_row = master_ws.max_row
        max_col = master_ws.max_column
        replace_row_count = 0
        master_merge_lookup = build_merge_lookup(master_ws)
        dept_merge_lookups = {}
        for dept_wb in dept_wbooks:
            if sheet_name in dept_wb.sheetnames:
                dept_merge_lookups[id(dept_wb)] = build_merge_lookup(dept_wb[sheet_name])

        for r in range(1, max_row + 1):
            row_replaced = False

            for dept_wb in dept_wbooks:
                if row_replaced: break # 해당 행이 이미 교체 완료되었다면 다음 파일 검사 패스

                if sheet_name in dept_wb.sheetnames:
                    dept_ws = dept_wb[sheet_name]
                    dept_ml = dept_merge_lookups.get(id(dept_wb))

                    for c in range(1, max_col + 1):
                        dept_cell_raw = dept_ws.cell(row=r, column=c)
                        d_eff = get_real_cell(dept_ws, dept_cell_raw, dept_ml)
                        if not is_blue_cell(dept_cell_raw) and not is_blue_cell(d_eff):
                            continue

                        print(f'  -> [파란색 셀 포착] {r}행 {c}열 / 해당 {r}행 전체를 이식합니다.')

                        for copy_c in range(1, max_col + 1):
                            m_cell_raw = master_ws.cell(row=r, column=copy_c)
                            d_cell_raw = dept_ws.cell(row=r, column=copy_c)
                            m_cell = get_real_cell(master_ws, m_cell_raw, master_merge_lookup)
                            d_cell = get_real_cell(dept_ws, d_cell_raw, dept_ml)
                            is_blue = is_blue_cell(d_cell_raw) or is_blue_cell(d_cell)

                            if d_cell.value is not None or is_blue:
                                m_cell.value = d_cell.value

                            if is_blue:
                                d_color = str(d_cell.fill.start_color.rgb).upper()
                                m_cell.fill = PatternFill(
                                    start_color=d_color, end_color=d_color, fill_type='solid'
                                )

                        replace_row_count += 1
                        row_replaced = True
                        break

        print(f'>> {sheet_name} 텍스트 취합 완료  / 총 {replace_row_count}개 행 교체 완료.\n')

# ------------------------------------------
# 실행
if __name__ == "__main__":
    import sys
    PROFILE = os.environ.get('CF_PROFILE', '').strip() in ('1', 'true', 'yes')

    print("====================================================")
    print(f" [{prev_year_short}년 {prev_month:02d}월 기준] 자동 취합")
    print("====================================================")
    start_time = time.time()

    if not os.path.exists(master_file):
        print(f' 취합용 파일이 경로에 없습니다: {master_file}')
        sys.exit(1)

    # 취합용 원본 Parameter 등: [1]DashBoard2! → DashBoard2! (COM 저장, 도면 보존)
    try:
        master_fixed = repair_master_external_link_refs(master_file)
        if master_fixed:
            print(f' 취합용 원본 외부 참조 수정·저장: {len(master_fixed)}건')
            for _name, coord, old, new in master_fixed[:8]:
                print(f'   - {_name}!{coord}: {old} → {new}')
            if len(master_fixed) > 8:
                print(f'   ... 외 {len(master_fixed) - 8}건')
        else:
            print(' 취합용 원본: 외부 참조([n]Sheet!) 수정 대상 없음')
    except Exception as e:
        print(f' 취합용 원본 외부 참조 수정 경고(취합은 계속): {e}')

    master_wb = openpyxl.load_workbook(
        master_file, data_only=False, keep_links=False, keep_vba=False
    )
    print(' 취합용 파일 로드')

    all_dept_xlsx = [
        f for f in glob.glob(os.path.join(target_folder, '*.xlsx'))
        if os.path.basename(f) != os.path.basename(master_file)
        and not os.path.basename(f).startswith('~$')
    ]
    dept_files = [
        f for f in all_dept_xlsx
        if dept_file_matches_prev_month(f, prev_year_short, prev_month, numeric_sheets)
    ]
    skipped_dept_files = [f for f in all_dept_xlsx if f not in dept_files]

    print(f' 전월({prev_year_short}.{prev_month:02d}) 취합 대상 부서 파일 총 {len(dept_files)}개 발견.')
    if skipped_dept_files:
        print(f'      (기준 외 파일 {len(skipped_dept_files)}개 제외)')

    if not dept_files:
        print(" 취합 요건을 충족하는 부서별 파일이 없습니다.")
        sys.exit(1)

    dept_wbs = load_dept_workbooks(dept_files)
    if not dept_wbs:
        print(" 부서 파일을 열 수 없습니다.")
        sys.exit(1)

    print("\n 숫자 시트 취합 시작.")
    if PROFILE:
        with cProfile.Profile() as pr:
            aggregate_numeric_sheets(master_wb, dept_wbs, numeric_sheets)
        pr.dump_stats('aggregate_numeric_sheets.prof')
        print(' (프로파일 저장: aggregate_numeric_sheets.prof)')
    else:
        aggregate_numeric_sheets(master_wb, dept_wbs, numeric_sheets)

    print(" 텍스트 시트 취합 시작.")
    dept_wbook_list = list(dept_wbs.values())
    if PROFILE:
        with cProfile.Profile() as pr2:
            aggregate_text_sheets(master_wb, dept_wbook_list, text_sheets)
        pr2.dump_stats('aggregate_text_sheets.prof')
        print(' (프로파일 저장: aggregate_text_sheets.prof)')
    else:
        aggregate_text_sheets(master_wb, dept_wbook_list, text_sheets)

    print(' DashBoard 텍스트 취합 시작. (DashBoard1만)')
    from dashboard import aggregate_dashboard_consolidation
    aggregate_dashboard_consolidation(
        master_path=master_file,
        dept_paths=dept_files,
        output_dir=base_path,
        sheet_names=['DashBoard1'],
    )

    for wb in dept_wbs.values():
        wb.close()
    
    # 5단계: 최종 취합본 파일 저장
    end_time = time.time()
    elapsed = end_time - start_time
    print("====================================================")
    print(f" 소요 시간: {int(elapsed // 60)}분 {int(elapsed % 60)}초")

    folder_file = os.path.join(base_path, '최종취합본')
    os.makedirs(folder_file, exist_ok=True)
    output_file = os.path.join(folder_file, f"도시개발본부_현금흐름_최종취합_완성본({prev_year_short}.{prev_month:02d}).xlsx")

    # openpyxl이 DashBoard(차트/도형)를 저장하면 xlsx 도면 XML이 깨져
    # Excel Open이 실패함 → 저장 전 시트 제거 + 저장 후 도면 잔여물 purge
    from dashboard import (
        strip_dashboard_sheets,
        purge_xlsx_drawing_parts,
        restore_dashboard_sheets_to_output,
    )

    # DashBoard1/2는 openpyxl 저장 시 도면이 깨지므로 제외 → 이후 COM 복원
    # (복원 Copy가 [1]Sheet! 외부참조를 만들면 restore 단계에서 내부참조로 되돌림)
    removed_dash = strip_dashboard_sheets(master_wb)
    if removed_dash:
        print(f" openpyxl 안전 저장을 위해 DashBoard 임시 제외: {', '.join(removed_dash)}")
    else:
        print(" DashBoard 제외 대상 없음")

    master_wb.save(output_file)
    master_wb.close()
    try:
        purge_xlsx_drawing_parts(output_file)
        print(" openpyxl 저장본 도면/차트 잔여 패키지 정리 완료")
    except Exception as e:
        print(f" 도면 잔여 패키지 정리 경고(복원에서 재시도): {e}")
    print(f" 최종 취합본 파일 저장완료 \n -> 경로: {output_file}")

    # 취합용 원본에서 DashBoard1/2 복원 + [n]외부참조 로컬화 + Parameter 링크 보정
    if removed_dash:
        print(" DashBoard 시트 도형·텍스트 복원 중...")
        try:
            restored = restore_dashboard_sheets_to_output(
                master_file, output_file, sheet_names=removed_dash,
            )
            if restored:
                print(f" DashBoard 복원 완료: {', '.join(restored)}")
            else:
                print(" DashBoard 복원 실패: 시트를 붙이지 못했습니다.")
        except Exception as e:
            print(f" DashBoard 복원 실패(숫자/텍스트 취합은 저장됨): {e}")
    else:
        print(" DashBoard 복원 생략 (제외된 시트 없음)")
    print("====================================================")

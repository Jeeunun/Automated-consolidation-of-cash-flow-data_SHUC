"""dashboard_processing.py — dashboard.py 로 통합됨. 하위 호환용."""
from dashboard import (
    extract_sort_key,
    normalize_text,
    process_textboxes_df,
    extract_and_process,
    build_consolidated,
    aggregate_dashboard_consolidation,
)

__all__ = [
    'extract_sort_key',
    'normalize_text',
    'process_textboxes_df',
    'extract_and_process',
    'build_consolidated',
    'aggregate_dashboard_consolidation',
]

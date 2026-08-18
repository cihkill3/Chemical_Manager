"""Workbook/session and layout helpers used by the synchronization engine."""

import os
import time


def is_file_locked(file_path):
    if not file_path or not os.path.exists(file_path):
        return False
    try:
        with open(file_path, "a+b"):
            pass
        return False
    except (IOError, OSError):
        return True


def check_and_wait_lock(file_path, log_fn=None, max_retries=None, retry_delay=60, is_auto_sync=False, check_stop_fn=None):
    if not file_path or not os.path.exists(file_path):
        return True
    attempts = 0
    while is_file_locked(file_path):
        if check_stop_fn and check_stop_fn():
            if log_fn:
                log_fn("사용자 요청으로 파일 대기를 즉시 중단했습니다.")
            return False
        attempts += 1
        if max_retries is not None and attempts > max_retries:
            if log_fn:
                log_fn(f"파일 사용 중 대기 시간 초과 ({max_retries}회 지남).")
            return False
        if log_fn:
            log_fn(f"{os.path.basename(file_path)} 파일이 사용/편집 중입니다. 1분 후 다시 확인합니다... ({attempts}회차)")
        for _ in range(int(retry_delay * 10)):
            if check_stop_fn and check_stop_fn():
                if log_fn:
                    log_fn("사용자 요청으로 파일 대기를 즉시 중단했습니다.")
                return False
            time.sleep(0.1)
    return True


def build_header_map(worksheet, header_row=1, scan_limit=30):
    last_column = worksheet.Cells(header_row, worksheet.Columns.Count).End(-4159).Column
    headers = {}
    for column in range(1, max(last_column + 1, scan_limit)):
        value = worksheet.Cells(header_row, column).Value
        if value:
            headers[str(value).strip()] = column
    return headers, max(headers.values(), default=0)


def reserve_missing_headers(header_map, requested_headers, last_column=0):
    """Reserve after real headers; this performs no Excel writes."""
    result = dict(header_map)
    next_column = max(last_column, max(result.values(), default=0))
    for header in requested_headers:
        if header and header not in result:
            next_column += 1
            result[header] = next_column
    return result


def write_reserved_headers(worksheet, header_map, header_row=1):
    for header, column in header_map.items():
        if not worksheet.Cells(header_row, column).Value:
            worksheet.Cells(header_row, column).Value = header


def meaningful_data_last_row(worksheet, header_map, key_kind="order"):
    """Return the last real data row without trusting a sparsely populated key column."""
    if key_kind == "db":
        candidates = ("Key", "Manufacturer", "Catalog No.", "Product Name")
    else:
        candidates = (
            "Order No.", "Product Name", "Original Product Name",
            "Manufacturer", "Catalog No.",
        )
    rows = [
        worksheet.Cells(worksheet.Rows.Count, header_map[header]).End(-4162).Row
        for header in candidates
        if header_map.get(header)
    ]
    return max([1, *rows])


def should_manage_product_name(order_number, current_formula):
    """Preserve a literal aliquot/manual name when the row has no Order No."""
    has_order = bool(str(order_number or "").strip())
    is_program_formula = str(current_formula or "").strip().startswith("=")
    return has_order or is_program_formula

import datetime

from core.sync_engine import SyncEngine
from core.excel_manager import should_manage_product_name


def test_manual_aliquot_name_without_order_number_is_preserved():
    assert not should_manage_product_name("", "Aliquot of reagent A")
    assert should_manage_product_name("", "=IFERROR(INDEX(...),Z2)")
    assert should_manage_product_name("26-0500", "Ordered reagent")


def test_excel_numeric_identifiers_become_plain_text():
    assert SyncEngine.as_excel_text(1.0) == "1"
    assert SyncEngine.as_excel_text(123.0) == "123"
    assert SyncEngine.as_excel_text("001") == "001"
    assert SyncEngine.as_excel_text("ABC.0") == "ABC.0"


def test_received_accepts_markers_and_supported_dates():
    accepted = [
        "O", "o", "ㅇ", "2026-08-14", "26-08-14", "260814",
        260814.0, datetime.date(2026, 8, 14), datetime.datetime(2026, 8, 14, 12, 0),
    ]
    assert all(SyncEngine.is_received(value) for value in accepted)


def test_received_rejects_blanks_invalid_dates_and_other_values():
    rejected = [None, "", "X", "2026/08/14", "2026-02-30", "26-13-01", "260230"]
    assert not any(SyncEngine.is_received(value) for value in rejected)

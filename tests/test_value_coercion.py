"""Unit tests for JSON-safe value coercion in the Graph Excel write paths.

The Graph workbook endpoints take values as a JSON body, so datetime/date
(not JSON-serialisable) and Decimal (rejected by requests' encoder) must be
converted before sending.  Regression guard for the
'Object of type datetime is not JSON serializable' failure.
"""
import datetime
from decimal import Decimal

from tropi_excel.excelfile import _to_excel_value, _to_excel_grid


def test_date_becomes_excel_serial():
    # 2020-01-01 is Excel serial 43831 (1899-12-30 epoch).
    assert _to_excel_value(datetime.date(2020, 1, 1)) == 43831


def test_datetime_midnight_matches_date_serial():
    assert _to_excel_value(datetime.datetime(2020, 1, 1, 0, 0, 0)) == 43831.0


def test_datetime_carries_fractional_day():
    # Noon = half a day past the date serial.
    assert _to_excel_value(datetime.datetime(2020, 1, 1, 12, 0, 0)) == 43831.5


def test_decimal_becomes_float():
    out = _to_excel_value(Decimal("12.50"))
    assert out == 12.5 and isinstance(out, float)


def test_passthrough_types_unchanged():
    assert _to_excel_value("text") == "text"
    assert _to_excel_value(7) == 7
    assert _to_excel_value(None) is None


def test_grid_coerces_every_cell_and_handles_none():
    grid = [[datetime.date(2020, 1, 1), "x"], [Decimal("1.5"), None]]
    assert _to_excel_grid(grid) == [[43831, "x"], [1.5, None]]
    assert _to_excel_grid(None) is None

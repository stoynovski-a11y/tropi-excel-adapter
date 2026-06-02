"""Unit tests for set_font_color's input guard.

The actual PATCH to range/format/font needs a live Graph session (covered by
the live test); here we only assert the cheap, pure validation guard so a bad
color string fails fast before any network call.
"""
import pytest

from tropi_excel.excelfile import ExcelFileClient


def _wb():
    # Build a client object without opening a session — we only exercise the
    # input guard, which runs before any session use.
    return ExcelFileClient.__new__(ExcelFileClient)


@pytest.mark.parametrize("bad", ["red", "FF0000", "#FFF", "#GG0000", "#FF00000", "", None, 0xFF0000])
def test_rejects_bad_color(bad):
    with pytest.raises(ValueError):
        ExcelFileClient.set_font_color(_wb(), "Sheet", "A1", bad)

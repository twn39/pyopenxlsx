"""Tests for DefinedNames façade and hyperlink helpers."""

from pyopenxlsx import Workbook, link_cell
from pyopenxlsx.hyperlink import is_external_url


def test_is_external_url():
    assert is_external_url("https://example.com")
    assert is_external_url("mailto:a@b.com")
    assert not is_external_url("Sheet2!A1")
    assert not is_external_url("A1")


def test_defined_names_define_and_contains():
    wb = Workbook()
    ws = wb.active
    dns = wb.defined_names
    dn = dns.define("Sales", "Sheet1!$A$1:$A$10")
    assert dn.name() == "Sales"
    assert "Sales" in dns
    assert dns["Sales"].refers_to() == "Sheet1!$A$1:$A$10"
    # redefine is idempotent
    dns.define("Sales", "Sheet1!$B$1:$B$5")
    assert dns.get("Sales").refers_to() == "Sheet1!$B$1:$B$5"
    # sheet-local by worksheet object
    dns.define("Local", "Sheet1!$C$1", sheet=ws)
    assert dns.exists("Local", local_sheet_id=0)
    wb.close()


def test_defined_names_legacy_append_still_works():
    wb = Workbook()
    dns = wb.defined_names
    dns.append("GlobalName", "Sheet1!$A$1")
    assert dns.count() == 1
    assert dns.get("GlobalName").name() == "GlobalName"
    names = [dn.name() for dn in dns]
    assert "GlobalName" in names
    wb.close()


def test_worksheet_link_external_and_internal(tmp_path):
    path = tmp_path / "links_facade.xlsx"
    with Workbook() as wb:
        ws = wb.active
        ws.link("A1", "https://www.example.com", text="Example", tooltip="go")
        assert ws["A1"].value == "Example"
        assert ws.has_hyperlink("A1")
        wb.create_sheet("Data")
        ws.link("A2", "Data!A1", text="Jump", internal=True)
        assert ws.has_hyperlink("A2")
        link_cell(ws, "A3", "https://x.ai", text="xAI")
        assert ws["A3"].value == "xAI"
        wb.save(path)
    assert path.exists()

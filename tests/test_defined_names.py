from pyopenxlsx import Workbook

def test_defined_names_basic():
    wb = Workbook()
    dns = wb.defined_names
    
    dns.append("GlobalName", "Sheet1!$A$1")
    assert dns.count() == 1
    assert dns.exists("GlobalName")
    
    dn = dns.get("GlobalName")
    assert dn.name() == "GlobalName"
    assert dn.refers_to() == "Sheet1!$A$1"
    assert dn.local_sheet_id() is None
    
    dns.remove("GlobalName")
    assert dns.count() == 0
    assert not dns.exists("GlobalName")

def test_defined_names_local():
    wb = Workbook()
    wb.create_sheet("Sheet2")
    
    dns = wb.defined_names
    # Local names use 0-based index for sheet
    dns.append("LocalName", "Sheet1!$B$1", local_sheet_id=0)
    dns.append("LocalName", "Sheet2!$B$1", local_sheet_id=1)
    
    assert dns.count() == 2
    assert dns.exists("LocalName", local_sheet_id=0)
    assert dns.exists("LocalName", local_sheet_id=1)
    
    dn0 = dns.get("LocalName", local_sheet_id=0)
    assert dn0.refers_to() == "Sheet1!$B$1"
    
    dn1 = dns.get("LocalName", local_sheet_id=1)
    assert dn1.refers_to() == "Sheet2!$B$1"

def test_defined_names_iteration():
    wb = Workbook()
    dns = wb.defined_names
    dns.append("N1", "Sheet1!$A$1")
    dns.append("N2", "Sheet1!$A$2")
    
    names = [dn.name() for dn in dns]
    assert "N1" in names
    assert "N2" in names
    assert len(names) == 2

def test_defined_names_save_load(tmp_path):
    filepath = tmp_path / "names.xlsx"
    wb = Workbook()
    dns = wb.defined_names
    dns.append("SavedName", "Sheet1!$C$1")
    wb.save(filepath)
    
    wb2 = Workbook(filepath)
    assert wb2.defined_names.exists("SavedName")
    assert wb2.defined_names.get("SavedName").refers_to() == "Sheet1!$C$1"
    wb2.close()

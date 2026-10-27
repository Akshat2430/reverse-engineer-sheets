from sheetre import refs as R


def test_parse_ref_forms():
    r = R.parse_ref("'My Sheet'!$B$2:C10", "X")
    assert (r.kind, r.sheet, r.c1, r.r1, r.c2, r.r2) == ("range", "My Sheet", 2, 2, 3, 10)
    assert R.parse_ref("A1", "X").kind == "cell"
    w = R.parse_ref("Data!A:B", "X")
    assert w.whole_col and w.c2 == 2
    assert R.parse_ref("[1]Budget!C5", "X").kind == "external"
    assert R.parse_ref("Sheet1!#REF!", "X").kind == "error"
    assert R.parse_ref("TaxRate", "X").kind == "name"


def test_refs_with_function_context():
    f = "=VLOOKUP(A2,Lookup!$A$2:$B$9,2,TRUE)+OFFSET(Data!B2,0,1)"
    rs = R.refs(f, "Model")
    funcs = {(r.text, r.func, r.arg) for r in rs}
    assert ("Lookup!$A$2:$B$9", "VLOOKUP", 1) in funcs
    assert ("Data!B2", "OFFSET", 0) in funcs


def test_r1c1_detects_same_and_different_patterns():
    a = R.r1c1("=B2*C2*1.5", 2, 4)
    b = R.r1c1("=B3*C3*1.5", 3, 4)
    c = R.r1c1("=B4*C4", 4, 4)
    assert a == b and a != c
    assert R.r1c1("=$A$1+B2", 2, 3) == R.r1c1("=$A$1+B5", 5, 3)


def test_literals_skip_structural_numbers():
    lits = R.literals("=VLOOKUP(A2,T!A:C,3,FALSE)*1.18+ROUND(B2,2)")
    assert [l.value for l in lits] == [1.18]
    assert lits[0].context == "multiply"


def test_comparisons_find_thresholds():
    pairs = R.comparisons("=IF(A2>=5000,1,IF(A2>100,2,3))")
    assert (">=", 5000.0) in pairs and (">", 100.0) in pairs


def test_cell_ranges_expand_within_bounds():
    r = R.parse_ref("S!A:A", "S")
    assert len(r.cells(max_row=5, max_col=3)) == 5

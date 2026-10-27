"""Direct reads of the .xlsx/.xlsm zip for things openpyxl hides or drops.

- sheet state (visible / hidden / veryHidden) and the sheet -> part mapping
- calcPr (iterative calculation settings)
- defined names exactly as stored
- external link parts: target file, sheet names, cached cell values
- x14:dataValidation in extLst (openpyxl drops these with a warning)
- cached error values, array formulas, threaded comments
"""
import posixpath
import re
import zipfile
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
}
R_ID = "{%s}id" % NS["r"]


class Package:
    def __init__(self, path):
        self.path = str(path)
        self.z = zipfile.ZipFile(self.path)
        self.names = set(self.z.namelist())

    def xml(self, name):
        return ET.fromstring(self.z.read(name))

    def rels(self, part):
        d, f = posixpath.split(part)
        rp = posixpath.join(d, "_rels", f + ".rels")
        if rp not in self.names:
            return {}
        out = {}
        for r in self.xml(rp).findall("rel:Relationship", NS):
            target = r.get("Target")
            if r.get("TargetMode") != "External":
                target = posixpath.normpath(posixpath.join(d, target)) if not target.startswith("/") else target[1:]
            out[r.get("Id")] = dict(type=r.get("Type").rsplit("/", 1)[-1], target=target,
                                    external=r.get("TargetMode") == "External")
        return out

    # ------------------------------------------------------------ workbook level
    def workbook(self):
        wb = self.xml("xl/workbook.xml")
        rels = self.rels("xl/workbook.xml")
        sheets = []
        for s in wb.findall("m:sheets/m:sheet", NS):
            sheets.append(dict(name=s.get("name"), state=s.get("state", "visible"),
                               part=rels.get(s.get(R_ID), {}).get("target")))
        names = []
        for dn in wb.findall("m:definedNames/m:definedName", NS):
            lsid = dn.get("localSheetId")
            names.append(dict(name=dn.get("name"), refers_to=dn.text or "",
                              scope=sheets[int(lsid)]["name"] if lsid is not None else None,
                              hidden=dn.get("hidden") == "1"))
        calc = wb.find("m:calcPr", NS)
        calc = dict(calc.attrib) if calc is not None else {}
        ext_ids = [e.get(R_ID) for e in wb.findall("m:externalReferences/m:externalReference", NS)]
        externals = []
        for i, rid in enumerate(ext_ids, start=1):
            part = rels.get(rid, {}).get("target")
            externals.append(self.external_link(i, part))
        return dict(sheets=sheets, defined_names=names, calc=calc, external_links=externals,
                    has_vba="xl/vbaProject.bin" in self.names)

    def external_link(self, index, part):
        info = dict(index=index, part=part, target=None, sheets=[], cached={})
        if not part or part not in self.names:
            return info
        rels = self.rels(part)
        info["target"] = next((r["target"] for r in rels.values() if r["external"]), None)
        root = self.xml(part)
        book = root.find("m:externalBook", NS)
        if book is None:
            return info
        sheet_names = [s.get("val") for s in book.findall("m:sheetNames/m:sheetName", NS)]
        info["sheets"] = sheet_names
        for sd in book.findall("m:sheetDataSet/m:sheetData", NS):
            sname = sheet_names[int(sd.get("sheetId", 0))] if sheet_names else "?"
            for cell in sd.iter("{%s}cell" % NS["m"]):
                v = cell.find("m:v", NS)
                info["cached"][f"{sname}!{cell.get('r')}"] = v.text if v is not None else None
        return info

    # ------------------------------------------------------------ sheet level
    def sheet_extras(self, part):
        """x14 validations, cached errors, array formulas for one worksheet part."""
        out = dict(x14_validations=[], cached_errors=[], array_formulas=[], x14_conditional=0)
        if not part or part not in self.names:
            return out
        raw = self.z.read(part).decode("utf-8", "replace")
        root = ET.fromstring(raw)
        for dv in root.iter("{%s}dataValidation" % NS["x14"]):
            f1 = dv.find("x14:formula1/xm:f", NS)
            f2 = dv.find("x14:formula2/xm:f", NS)
            sq = dv.find("xm:sqref", NS)
            out["x14_validations"].append(dict(type=dv.get("type"), operator=dv.get("operator"),
                                               formula1=f1.text if f1 is not None else None,
                                               formula2=f2.text if f2 is not None else None,
                                               sqref=sq.text if sq is not None else None))
        out["x14_conditional"] = len(list(root.iter("{%s}conditionalFormatting" % NS["x14"])))
        for c in root.iter("{%s}c" % NS["m"]):
            if c.get("t") == "e":
                v = c.find("m:v", NS)
                f = c.find("m:f", NS)
                out["cached_errors"].append(dict(cell=c.get("r"), error=v.text if v is not None else None,
                                                 has_formula=f is not None))
            f = c.find("m:f", NS)
            if f is not None and f.get("t") == "array":
                out["array_formulas"].append(dict(cell=c.get("r"), ref=f.get("ref"), formula=f.text))
        return out

    def threaded_comment_parts(self):
        return sorted(n for n in self.names if re.match(r"xl/threadedComments/.*\.xml$", n))

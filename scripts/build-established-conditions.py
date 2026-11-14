"""Build established-conditions.json from the Established Condition Database spreadsheet.

Usage (from the repo root):
    python scripts/build-established-conditions.py "C:\\path\\to\\Established Condition Database.xlsx"

Reads the "New List" sheet (the list published on the EI Colorado website) and
writes established-conditions.json next to index.html. Uses only the Python
standard library, so nothing needs to be installed.
"""
import datetime
import json
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

SHEET_NAME = "New List"
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def column_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def read_sheet_rows(path, sheet_name):
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall(NS + "si"):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {r.get("Id"): r.get("Target") for r in rels.findall(PKG_REL_NS + "Relationship")}
        sheet_path = None
        for sheet in workbook.find(NS + "sheets"):
            if sheet.get("name") == sheet_name:
                target = targets[sheet.get(REL_NS + "id")]
                sheet_path = "xl/" + target.lstrip("/").replace("xl/", "", 1)
        if not sheet_path:
            raise SystemExit(f'The workbook has no sheet named "{sheet_name}".')

        rows = []
        for row in ET.fromstring(z.read(sheet_path)).iter(NS + "row"):
            cells = {}
            for c in row.findall(NS + "c"):
                v = c.find(NS + "v")
                kind = c.get("t")
                if kind == "s" and v is not None:
                    value = shared[int(v.text)]
                elif kind == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(NS + "t"))
                else:
                    value = v.text if v is not None else ""
                if value != "":
                    cells[column_index(c.get("r"))] = value
            if cells:
                rows.append([cells.get(i, "") for i in range(4)])
        return rows


def excel_date(value):
    """Excel stores dates as day counts. Returns YYYY-MM-DD, or '' if not a date."""
    try:
        serial = int(float(value))
    except (TypeError, ValueError):
        return ""
    return (datetime.date(1899, 12, 30) + datetime.timedelta(days=serial)).isoformat()


def main():
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    rows = read_sheet_rows(sys.argv[1], SHEET_NAME)

    # Find the header row (the first row whose first cell is "Medical Diagnosis").
    start = next(i for i, r in enumerate(rows) if clean(r[0]).lower() == "medical diagnosis") + 1

    conditions = {}
    for name, details, answer, reviewed in rows[start:]:
        name, details = clean(name), clean(details)
        answer = clean(answer).lower()
        if not name:
            continue
        if answer not in ("yes", "no"):
            raise SystemExit(f'"{name}" has an unexpected Established Condition value: "{answer}".')
        entry = {"name": name, "established": answer == "yes"}
        if details and details.lower() != name.lower():
            entry["details"] = details
        date = excel_date(reviewed)
        if date:
            entry["reviewed"] = date
        key = name.lower()
        if key in conditions and conditions[key]["established"] != entry["established"]:
            raise SystemExit(f'"{name}" appears twice with different answers. Fix the spreadsheet first.')
        conditions[key] = entry

    ordered = sorted(conditions.values(), key=lambda e: e["name"].lower())
    out = Path(__file__).resolve().parent.parent / "established-conditions.json"
    out.write_text(
        json.dumps({"source": "EI Colorado Established Conditions list", "conditions": ordered}, ensure_ascii=False, indent=0),
        encoding="utf-8",
    )
    yes = sum(1 for e in ordered if e["established"])
    print(f"Wrote {len(ordered)} conditions ({yes} established, {len(ordered) - yes} not) to {out}")


if __name__ == "__main__":
    main()

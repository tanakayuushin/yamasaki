#!/usr/bin/env python3
"""
.xlsx に、文字だけの簡単なシートを末尾に追加する（メール文面などの設定用）。

  python3 add_text_sheet.py 入力.xlsx 出力.xlsx シート名 セル=文字 [セル=文字 ...]

  文字の中の \\n は改行になる。例:
    python3 add_text_sheet.py in.xlsx out.xlsx 文面 "A1=件名" "B1=お見積りのお願い" "A3=本文" "B3=【会社名】\\n【担当者名】 様\\n..."
"""
import re
import sys
import zipfile

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def add_text_sheet(src, out, sheet_name, cells, col_widths=None):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    patched = {}
    wb = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ct = zin.read("[Content_Types].xml").decode("utf-8")
    styles = zin.read("xl/styles.xml").decode("utf-8")

    # 折り返し＋上寄せの書式を styles.xml に 1 つ追加
    m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', styles, re.S)
    xf_count = int(m.group(1))
    wrap_xf = xf_count
    new_xfs = m.group(2) + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
    styles = styles.replace(m.group(0), '<cellXfs count="%d">%s</cellXfs>' % (xf_count + 1, new_xfs))
    patched["xl/styles.xml"] = styles

    n = 1
    while "xl/worksheets/sheet%d.xml" % n in names:
        n += 1
    path = "xl/worksheets/sheet%d.xml" % n
    sheet_id = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
    rid = "rId%d" % (max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1)

    rows = {}
    for ref, text in cells:
        mm = re.match(r"([A-Z]+)(\d+)$", ref)
        r = int(mm.group(2))
        rows.setdefault(r, []).append((col_to_num(mm.group(1)), ref, text))
    row_xml = ""
    for r in sorted(rows):
        cs = ""
        for _, ref, text in sorted(rows[r]):
            style = ' s="%d"' % wrap_xf if "\n" in text else ""
            cs += '<c r="%s" t="inlineStr"%s><is><t xml:space="preserve">%s</t></is></c>' % (ref, style, esc(text))
        row_xml += '<row r="%d">%s</row>' % (r, cs)
    cols_xml = ""
    if col_widths:
        cols_xml = "<cols>" + "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (col_to_num(c), col_to_num(c), w) for c, w in col_widths.items()) + "</cols>"
    patched[path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                     '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="%s">'
                     '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>%s'
                     '<sheetData>%s</sheetData></worksheet>') % (NS_R, cols_xml, row_xml)

    wb = wb.replace("</sheets>", '<sheet xmlns:r="%s" name="%s" sheetId="%d" r:id="%s"/></sheets>' % (NS_R, esc(sheet_name), sheet_id, rid))
    patched["xl/workbook.xml"] = wb
    patched["xl/_rels/workbook.xml.rels"] = rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (rid, n))
    patched["[Content_Types].xml"] = ct.replace("</Types>",
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % path)
    if "docProps/app.xml" in names:
        app = zin.read("docProps/app.xml").decode("utf-8")
        app = re.sub(r'<vt:i4>(\d+)</vt:i4>', lambda m: "<vt:i4>%d</vt:i4>" % (int(m.group(1)) + 1), app, count=1)
        app = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)(" baseType="lpstr">)(.*?)(</vt:vector>)',
                     lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3) + m.group(4) + "<vt:lpstr>%s</vt:lpstr>" % esc(sheet_name) + m.group(5), app, count=1, flags=re.S)
        patched["docProps/app.xml"] = app

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = patched.get(item.filename)
            zout.writestr(item, zin.read(item.filename) if data is None else data.encode("utf-8"))
            written.add(item.filename)
        for name, data in patched.items():
            if name not in written:
                zout.writestr(name, data.encode("utf-8"))


if __name__ == "__main__":
    if len(sys.argv) < 5:
        print(__doc__)
        sys.exit(1)
    cells = []
    for a in sys.argv[4:]:
        ref, text = a.split("=", 1)
        cells.append((ref, text.replace("\\n", "\n")))
    add_text_sheet(sys.argv[1], sys.argv[2], sys.argv[3], cells, {"A": 12, "B": 70})
    print("作成しました:", sys.argv[2])

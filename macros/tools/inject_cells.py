#!/usr/bin/env python3
"""
.xlsx/.xlsm のセルに値を書き込む（書式はそのまま。XML を直接書き換えるので図形やマクロも保持。無い行・セルは作る）。

  python3 inject_cells.py 入力 出力 シート名 値.json

  値.json は {"B3": "土木", "M3": 1000000, ...} の形。数値は数値セル、それ以外は文字列セルになる。
"""
import json
import re
import sys
import zipfile


def esc(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def inject(src, out, sheet_name, values):
    zin = zipfile.ZipFile(src)
    wb = zin.read("xl/workbook.xml").decode("utf-8")
    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb).group(1)
    target = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
             re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    path = "xl/" + target.group(1).lstrip("/").replace("xl/", "", 1)
    xml = zin.read(path).decode("utf-8")

    def col_num(col):
        n = 0
        for ch in col:
            n = n * 26 + (ord(ch) - 64)
        return n

    for ref, val in values.items():
        mm = re.match(r"([A-Z]+)(\d+)$", ref)
        col, rownum = mm.group(1), int(mm.group(2))
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            new_inner = "<v>%s</v>" % val
            new_attr = ""
        else:
            new_inner = '<is><t xml:space="preserve">%s</t></is>' % esc(str(val))
            new_attr = ' t="inlineStr"'
        pat = re.compile(r'<c r="%s"((?: [a-z]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % ref, re.S)
        m = pat.search(xml)
        if m:
            attrs = re.sub(r' t="[^"]*"', "", m.group(1))
            xml = xml[:m.start()] + '<c r="%s"%s%s>%s</c>' % (ref, attrs, new_attr, new_inner) + xml[m.end():]
            continue
        new_cell = '<c r="%s"%s>%s</c>' % (ref, new_attr, new_inner)
        rm = re.search(r'<row r="%d"[^>]*?(/>|>)(.*?)(</row>)?' % rownum, xml, re.S)
        rm = re.search(r'<row r="%d"[^>]*>.*?</row>|<row r="%d"[^>]*/>' % (rownum, rownum), xml, re.S)
        if rm:
            row_xml = rm.group(0)
            if row_xml.endswith("/>"):
                row_xml = row_xml[:-2] + ">" + new_cell + "</row>"
            else:
                # 列順を保って挿入
                cells = list(re.finditer(r'<c r="([A-Z]+)\d+"', row_xml))
                pos = None
                for cm in cells:
                    if col_num(cm.group(1)) > col_num(col):
                        pos = cm.start()
                        break
                if pos is None:
                    row_xml = row_xml.replace("</row>", new_cell + "</row>")
                else:
                    row_xml = row_xml[:pos] + new_cell + row_xml[pos:]
            xml = xml[:rm.start()] + row_xml + xml[rm.end():]
        else:
            # 行ごと作る（行番号順の位置に挿入）
            new_row = '<row r="%d">%s</row>' % (rownum, new_cell)
            pos = None
            for rm2 in re.finditer(r'<row r="(\d+)"', xml):
                if int(rm2.group(1)) > rownum:
                    pos = rm2.start()
                    break
            if pos is None:
                xml = xml.replace("</sheetData>", new_row + "</sheetData>")
            else:
                xml = xml[:pos] + new_row + xml[pos:]

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == path:
                data = xml.encode("utf-8")
            zout.writestr(item, data)


if __name__ == "__main__":
    if len(sys.argv) != 5:
        print(__doc__)
        sys.exit(1)
    inject(sys.argv[1], sys.argv[2], sys.argv[3], json.load(open(sys.argv[4], encoding="utf-8")))
    print("作成しました:", sys.argv[2])

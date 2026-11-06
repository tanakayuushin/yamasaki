#!/usr/bin/env python3
"""
.xlsx/.xlsm の既存セルに値を書き込む（書式はそのまま。XML を直接書き換えるので図形やマクロも保持）。

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

    for ref, val in values.items():
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            new_inner = "<v>%s</v>" % val
            new_attr = ""
        else:
            new_inner = '<is><t xml:space="preserve">%s</t></is>' % esc(str(val))
            new_attr = ' t="inlineStr"'
        pat = re.compile(r'<c r="%s"((?: [a-z]+="[^"]*")*)\s*(?:/>|>.*?</c>)' % ref, re.S)
        m = pat.search(xml)
        if not m:
            raise SystemExit("セルが見つかりません（行や列が存在しないため書けません）: " + ref)
        attrs = re.sub(r' t="[^"]*"', "", m.group(1))
        xml = xml[:m.start()] + '<c r="%s"%s%s>%s</c>' % (ref, attrs, new_attr, new_inner) + xml[m.end():]

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

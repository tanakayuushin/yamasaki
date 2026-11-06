#!/usr/bin/env python3
"""
.xlsx の中のシートを XML レベルで複製して、別名のシートとして追加する。
（書式・列幅・オートフィルタ・印刷設定ごとそのまま複製される。openpyxl では落ちる情報も保持）

  python3 add_sheet_copy.py 入力.xlsx 出力.xlsx 複製元シート名 新しいシート名 [挿入位置(0始まり)] [--clear-data 見出し行]

  --clear-data N を付けると、新しいシート（複製先ではなく「複製元」のシート）の N 行目より下の値を消す。
  例）依頼リストに入力用、Sheet1 を結果用（空）にする:
      python3 add_sheet_copy.py book.xlsx out.xlsx Sheet1 依頼リスト 0 --clear-data 2
"""
import re
import sys
import zipfile

NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def xml_escape(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def add_sheet_copy(src_xlsx, out_xlsx, src_sheet, new_sheet, position=0, clear_data_header_row=None):
    zin = zipfile.ZipFile(src_xlsx)
    names = zin.namelist()
    patched = {}

    wb = zin.read("xl/workbook.xml").decode("utf-8")
    wb_rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ct = zin.read("[Content_Types].xml").decode("utf-8")

    sheets = re.findall(r'<sheet [^>]*/>', wb)
    src_tag = None
    for t in sheets:
        if re.search(r'name="%s"' % re.escape(xml_escape(src_sheet)), t):
            src_tag = t
    if not src_tag:
        raise SystemExit("シートが見つかりません: " + src_sheet)
    src_rid = re.search(r'r:id="([^"]+)"', src_tag).group(1)
    src_part = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % src_rid, wb_rels)
    if not src_part:
        src_part = re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % src_rid, wb_rels)
    src_path = "xl/" + src_part.group(1).lstrip("/").replace("xl/", "", 1) if not src_part.group(1).startswith("/") else src_part.group(1).lstrip("/")
    if not src_path.startswith("xl/"):
        src_path = "xl/" + src_path

    # 新しいシートのパーツ名
    n = 1
    while "xl/worksheets/sheet%d.xml" % n in names:
        n += 1
    new_path = "xl/worksheets/sheet%d.xml" % n
    sheet_ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', wb)]
    new_sheet_id = max(sheet_ids) + 1
    rids = [int(x) for x in re.findall(r'Id="rId(\d+)"', wb_rels)]
    new_rid = "rId%d" % (max(rids) + 1)

    # シート本体をコピー（codeName は外して、後で付け直せるようにする）
    sheet_xml = zin.read(src_path).decode("utf-8")
    sheet_xml = re.sub(r'\s+codeName="[^"]*"', "", sheet_xml)
    sheet_xml = re.sub(r'<tabSelected="1"', "", sheet_xml)
    sheet_xml = sheet_xml.replace(' tabSelected="1"', "")
    patched[new_path] = sheet_xml

    # シートの関連（印刷設定など）は複製して別パーツにする
    src_rels_path = src_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if src_rels_path in names:
        rels = zin.read(src_rels_path).decode("utf-8")
        for m in re.finditer(r'<Relationship [^>]*/>', rels):
            tag = m.group(0)
            target = re.search(r'Target="([^"]+)"', tag).group(1)
            if "printerSettings" in target:
                old_bin = "xl/" + target.replace("../", "")
                k = 1
                while re.sub(r"\d+\.bin$", "%d.bin" % k, old_bin) in names or re.sub(r"\d+\.bin$", "%d.bin" % k, old_bin) in patched:
                    k += 1
                new_bin = re.sub(r"\d+\.bin$", "%d.bin" % k, old_bin)
                patched[new_bin] = zin.read(old_bin)
                rels = rels.replace(target, "../" + new_bin[3:])
            elif "drawing" in target:
                rels = rels.replace(tag, "")  # 図形は複製しない（ボタンは後で builder が付ける）
                sheet_xml = re.sub(r'<drawing [^>]*/>', "", sheet_xml)
                patched[new_path] = sheet_xml
        patched[new_path.replace("worksheets/", "worksheets/_rels/") + ".rels"] = rels

    # workbook.xml：シート一覧に挿入、definedName の localSheetId を調整
    new_tag = '<sheet xmlns:r="%s" name="%s" sheetId="%d" r:id="%s"/>' % (NS_R, xml_escape(new_sheet), new_sheet_id, new_rid)
    sheet_list = list(sheets)
    src_index = sheet_list.index(src_tag)
    sheet_list.insert(position, new_tag)
    wb = wb.replace("".join(sheets), "".join(sheet_list))

    def shift_local(m):
        i = int(m.group(1))
        return 'localSheetId="%d"' % (i + 1 if i >= position else i)
    defined = re.findall(r'<definedName [^>]*>.*?</definedName>', wb, re.S)
    new_defined = []
    for d in defined:
        d2 = re.sub(r'localSheetId="(\d+)"', shift_local, d)
        new_defined.append(d2)
        lm = re.search(r'localSheetId="(\d+)"', d)
        if lm and int(lm.group(1)) == src_index:
            # 複製元シート固有の名前（オートフィルタ範囲など）を新シート用にも作る
            d3 = re.sub(r'localSheetId="\d+"', 'localSheetId="%d"' % position, d)
            d3 = d3.replace(xml_escape(src_sheet) + "!", "'" + xml_escape(new_sheet) + "'!")
            new_defined.append(d3)
    if defined:
        wb = wb.replace("".join(defined), "".join(new_defined))
    # 先頭シートを表示状態に
    wb = re.sub(r'(<workbookView [^>]*?)\s*activeTab="\d+"', r"\1", wb)
    patched["xl/workbook.xml"] = wb

    wb_rels = wb_rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/></Relationships>' % (new_rid, n))
    patched["xl/_rels/workbook.xml.rels"] = wb_rels

    ct = ct.replace("</Types>",
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>' % new_path)
    patched["[Content_Types].xml"] = ct

    # docProps/app.xml のシート名一覧（あれば）を更新
    if "docProps/app.xml" in names:
        app = zin.read("docProps/app.xml").decode("utf-8")
        app = re.sub(r'<vt:i4>(\d+)</vt:i4>', lambda m: "<vt:i4>%d</vt:i4>" % (int(m.group(1)) + 1), app, count=1)
        app = re.sub(r'(<TitlesOfParts><vt:vector size=")(\d+)(" baseType="lpstr">)',
                     lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3) + "<vt:lpstr>%s</vt:lpstr>" % xml_escape(new_sheet), app, count=1)
        patched["docProps/app.xml"] = app

    # 複製元シートのデータ行の値を消す（結果用の空シートにする）
    if clear_data_header_row is not None:
        src_xml = zin.read(src_path).decode("utf-8")
        def strip_row(m):
            row = m.group(0)
            r = int(re.search(r'<row r="(\d+)"', row).group(1))
            if r <= clear_data_header_row:
                return row
            row = re.sub(r'(<c [^>]*?)\s+t="[^"]*"([^>]*)>.*?</c>', r"\1\2/>", row, flags=re.S)
            row = re.sub(r'(<c [^>]*?)>.*?</c>', r"\1/>", row, flags=re.S)
            return row
        src_xml = re.sub(r'<row [^>]*>.*?</row>', strip_row, src_xml, flags=re.S)
        src_xml = re.sub(r'<row [^>]*/>', lambda m: m.group(0), src_xml)
        patched[src_path] = src_xml

    with zipfile.ZipFile(out_xlsx, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = patched.get(item.filename)
            if data is None:
                zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr(item, data if isinstance(data, bytes) else data.encode("utf-8"))
            written.add(item.filename)
        for name, data in patched.items():
            if name not in written:
                zout.writestr(name, data if isinstance(data, bytes) else data.encode("utf-8"))


if __name__ == "__main__":
    args = sys.argv[1:]
    clear = None
    if "--clear-data" in args:
        i = args.index("--clear-data")
        clear = int(args[i + 1])
        del args[i:i + 2]
    if len(args) not in (4, 5):
        print(__doc__)
        sys.exit(1)
    pos = int(args[4]) if len(args) == 5 else 0
    add_sheet_copy(args[0], args[1], args[2], args[3], pos, clear)
    print("作成しました:", args[1])

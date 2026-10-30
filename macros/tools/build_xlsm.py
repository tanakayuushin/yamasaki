#!/usr/bin/env python3
"""
元の .xlsx に VBA 標準モジュール（.bas）を組み込んで .xlsm を作るツール。

  python3 build_xlsm.py 元ブック.xlsx モジュール.bas 出力.xlsm [実行するSub名 ボタンの文字]

最後の 2 つを指定すると、先頭シートの表の右側（O2 付近）に
クリックでそのマクロを実行する図形ボタンを置きます。

Excel 本体を使わずに、MS-OVBA 形式の vbaProject.bin（CP932 / 日本語対応）を
生成し、元の xlsx のシート・書式はそのまま保ったまま追加します。
標準ライブラリのみで動作します。
"""
import re
import struct
import sys
import uuid
import zipfile

CODEPAGE = 932
ENC = "cp932"
PROJECT_NAME = "VBAProject"

# ---------------------------------------------------------------- MS-OVBA 圧縮

def _copy_token_params(pos_in_chunk):
    bit_count = 4
    while (1 << bit_count) < pos_in_chunk:
        bit_count += 1
    bit_count = max(bit_count, 4)
    length_mask = 0xFFFF >> bit_count
    max_len = length_mask + 3
    return bit_count, length_mask, max_len


def _compress_chunk(data):
    """1 チャンク（<=4096 バイト）を圧縮してトークン列を返す"""
    out = bytearray()
    pos = 0
    n = len(data)
    while pos < n:
        flag = 0
        tokens = bytearray()
        for bit in range(8):
            if pos >= n:
                break
            bit_count, length_mask, max_len = _copy_token_params(pos)
            best_len, best_off = 0, 0
            if pos > 0:
                max_off = min(pos, 1 << bit_count)
                for off in range(1, max_off + 1):
                    start = pos - off
                    l = 0
                    while l < max_len and pos + l < n and data[start + l] == data[pos + l]:
                        l += 1
                    if l > best_len:
                        best_len, best_off = l, off
                        if l == max_len:
                            break
            if best_len >= 3:
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                tokens += struct.pack("<H", token)
                flag |= 1 << bit
                pos += best_len
            else:
                tokens.append(data[pos])
                pos += 1
        out.append(flag)
        out += tokens
    return bytes(out)


def compress(data: bytes) -> bytes:
    out = bytearray(b"\x01")
    for i in range(0, len(data), 4096):
        chunk = data[i:i + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) < 4096:
            header = ((len(comp) + 2 - 3) & 0x0FFF) | 0x3000 | 0x8000
            out += struct.pack("<H", header) + comp
        else:  # 圧縮しない方が小さい場合は生データ（4096 バイトに揃える）
            raw = chunk + b"\x00" * (4096 - len(chunk))
            header = ((4096 + 2 - 3) & 0x0FFF) | 0x3000
            out += struct.pack("<H", header) + raw
    return bytes(out)


def decompress(data: bytes) -> bytes:
    """検証用の展開（MS-OVBA 2.4.1.3.1）"""
    assert data[0] == 1
    out = bytearray()
    i = 1
    while i < len(data):
        header = struct.unpack_from("<H", data, i)[0]
        size = (header & 0x0FFF) + 3
        compressed = bool(header & 0x8000)
        chunk = data[i + 2:i + size]
        i += size
        start = len(out)
        if not compressed:
            out += chunk
            continue
        j = 0
        while j < len(chunk):
            flag = chunk[j]
            j += 1
            for bit in range(8):
                if j >= len(chunk):
                    break
                if flag & (1 << bit):
                    token = struct.unpack_from("<H", chunk, j)[0]
                    j += 2
                    pos = len(out) - start
                    bit_count, length_mask, _ = _copy_token_params(pos)
                    length = (token & length_mask) + 3
                    offset = (token >> (16 - bit_count)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
                else:
                    out.append(chunk[j])
                    j += 1
    return bytes(out)

# ---------------------------------------------------------------- CFB (OLE) 書き出し

class _Entry:
    def __init__(self, name, etype, data=b"", children=None):
        self.name, self.etype, self.data = name, etype, data
        self.children = children or []
        self.sid = None
        self.left = self.right = self.child = 0xFFFFFFFF
        self.start = 0xFFFFFFFE
        self.size = len(data)


def _sort_key(e):
    return (len(e.name), e.name.upper())


def write_cfb(root_children):
    SEC = 512
    MINI = 64
    root = _Entry("Root Entry", 5, children=root_children)

    # 1. ディレクトリエントリを平坦化（子は長さ→名前順。右リンクで連結）
    entries = []

    def visit(e):
        e.sid = len(entries)
        entries.append(e)
        kids = sorted(e.children, key=_sort_key)
        for k in kids:
            visit(k)
        if kids:
            e.child = kids[0].sid
            for a, b in zip(kids, kids[1:]):
                a.right = b.sid

    visit(root)

    # 2. ミニストリーム（4096 未満）と通常ストリームを振り分け
    mini_stream = bytearray()
    mini_fat = []
    big_streams = []  # (entry, data)
    for e in entries:
        if e.etype != 2:
            continue
        if e.size < 4096:
            nsec = (e.size + MINI - 1) // MINI
            e.start = len(mini_fat) if nsec else 0xFFFFFFFE
            for k in range(nsec):
                mini_fat.append(len(mini_fat) + 1 if k < nsec - 1 else 0xFFFFFFFE)
            mini_stream += e.data + b"\x00" * (nsec * MINI - e.size)
        else:
            big_streams.append(e)

    root.data = bytes(mini_stream)
    root.size = len(mini_stream)

    # 3. セクタ配置: [ディレクトリ][ミニFAT][ミニストリーム][大ストリーム...][FAT]
    sectors = []  # list of bytes(512)
    fat = []

    def add_chain(data):
        if not data:
            return 0xFFFFFFFE
        nsec = (len(data) + SEC - 1) // SEC
        first = len(sectors)
        for k in range(nsec):
            sectors.append(data[k * SEC:(k + 1) * SEC].ljust(SEC, b"\x00"))
            fat.append(first + k + 1 if k < nsec - 1 else 0xFFFFFFFE)
        return first

    dir_bytes_len = ((len(entries) + 3) // 4) * 4 * 128
    dir_nsec = dir_bytes_len // SEC
    dir_start = len(sectors)
    for k in range(dir_nsec):  # 仮確保、後で中身を差し替える
        sectors.append(b"\x00" * SEC)
        fat.append(dir_start + k + 1 if k < dir_nsec - 1 else 0xFFFFFFFE)

    minifat_bytes = b"".join(struct.pack("<I", v) for v in mini_fat)
    minifat_start = add_chain(minifat_bytes)
    minifat_nsec = (len(minifat_bytes) + SEC - 1) // SEC if minifat_bytes else 0

    root.start = add_chain(root.data)
    for e in big_streams:
        e.start = add_chain(e.data)

    # FAT 自身のセクタ
    total = len(sectors)
    fat_nsec = 1
    while (total + fat_nsec) * 4 > fat_nsec * SEC:
        fat_nsec += 1
    fat_start = len(sectors)
    for k in range(fat_nsec):
        fat.append(0xFFFFFFFD)
        sectors.append(b"")
    fat_bytes = b"".join(struct.pack("<I", v) for v in fat)
    fat_bytes = fat_bytes.ljust(fat_nsec * SEC, b"\xff")
    for k in range(fat_nsec):
        sectors[fat_start + k] = fat_bytes[k * SEC:(k + 1) * SEC]

    # 4. ディレクトリ本体
    dir_bytes = bytearray()
    for e in entries:
        name16 = e.name.encode("utf-16-le") + b"\x00\x00"
        rec = bytearray(128)
        rec[0:len(name16)] = name16
        struct.pack_into("<H", rec, 64, len(name16))
        rec[66] = e.etype
        rec[67] = 1  # black
        struct.pack_into("<III", rec, 68, e.left, e.right, e.child)
        if e.etype == 5:
            rec[80:96] = b"\x00" * 16
        struct.pack_into("<IQQ", rec, 100, 0, 0, 0)
        start = e.start if e.etype in (2, 5) else 0
        size = e.size if e.etype in (2, 5) else 0
        struct.pack_into("<I", rec, 116, start if e.etype in (2, 5) else 0)
        struct.pack_into("<Q", rec, 120, size)
        dir_bytes += rec
    pad_entries = dir_bytes_len // 128 - len(entries)
    for _ in range(pad_entries):
        rec = bytearray(128)
        rec[66] = 0
        rec[67] = 1
        struct.pack_into("<III", rec, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
        struct.pack_into("<I", rec, 116, 0)
        dir_bytes += rec
    for k in range(dir_nsec):
        sectors[dir_start + k] = bytes(dir_bytes[k * SEC:(k + 1) * SEC])

    # 5. ヘッダ
    hdr = bytearray(SEC)
    hdr[0:8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
    struct.pack_into("<HHHHHH", hdr, 24, 0x003E, 0x0003, 0xFFFE, 0x0009, 0x0006, 0)
    struct.pack_into("<I", hdr, 44, fat_nsec)
    struct.pack_into("<I", hdr, 48, dir_start)
    struct.pack_into("<I", hdr, 56, 4096)
    struct.pack_into("<I", hdr, 60, minifat_start if minifat_nsec else 0xFFFFFFFE)
    struct.pack_into("<I", hdr, 64, minifat_nsec)
    struct.pack_into("<I", hdr, 68, 0xFFFFFFFE)
    struct.pack_into("<I", hdr, 72, 0)
    difat = [fat_start + k for k in range(fat_nsec)] + [0xFFFFFFFF] * (109 - fat_nsec)
    hdr[76:512] = b"".join(struct.pack("<I", v) for v in difat)

    return bytes(hdr) + b"".join(sectors)

# ---------------------------------------------------------------- VBA プロジェクト

def _rec(rid, payload):
    return struct.pack("<HI", rid, len(payload)) + payload


def _build_dir(modules, references):
    mb = lambda s: s.encode(ENC)
    u16 = lambda s: s.encode("utf-16-le")
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))            # SYSKIND Win32
    d += _rec(0x0002, struct.pack("<I", 0x409))        # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))        # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", CODEPAGE))     # CODEPAGE
    d += _rec(0x0004, mb(PROJECT_NAME))                # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")         # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")         # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))            # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)         # VERSION major=1 minor=0
    d += _rec(0x000C, b"") + _rec(0x003C, b"")         # CONSTANTS
    for name, libid in references:
        d += _rec(0x0016, mb(name)) + _rec(0x003E, u16(name))
        lib = mb(libid)
        d += struct.pack("<HI", 0x000D, 4 + len(lib) + 6) + struct.pack("<I", len(lib)) + lib + b"\x00" * 6
    d += _rec(0x000F, struct.pack("<H", len(modules)))  # MODULES count
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))        # COOKIE
    for name, mtype, _src in modules:
        d += _rec(0x0019, mb(name)) + _rec(0x0047, u16(name))          # NAME
        d += _rec(0x001A, mb(name)) + _rec(0x0032, u16(name))          # STREAMNAME
        d += _rec(0x001C, b"") + _rec(0x0048, b"")                     # DOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))                        # OFFSET
        d += _rec(0x001E, struct.pack("<I", 0))                        # HELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))                   # COOKIE
        d += _rec(0x0021 if mtype == "std" else 0x0022, b"")           # TYPE
        d += _rec(0x002B, b"")                                         # end of module
    d += _rec(0x0010, b"")                                             # terminator
    return bytes(d)


def _build_project_stream(modules):
    lines = ['ID="{%s}"' % str(uuid.uuid4()).upper()]
    for name, mtype, _ in modules:
        lines.append(("Document=%s/&H00000000" if mtype == "doc" else "Module=%s") % name)
    lines += [
        'Name="%s"' % PROJECT_NAME,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="5E5C5EE2A21E6422642264226422"',
        'DPB="ACAEAC10D4CFD5CFD5CF"',
        'GC="ECEEEC50140F150F15F0"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    for name, _, _ in modules:
        lines.append("%s=0, 0, 0, 0, C" % name)
    return ("\r\n".join(lines) + "\r\n").encode(ENC)


def _build_projectwm(modules):
    out = bytearray()
    for name, _, _ in modules:
        out += name.encode(ENC) + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    out += b"\x00\x00"
    return bytes(out)


DOC_HEADER = (
    'Attribute VB_Name = "{name}"\r\n'
    'Attribute VB_Base = "0{{{guid}}}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)


def build_vba_project(std_name, std_source, sheet_codenames):
    modules = [("ThisWorkbook", "doc", DOC_HEADER.format(name="ThisWorkbook", guid="00020819-0000-0000-C000-000000000046"))]
    for cn in sheet_codenames:
        modules.append((cn, "doc", DOC_HEADER.format(name=cn, guid="00020820-0000-0000-C000-000000000046")))
    modules.append((std_name, "std", std_source))

    references = [
        ("stdole", r"*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\Windows\System32\stdole2.tlb#OLE Automation"),
        ("Office", r"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files\Common Files\Microsoft Shared\OFFICE16\MSO.DLL#Microsoft Office 16.0 Object Library"),
    ]

    vba_children = [
        _Entry("dir", 2, compress(_build_dir(modules, references))),
        _Entry("_VBA_PROJECT", 2, b"\xCC\x61\xFF\xFF\x00\x00\x00"),
    ]
    for name, _, src in modules:
        vba_children.append(_Entry(name, 2, compress(src.encode(ENC))))

    root_children = [
        _Entry("VBA", 1, children=vba_children),
        _Entry("PROJECT", 2, _build_project_stream(modules)),
        _Entry("PROJECTwm", 2, _build_projectwm(modules)),
    ]
    return write_cfb(root_children)

# ---------------------------------------------------------------- シート上の実行ボタン

DRAWING_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="MacroButton"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln></xdr:spPr><xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>{label}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"""


def add_button(zin, names, patched, sheet_file, macro, label, cell_from=(14, 1), cell_to=(17, 4)):
    """sheet_file に、クリックで macro を実行する図形ボタンを追加する（0 始まりの列・行で位置指定）"""
    if "<drawing " in patched.get(sheet_file, zin.read(sheet_file).decode("utf-8")):
        return  # 既に図形があるシートには手を付けない
    n = 1
    while "xl/drawings/drawing%d.xml" % n in names or "xl/drawings/drawing%d.xml" % n in patched:
        n += 1
    drawing_path = "xl/drawings/drawing%d.xml" % n
    patched[drawing_path] = DRAWING_XML.format(c1=cell_from[0], r1=cell_from[1], c2=cell_to[0], r2=cell_to[1],
                                               macro=macro, label=label)

    rels_path = sheet_file.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if rels_path in names:
        rels = patched.get(rels_path) or zin.read(rels_path).decode("utf-8")
    else:
        rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = "rId%d" % (max(ids) + 1 if ids else 1)
    rels = rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing%d.xml"/></Relationships>' % (rid, n))
    patched[rels_path] = rels

    xml = patched.get(sheet_file) or zin.read(sheet_file).decode("utf-8")
    tag = '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="%s"/>' % rid
    for marker in ("<legacyDrawing", "<legacyDrawingHF", "<picture", "<oleObjects", "<controls",
                   "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"):
        if marker in xml:
            xml = xml.replace(marker, tag + marker, 1)
            break
    patched[sheet_file] = xml

    ct = patched["[Content_Types].xml"]
    patched["[Content_Types].xml"] = ct.replace("</Types>",
        '<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>' % drawing_path)

# ---------------------------------------------------------------- xlsx → xlsm

def load_bas(path):
    text = open(path, encoding="utf-8-sig").read()
    m = re.search(r'^Attribute VB_Name = "([^"]+)"', text, re.M)
    if not m:
        raise SystemExit("Attribute VB_Name 行が見つかりません: " + path)
    name = m.group(1)
    body = "\n".join(l for l in text.splitlines() if not l.startswith("Attribute VB_Name"))
    src = 'Attribute VB_Name = "%s"\r\n' % name + body.replace("\r\n", "\n").replace("\n", "\r\n")
    if not src.endswith("\r\n"):
        src += "\r\n"
    return name, src


def build_xlsm(src_xlsx, bas_path, out_xlsm, button_macro=None, button_label=None):
    mod_name, mod_src = load_bas(bas_path)
    zin = zipfile.ZipFile(src_xlsx)
    names = zin.namelist()

    # シートの一覧と codeName
    wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
    sheet_files = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    codenames = []
    patched = {}
    for n in sheet_files:
        xml = zin.read(n).decode("utf-8")
        m = re.search(r'<sheetPr[^>]*codeName="([^"]+)"', xml)
        if m:
            cn = m.group(1)
        else:
            cn = "Sheet" + re.search(r"sheet(\d+)\.xml$", n).group(1)
            if "<sheetPr" in xml:
                xml = re.sub(r"<sheetPr", '<sheetPr codeName="%s"' % cn, xml, count=1)
            else:
                xml = re.sub(r"(<worksheet[^>]*>)", r'\1<sheetPr codeName="%s"/>' % cn, xml, count=1)
            patched[n] = xml
        codenames.append(cn)

    if 'codeName=' not in wb_xml.split("<workbookPr")[1].split("/>")[0] if "<workbookPr" in wb_xml else True:
        if "<workbookPr" in wb_xml:
            wb_xml = re.sub(r"<workbookPr", '<workbookPr codeName="ThisWorkbook"', wb_xml, count=1)
        else:
            wb_xml = re.sub(r"(</fileVersion>|<fileVersion[^>]*/>)", r'\1<workbookPr codeName="ThisWorkbook"/>', wb_xml, count=1)
    patched["xl/workbook.xml"] = wb_xml

    ct = zin.read("[Content_Types].xml").decode("utf-8")
    ct = ct.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    if "/xl/vbaProject.bin" not in ct:
        ct = ct.replace("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    patched["[Content_Types].xml"] = ct

    rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = "rId%d" % (max(ids) + 1 if ids else 1)
    rels = rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>' % rid)
    patched["xl/_rels/workbook.xml.rels"] = rels

    if button_macro:
        add_button(zin, names, patched, sheet_files[0], button_macro, button_label or button_macro)

    vba_bin = build_vba_project(mod_name, mod_src, codenames)

    with zipfile.ZipFile(out_xlsm, "w", zipfile.ZIP_DEFLATED) as zout:
        written = set()
        for item in zin.infolist():
            data = patched.get(item.filename, None)
            if data is None:
                zout.writestr(item, zin.read(item.filename))
            else:
                zout.writestr(item, data.encode("utf-8"))
            written.add(item.filename)
        for name, data in patched.items():  # 新規追加したパーツ（drawing など）
            if name not in written:
                zout.writestr(name, data.encode("utf-8"))
        zout.writestr("xl/vbaProject.bin", vba_bin)
    return mod_name, codenames


if __name__ == "__main__":
    if len(sys.argv) not in (4, 6):
        print(__doc__)
        sys.exit(1)
    bm = sys.argv[4] if len(sys.argv) == 6 else None
    bl = sys.argv[5] if len(sys.argv) == 6 else None
    name, cns = build_xlsm(sys.argv[1], sys.argv[2], sys.argv[3], bm, bl)
    print("作成しました:", sys.argv[3], "モジュール:", name, "シート:", cns)

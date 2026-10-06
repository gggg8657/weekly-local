"""XLSX 쓰기·읽기 — 표준 라이브러리(zipfile + XML)만. openpyxl 없이 동작한다.

쓰기: Book().sheet(...) 에 셀(글·숫자·날짜·서식 있는 글 조각), 열 폭, 병합, 틀 고정, 목록 검사(드롭다운), 인쇄 설정, 숨김 시트를 넣고 save().
읽기: read(data) → [{"name", "rows": [[셀 값…]], "rich": {(r,c): [조각…]}, "merged": [...], "hidden"}] — 병합 셀은 왼쪽 위 값으로 채운다,
      날짜 서식 숫자는 date 로, 서식 있는 글(rich text)은 조각별 굵게·색·취소선을 함께 돌려준다.
"""
import datetime
import html
import io
import re
import zipfile
from xml.sax.saxutils import escape

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
EPOCH = datetime.date(1899, 12, 30)


def col_name(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def ref(r, c):
    return f"{col_name(c)}{r + 1}"


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


class Styles:
    """cellXfs 를 서식 묶음(dict) 으로 등록 → 번호. font/fill/border/numFmt/alignment 를 자동으로 모은다."""

    def __init__(self):
        self.fonts = ['<font><sz val="10"/><name val="맑은 고딕"/></font>']
        self.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>']
        self.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>']
        self.numfmts = {}
        self.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
        self.cache = {}

    def font_xml(self, f):
        return ("<font>" + ("<b/>" if f.get("bold") else "") + ("<i/>" if f.get("italic") else "") + ("<strike/>" if f.get("strike") else "")
                + ("<u/>" if f.get("underline") else "") + f'<sz val="{f.get("size", 10)}"/>'
                + (f'<color rgb="FF{f["color"].lstrip("#").upper()}"/>' if f.get("color") else "") + f'<name val="{escape(f.get("name", "맑은 고딕"))}"/></font>')

    def _idx(self, lst, xml):
        if xml not in lst:
            lst.append(xml)
        return lst.index(xml)

    def get(self, st=None):
        st = st or {}
        key = repr(sorted(st.items()))
        if key in self.cache:
            return self.cache[key]
        font = self._idx(self.fonts, self.font_xml(st.get("font") or {}))
        fill = 0
        if st.get("fill"):
            fill = self._idx(self.fills, f'<fill><patternFill patternType="solid"><fgColor rgb="FF{st["fill"].lstrip("#").upper()}"/><bgColor indexed="64"/></patternFill></fill>')
        border = 0
        if st.get("border"):
            b = st["border"] if isinstance(st["border"], dict) else {k: "thin" for k in ("left", "right", "top", "bottom")}
            border = self._idx(self.borders, "<border>" + "".join(f'<{k} style="{b[k]}"><color auto="1"/></{k}>' if b.get(k) else f"<{k}/>"
                                                                      for k in ("left", "right", "top", "bottom")) + "<diagonal/></border>")
        nf = 0
        if st.get("numfmt"):
            if st["numfmt"] not in self.numfmts:
                self.numfmts[st["numfmt"]] = 164 + len(self.numfmts)
            nf = self.numfmts[st["numfmt"]]
        al = st.get("align") or {}
        al_xml = ("<alignment" + "".join(f' {k}="{v}"' for k, v in al.items()) + "/>") if al else ""
        xf = (f'<xf numFmtId="{nf}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0"'
              + (' applyNumberFormat="1"' if nf else "") + (' applyAlignment="1">' + al_xml + "</xf>" if al_xml else "/>"))
        self.xfs.append(xf)
        self.cache[key] = len(self.xfs) - 1
        return self.cache[key]

    def xml(self):
        nf = "".join(f'<numFmt numFmtId="{i}" formatCode="{escape(c)}"/>' for c, i in self.numfmts.items())
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<styleSheet {NS}>' + (f'<numFmts count="{len(self.numfmts)}">{nf}</numFmts>' if nf else "")
                + f'<fonts count="{len(self.fonts)}">{"".join(self.fonts)}</fonts><fills count="{len(self.fills)}">{"".join(self.fills)}</fills>'
                + f'<borders count="{len(self.borders)}">{"".join(self.borders)}</borders>'
                + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                + f'<cellXfs count="{len(self.xfs)}">{"".join(self.xfs)}</cellXfs>'
                + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')


class Sheet:
    def __init__(self, book, name, hidden=False):
        self.book, self.name, self.hidden = book, name, hidden
        self.cells = {}  # (r,c) → (값, 스타일 번호)
        self.widths, self.heights, self.merges, self.validations = {}, {}, [], []
        self.freeze = None
        self.print_fit = False
        self.landscape = False

    def set(self, r, c, value, style=None):
        """value: str | int | float | date | [조각] (조각 = {"t", "bold", "italic", "color", "strike", "size"}) — 서식 있는 글"""
        self.cells[(r, c)] = (value, self.book.styles.get(style) if style is not None else 0)

    def width(self, c, w):
        self.widths[c] = w

    def merge(self, r1, c1, r2, c2):
        self.merges.append(f"{ref(r1, c1)}:{ref(r2, c2)}")

    def validate_list(self, r1, c1, r2, c2, items=None, formula=None):
        f = formula or '"' + ",".join(items) + '"'
        self.validations.append(f'<dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="{ref(r1, c1)}:{ref(r2, c2)}"><formula1>{escape(f)}</formula1></dataValidation>')

    def _cell(self, r, c, v, s):
        a = ref(r, c)
        st = f' s="{s}"' if s else ""
        if v is None or v == "":
            return f'<c r="{a}"{st}/>'
        if isinstance(v, bool):
            v = "O" if v else ""
        if isinstance(v, (datetime.date, datetime.datetime)):
            d = v.date() if isinstance(v, datetime.datetime) else v
            return f'<c r="{a}"{st}><v>{(d - EPOCH).days}</v></c>'
        if isinstance(v, (int, float)):
            return f'<c r="{a}"{st}><v>{v}</v></c>'
        if isinstance(v, list):  # rich text
            runs = ""
            for x in v:
                f = self.book.styles.font_xml(x).replace("<font>", "<rPr>").replace("</font>", "</rPr>").replace("<name ", "<rFont ")
                runs += f'<r>{f}<t xml:space="preserve">{escape(x.get("t", ""))}</t></r>'
            return f'<c r="{a}" t="inlineStr"{st}><is>{runs}</is></c>'
        return f'<c r="{a}" t="inlineStr"{st}><is><t xml:space="preserve">{escape(str(v))}</t></is></c>'

    def xml(self):
        rows = {}
        for (r, c), (v, s) in self.cells.items():
            rows.setdefault(r, []).append((c, v, s))
        sd = ""
        for r in sorted(rows):
            ht = f' ht="{self.heights[r]}" customHeight="1"' if r in self.heights else ""
            sd += f'<row r="{r + 1}"{ht}>' + "".join(self._cell(r, c, v, s) for c, v, s in sorted(rows[r], key=lambda x: x[0])) + "</row>"
        cols = "".join(f'<col min="{c + 1}" max="{c + 1}" width="{w}" customWidth="1"/>' for c, w in sorted(self.widths.items()))
        pane = ""
        if self.freeze:
            r, c = self.freeze
            pane = (f'<pane {"xSplit=" + chr(34) + str(c) + chr(34) + " " if c else ""}ySplit="{r}" topLeftCell="{ref(r, c)}" activePane="bottomLeft" state="frozen"/>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<worksheet {NS}>' + ('<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' if self.print_fit else "")
                + f'<sheetViews><sheetView workbookViewId="0">{pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="16.5"/>'
                + (f"<cols>{cols}</cols>" if cols else "") + f"<sheetData>{sd}</sheetData>"
                + (f'<mergeCells count="{len(self.merges)}">' + "".join(f'<mergeCell ref="{m}"/>' for m in self.merges) + "</mergeCells>" if self.merges else "")
                + (f'<dataValidations count="{len(self.validations)}">{"".join(self.validations)}</dataValidations>' if self.validations else "")
                + '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>'
                + (f'<pageSetup paperSize="9" orientation="{"landscape" if self.landscape else "portrait"}"' + (' fitToWidth="1" fitToHeight="1"' if self.print_fit else "") + "/>")
                + "</worksheet>")


class Book:
    def __init__(self):
        self.styles = Styles()
        self.sheets = []
        self.custom = {}

    def sheet(self, name, hidden=False):
        s = Sheet(self, name, hidden)
        self.sheets.append(s)
        return s

    def save(self):
        n = len(self.sheets)
        ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
              '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
              '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
              + "".join(f'<Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(n))
              + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
        first = next((i for i, s in enumerate(self.sheets) if not s.hidden), 0)
        wb = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {NS}><bookViews><workbookView activeTab="{first}"/></bookViews><sheets>'
              + "".join(f'<sheet name="{escape(s.name)}" sheetId="{i + 1}" r:id="rId{i + 1}"' + (' state="hidden"' if s.hidden else "") + "/>" for i, s in enumerate(self.sheets))
              + "</sheets></workbook>")
        wrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 + "".join(f'<Relationship Id="rId{i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(n))
                 + f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                f'<dc:creator>weekly-local</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created></cp:coreProperties>')
        sheet_xml = [s.xml() for s in self.sheets]  # 스타일이 다 모인 뒤에 styles.xml
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", ct)
            z.writestr("_rels/.rels", rels)
            z.writestr("xl/workbook.xml", wb)
            z.writestr("xl/_rels/workbook.xml.rels", wrels)
            z.writestr("xl/styles.xml", self.styles.xml())
            for i, x in enumerate(sheet_xml):
                z.writestr(f"xl/worksheets/sheet{i + 1}.xml", x)
            z.writestr("docProps/core.xml", core)
        return buf.getvalue()


# ── 읽기 ────────────────────────────────────────────────────────────────
DATE_FMT_IDS = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}


def _txt(x):
    return "".join(re.findall(r"<t(?: [^>]*)?>([^<]*)</t>", x, re.S))


def _unesc(s):
    return html.unescape(s)  # &#48512; 같은 문자 참조까지 (openpyxl·LibreOffice 가 저장한 파일)


def _runs(x):
    """<si>/<is> 안의 서식 조각 → [{"t", "bold", "color", "strike"}] (조각이 없으면 글 하나)"""
    rs = re.findall(r"<r>(.*?)</r>", x, re.S)
    if not rs:
        return [{"t": _unesc(_txt(x))}]
    out = []
    for r in rs:
        pr = (re.search(r"<rPr>(.*?)</rPr>", r, re.S) or [None, ""])[1] if re.search(r"<rPr>", r) else ""
        col = re.search(r'<color rgb="(?:FF)?([0-9A-Fa-f]{6})"', pr)
        out.append({"t": _unesc(_txt(r)), "bold": "<b/>" in pr or "<b " in pr, "strike": "<strike" in pr, "color": "#" + col.group(1).upper() if col else None})
    return out


def read(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()
    rd = lambda n: z.read(n).decode("utf-8") if n in names else ""
    shared = [(_unesc(_txt(si)), _runs(si)) for si in re.findall(r"<si>(.*?)</si>", rd("xl/sharedStrings.xml"), re.S)]
    sty = rd("xl/styles.xml")
    custom_date = {int(i) for i, code in re.findall(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"', sty) if re.search(r"[yYdD]", _unesc(code)) and "[" not in code[:2]}
    xfs = re.findall(r"<xf ([^>]*)/?>", (re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", sty, re.S) or [None, ""])[1]) if sty else []
    fonts = re.findall(r"<font>(.*?)</font>|<font/>", (re.search(r"<fonts[^>]*>(.*?)</fonts>", sty, re.S) or [None, ""])[1], re.S) if sty else []
    xf_date, xf_font = [], []
    for a in xfs:
        m = re.search(r'numFmtId="(\d+)"', a)
        nid = int(m.group(1)) if m else 0
        xf_date.append(nid in DATE_FMT_IDS or nid in custom_date)
        fm = re.search(r'fontId="(\d+)"', a)
        f = fonts[int(fm.group(1))] if fm and int(fm.group(1)) < len(fonts) else ""
        col = re.search(r'<color rgb="(?:FF)?([0-9A-Fa-f]{6})"', f or "")
        xf_font.append({"bold": "<b/>" in (f or ""), "strike": "<strike" in (f or ""), "color": "#" + col.group(1).upper() if col else None, "italic": "<i/>" in (f or "")})
    wb = rd("xl/workbook.xml")
    rels = dict(re.findall(r'<Relationship [^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rd("xl/_rels/workbook.xml.rels")))
    rels.update({k: v for v, k in re.findall(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rd("xl/_rels/workbook.xml.rels"))})
    out = []
    for m in re.finditer(r"<sheet ([^>]*)/>", wb):
        a = m.group(1)
        name = _unesc(re.search(r'name="([^"]*)"', a).group(1))
        rid = re.search(r'r:id="([^"]+)"', a).group(1)
        tgt = rels.get(rid, "")
        path = "xl/" + tgt.lstrip("/").replace("xl/", "", 1) if not tgt.startswith("/") else tgt.lstrip("/")
        x = rd(path)
        cells, rich, fontmap = {}, {}, {}
        for c in re.finditer(r"<c ([^>]*?)(?:/>|>(.*?)</c>)", x, re.S):
            attrs, body = c.group(1), c.group(2) or ""
            r_ = re.search(r'r="([A-Z]+)(\d+)"', attrs)
            if not r_:
                continue
            rr, cc = int(r_.group(2)) - 1, col_index(r_.group(1))
            t = (re.search(r't="(\w+)"', attrs) or [None, ""])[1]
            s = int((re.search(r's="(\d+)"', attrs) or [None, "0"])[1])
            v = re.search(r"<v>([^<]*)</v>", body)
            val = None
            if t == "s" and v:
                val, rich[(rr, cc)] = shared[int(v.group(1))]
            elif t == "inlineStr":
                isx = (re.search(r"<is>(.*?)</is>", body, re.S) or [None, ""])[1]
                val, rich[(rr, cc)] = _unesc(_txt(isx)), _runs(isx)
            elif t == "str" and v:
                val = _unesc(v.group(1))
            elif t == "b" and v:
                val = v.group(1) == "1"
            elif v:
                num = float(v.group(1))
                if s < len(xf_date) and xf_date[s]:
                    val = EPOCH + datetime.timedelta(days=int(num))
                else:
                    val = int(num) if num.is_integer() else num
            if val is not None and val != "":
                cells[(rr, cc)] = val
            if s < len(xf_font):
                fontmap[(rr, cc)] = xf_font[s]
        merged = re.findall(r'<mergeCell ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', x)
        for c1, r1, c2, r2 in merged:  # 병합 셀은 왼쪽 위 값으로 채운다
            v0 = cells.get((int(r1) - 1, col_index(c1)))
            for rr in range(int(r1) - 1, int(r2)):
                for cc in range(col_index(c1), col_index(c2) + 1):
                    if v0 is not None and (rr, cc) not in cells:
                        cells[(rr, cc)] = v0
        nr = max([k[0] for k in cells] or [-1]) + 1
        nc = max([k[1] for k in cells] or [-1]) + 1
        rows = [[cells.get((r, c)) for c in range(nc)] for r in range(nr)]
        out.append({"name": name, "rows": rows, "rich": rich, "font": fontmap, "hidden": 'state="hidden"' in a or 'state="veryHidden"' in a})
    return out

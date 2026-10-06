"""PPTX 쓰기·읽기 — 표준 라이브러리(zipfile + XML)만.

쓰기: deck(slides) — 슬라이드마다 제목 + 글상자 여러 개. 글상자 문단 = {"lvl": 0~3, "runs": [{"t", "bold", "color", "strike", "size"}], "bullet": "∙"}
      lvl 마다 왼쪽 여백·내어쓰기(marL·indent)로 들여쓴다. 숨겨 넣을 JSON 은 customXml/item1.xml(발표 문서 관계 customXml).
읽기: read(data) → [{"title", "boxes": [[{"lvl", "t", "runs"}]]}] — 슬라이드 순서대로, 글상자 위→아래 순서.
"""
import html
import io
import re
import zipfile
from xml.sax.saxutils import escape

A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
EMU = 12700  # 1pt
SW, SH = 12192000, 6858000  # 16:9
FONT = "맑은 고딕"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

THEME = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme {A} name="weekly"><a:themeElements>
<a:clrScheme name="weekly"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F2937"/></a:dk2><a:lt2><a:srgbClr val="F2F2F2"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="FF6600"/></a:accent2>
<a:accent3><a:srgbClr val="0000FF"/></a:accent3><a:accent4><a:srgbClr val="7F7F7F"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="weekly"><a:majorFont><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="weekly"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'''

MASTER = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster {A}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr sz="2400" b="1"/></a:lvl1pPr></p:titleStyle>
<p:bodyStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1400"/></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>'''

LAYOUT = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout {A} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'''


def _rels(items):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(f'<Relationship Id="{i}" Type="{REL}/{t}" Target="{g}"/>' for i, t, g in items) + "</Relationships>")


def _run(r, size):
    sz = int((r.get("size") or size) * 100)
    attrs = f' lang="ko-KR" sz="{sz}"' + (' b="1"' if r.get("bold") else "") + (' strike="sngStrike"' if r.get("strike") else "")
    fill = f'<a:solidFill><a:srgbClr val="{r["color"].lstrip("#").upper()}"/></a:solidFill>' if r.get("color") else ""
    return f'<a:r><a:rPr{attrs} dirty="0">{fill}<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr><a:t>{escape(r.get("t", ""))}</a:t></a:r>'


def _para(p, size):
    lvl = int(p.get("lvl") or 0)
    bullet = p.get("bullet")
    mar = int(p.get("marL", lvl * 0.35 * 914400 + (0.25 * 914400 if bullet else 0)))
    ind = -int(0.25 * 914400) if bullet else 0
    bu = f'<a:buFont typeface="{FONT}"/><a:buChar char="{escape(bullet)}"/>' if bullet else "<a:buNone/>"
    algn = f' algn="{p["align"]}"' if p.get("align") else ""
    runs = "".join(_run(r, size) for r in p.get("runs") or []) or f'<a:endParaRPr lang="ko-KR" sz="{int(size * 100)}"/>'
    return f'<a:p><a:pPr lvl="{lvl}" marL="{mar}" indent="{ind}"{algn}>{bu}</a:pPr>{runs}</a:p>'


def _shape(sid, name, x, y, w, h, paras, size=14, fill=None, line=None, title=False):
    sp = (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
          f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
          + (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>")
          + (f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line else "")
          + f'</p:spPr><p:txBody><a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"><a:normAutofit/></a:bodyPr><a:lstStyle/>'
          + "".join(_para(p, size) for p in paras) + "</p:txBody></p:sp>")
    return sp


def slide_xml(shapes):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {A}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            + "".join(shapes) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')


def deck(slides, embed=None, title="주간보고"):
    """slides = [[shape xml…]] → pptx bytes"""
    n = len(slides)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
          '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
          '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
          '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
          + "".join(f'<Override PartName="/ppt/slides/slide{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(n))
          + '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
    pres = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation {A} saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:sldIdLst>' + "".join(f'<p:sldId id="{256 + i}" r:id="rId{i + 3}"/>' for i in range(n)) + "</p:sldIdLst>"
            f'<p:sldSz cx="{SW}" cy="{SH}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>')
    prels = [("rId1", "slideMaster", "slideMasters/slideMaster1.xml"), ("rId2", "theme", "theme/theme1.xml")] + \
            [(f"rId{i + 3}", "slide", f"slides/slide{i + 1}.xml") for i in range(n)]
    if embed is not None:
        prels.append(("rIdWeekly", "customXml", "../customXml/item1.xml"))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
        z.writestr("ppt/presentation.xml", pres)
        z.writestr("ppt/_rels/presentation.xml.rels", _rels(prels))
        z.writestr("ppt/slideMasters/slideMaster1.xml", MASTER)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", _rels([("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"), ("rId2", "theme", "../theme/theme1.xml")]))
        z.writestr("ppt/slideLayouts/slideLayout1.xml", LAYOUT)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", _rels([("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]))
        z.writestr("ppt/theme/theme1.xml", THEME)
        for i, sh in enumerate(slides):
            z.writestr(f"ppt/slides/slide{i + 1}.xml", slide_xml(sh))
            z.writestr(f"ppt/slides/_rels/slide{i + 1}.xml.rels", _rels([("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")]))
        z.writestr("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                   f'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{escape(title)}</dc:title><dc:creator>weekly-local</dc:creator></cp:coreProperties>')
        if embed is not None:
            z.writestr("customXml/item1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><weekly xmlns="urn:weekly-local"><![CDATA['
                       + embed.replace("]]>", "]]]]><![CDATA[>") + "]]></weekly>")
    return buf.getvalue()


# ── 읽기 ────────────────────────────────────────────────────────────────
def _unesc(s):
    return html.unescape(s)  # &#48512; 같은 문자 참조까지 (다른 프로그램이 다시 저장한 파일)


def read(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    pres = z.read("ppt/presentation.xml").decode("utf-8")
    rels = z.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    rmap = {}
    for m in re.finditer(r"<Relationship ([^>]*)/>", rels):
        a = m.group(1)
        rmap[re.search(r'Id="([^"]+)"', a).group(1)] = re.search(r'Target="([^"]+)"', a).group(1)
    out = []
    for rid in re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres):
        x = z.read("ppt/" + rmap[rid].lstrip("/").replace("ppt/", "")).decode("utf-8")
        shapes = []
        for sp in re.findall(r"<p:sp>.*?</p:sp>", x, re.S):
            off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"', sp)
            is_title = 'type="title"' in sp or 'type="ctrTitle"' in sp
            paras = []
            for p in re.findall(r"<a:p>.*?</a:p>|<a:p/>", sp, re.S):
                ppr = re.search(r"<a:pPr([^>]*)/?>", p)
                lvl = int((re.search(r'lvl="(\d+)"', ppr.group(1)) or [0, "0"])[1]) if ppr else 0
                mar = int((re.search(r'marL="(-?\d+)"', ppr.group(1)) or [0, "0"])[1]) if ppr else 0
                runs = []
                for r in re.findall(r"<a:r>.*?</a:r>", p, re.S):
                    rpr = (re.search(r"<a:rPr([^>]*)>(.*?)</a:rPr>|<a:rPr([^>]*)/>", r, re.S))
                    attrs = (rpr.group(1) or rpr.group(3) or "") if rpr else ""
                    inner = (rpr.group(2) or "") if rpr else ""
                    col = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"', inner)
                    runs.append({"t": _unesc("".join(re.findall(r"<a:t>([^<]*)</a:t>", r))), "bold": bool(re.search(r'\bb="1"', attrs)),
                                 "strike": bool(re.search(r'strike="(sng|dbl)Strike"', attrs)), "color": "#" + col.group(1).upper() if col else None})
                t = "".join(r["t"] for r in runs)
                paras.append({"lvl": lvl, "marL": mar, "t": t, "runs": runs})
            shapes.append({"y": int(off.group(2)) if off else 0, "x": int(off.group(1)) if off else 0, "title": is_title, "paras": paras})
        shapes.sort(key=lambda s: (s["y"], s["x"]))
        out.append({"shapes": shapes})
    return out


def find_embedded(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    if "customXml/item1.xml" in z.namelist():
        m = re.search(r"<!\[CDATA\[(.*)\]\]>", z.read("customXml/item1.xml").decode("utf-8"), re.S)
        if m:
            return m.group(1).replace("]]]]><![CDATA[>", "]]>")
    return None

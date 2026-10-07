"""주간보고 문서 → 미리보기 HTML / 텍스트 / DOCX / HWPX.

문서(doc) 구조 — 개인 보고와 실 취합본이 같다:
  {"title", "subtitle", "rows": [{"label": 부서, "sub": 이름(개인), "items": [항목]}], "expand_mode": paren|note|off, "expand_lang"}
  항목: {"id", "text", "kind": done|plan, "cat": goal|perf|event|etc, "core", "msit", "nobbs", "place", "people"}
표: 왼쪽 열(부서) | 수행한 일 | 향후 2주 계획 — 머리 2줄(부서 칸 세로 병합, '주요 업무 실적 및 계획' 가로 병합).
칸 안: '□ 분류' 머리 → '○ 항목 (@장소, 참석자)'. 핵심 = 주황, 과기정통부 = 파랑(둘 다면 파랑 굵게), BBS 비게시 = 취소선.
bbs=True 면 비게시 항목을 뺀 게시용.

HWPX 는 templates/<양식>.hwpx 를 열어 자리표시자를 바꾼다(양식 교체 가능):
  {{제목}} {{부제}} 같은 글 자리 → 값,  {{본문표}} 가 든 문단 → 생성한 표,  {{약어풀이}} 문단 → '※ 약어' 목록(없으면 문단 삭제).
  글자색·취소선·굵게는 자리표시자 글자모양을 복제해 header.xml 에 새 charPr 로 더한다(양식 글꼴 유지).
기본 양식 default.hwpx 는 kordoc(kordoc-local)의 `generate --preset 통지` 로 만든 문서에 자리표시자만 넣은 것이다.
"""
import datetime
import html
import io
import json
import os
import re
import zipfile
from xml.sax.saxutils import escape

import rules

HERE = os.path.dirname(os.path.abspath(__file__))
TPL_DIR = os.path.join(HERE, "templates")


def templates():
    out = []
    for fn in sorted(os.listdir(TPL_DIR)):
        if fn.endswith(".json"):
            try:
                t = rules.load_json(os.path.join(TPL_DIR, fn), {})
                if t.get("hwpx") and not os.path.exists(os.path.join(TPL_DIR, t["hwpx"])):
                    continue  # 매핑만 있고 양식 파일이 없으면 목록에서 뺀다
                out.append({"id": fn[:-5], "name": t.get("name", fn[:-5])})
            except Exception:
                pass
    return sorted(out, key=lambda t: t["id"] != DEFAULT_TPL)


# 기관 공식 양식(templates/kaeri_weekly.hwpx + .json)은 저장소에 없다 — 있으면 기본, 없으면 임시 양식 'default'
DEFAULT_TPL = "kaeri_weekly" if os.path.exists(os.path.join(TPL_DIR, "kaeri_weekly.json")) and os.path.exists(os.path.join(TPL_DIR, "kaeri_weekly.hwpx")) else "default"


def template(tid=None):
    tid = tid if re.fullmatch(r"[\w\-가-힣]+", tid or "") and os.path.exists(os.path.join(TPL_DIR, (tid or "") + ".json")) else DEFAULT_TPL
    p = os.path.join(TPL_DIR, tid + ".json")
    t = rules.load_json(p if os.path.exists(p) else os.path.join(TPL_DIR, "default.json"), {})
    base = rules.load_json(os.path.join(TPL_DIR, "default.json"), {})
    return {**base, **t, "page": {**rules.PAGE, **base.get("page", {}), **t.get("page", {})}}


# ── 공통 구조: 표(수행/계획) → 행(부서) → 칸 → 문단 → 조각(run) ───────────────────
def style_of(it):
    return {"color": "msit" if it.get("msit") else "core" if it.get("core") else None,
            "bold": bool(it.get("msit") and it.get("core")), "strike": bool(it.get("nobbs"))}


def bullets(tpl):
    b = tpl.get("bullets") or {}
    return b.get("project") or "∙ ", b.get("levels") or ["- ", "· ", "· "]


def yy(d, q="‘"):
    return f"{q}{d.year % 100:02d}.{d.month}.{d.day}."


def short_periods(week_key):
    """공식 양식 표 머리 오른쪽 기간: 수행 "'26.3.3.~'26.3.6."(이번 주 월~금), 계획 "'26.3.9.~'26.3.20."(다음 주 월 ~ 그다음 주 금)"""
    mon = datetime.date.fromisoformat(rules.week_of(week_key)["mon"])
    d = datetime.timedelta
    return {"done": f"{yy(mon)}~{yy(mon + d(days=4), '’')}", "plan": f"{yy(mon + d(days=7))}~{yy(mon + d(days=18), '’')}"}  # 양식: ‘26.3.3.~’26.3.6.


def layout(doc, gl, tpl, bbs=False):
    """→ {"tables": {"done"|"plan": {"title", "period", "rows": [{"label": [문단], "cell": [문단]}]}}, "rows", "notes", "descs", "unknown", "extra", "legend"}
    칸 안 구조(공식 양식): 과제명별로 묶어 '∙ (과제명)'(굵게) 한 줄 → 그 아래 '- (기간) 내용' (하위는 '· ', 단계마다 안으로).
    분류(중점목표·성과·행사·기타)는 출력하지 않는다. 문단 = {"kind": proj|item|abbr|label, "left"(em), "hang"(em, 내어쓰기), "runs"}."""
    pb, lv_b = bullets(tpl)
    ind, item_left = float(tpl.get("indent_em", 2)), float(tpl.get("item_indent_em", 1))
    rows = []
    for r in doc.get("rows") or []:
        its = []
        for blk in rules.blocks(rules.sort_items([i for i in r.get("items") or [] if (i.get("text") or "").strip()])):
            if bbs and blk[0].get("nobbs"):
                continue  # 비게시 상위 항목은 하위까지 통째로 뺀다
            its += [dict(i, kind=blk[0].get("kind"), project=blk[0].get("project") or "") for i in blk if not (bbs and i.get("nobbs"))]
        rows.append((r, its))
    # 과제명 묶음 순서로 다시 줄 세운 뒤 약어는 그 순서로 첫 등장에만 풀이
    grouped = []
    for r, its in rows:
        g = {}
        for kind in rules.KINDS:
            proj = {}
            for blk in rules.blocks([i for i in its if i.get("kind") == kind]):
                proj.setdefault((blk[0].get("project") or "").strip(), []).append(blk)
            g[kind] = proj
        po = (doc.get("project_order") or {}).get(r.get("label") or "") or []
        if po:  # 과제 등록부 순서대로(등록 안 된 과제는 뒤에, 나온 순서)
            g = {kind: dict(sorted(pr.items(), key=lambda kv: po.index(kv[0]) if kv[0] in po else len(po))) for kind, pr in g.items()}
        grouped.append((r, g))
    order = [i for _, g in grouped for k in rules.KINDS for blks in g[k].values() for b in blks for i in b]
    r_sub = {id(i): r.get("sub") or "" for r, g in grouped for k in rules.KINDS for blks in g[k].values() for b in blks for i in b}  # 개인 보고서의 작성자
    texts, notes, descs, unknown, lines = rules.expand_items([i["text"].strip() + rules.ext_suffix(i, doc.get("attendee") or "external", doc.get("writer", "yes") != "no", r_sub.get(id(i), ""), doc.get("online") or "show") for i in order], gl, doc.get("expand_mode") or "note",
                                                             doc.get("levels"), bool(doc.get("desc_block")), doc.get("first_only", True) is not False)
    exp = {id(i): (t, ln) for i, t, ln in zip(order, texts, lines)}
    out = []
    for r, g in grouped:
        cell = {}
        for kind in rules.KINDS:
            paras = []
            for name, blks in g[kind].items():
                paras.append({"kind": "proj", "left": 0, "hang": rules.em(pb),
                              "runs": [{"t": pb}, {"t": f"({name})" if name else rules.NO_PROJECT, "bold": True}]})
                for b in blks:
                    for it in b:
                        d = max(0, min(len(lv_b) - 1, int(it.get("depth") or 0)))
                        text, ln = exp[id(it)]
                        st = style_of(it)
                        left, hang = item_left + d * ind, rules.em(lv_b[d])
                        per = (it.get("period") or "").strip()
                        runs = [{"t": lv_b[d]}] + ([{"t": f"({per}) ", **st}] if per else []) + [{"t": text, **st}]
                        paras.append({"kind": "item", "depth": d, "left": left, "hang": hang, "runs": runs})
                        for a, body in ln:
                            paras.append({"kind": "abbr", "depth": d, "left": left + hang + ind / 2, "hang": 0, "small": True,
                                          "runs": [{"t": a, "bold": True}, {"t": ": " + body}]})
            cell[kind] = paras
        label = [{"kind": "label", "left": 0, "hang": 0, "runs": [{"t": x}]} for x in split_label(r.get("label") or "", tpl)]
        if r.get("sub"):
            label.append({"kind": "label", "left": 0, "hang": 0, "runs": [{"t": f"({r['sub']})"}]})
        out.append({"label": label, "done": cell["done"], "plan": cell["plan"]})
    titles = tpl.get("titles") or rules.KINDS
    periods = short_periods(doc["week"]) if tpl.get("short_period") and doc.get("week") else (doc.get("periods") or {})
    tables = {k: {"title": titles.get(k, rules.KINDS[k]), "period": periods.get(k, ""), "rows": [{"label": r["label"], "cell": r[k]} for r in out]}
              for k in rules.KINDS}
    used = {k for _, its in rows for i in its for k in ("core", "msit", "nobbs") if i.get(k)}
    legend = tpl.get("legend") if used and not bbs and doc.get("legend", True) else ""
    extra = []
    if notes:
        extra.append([{"t": "※ 약어"}])  # 약어 하나당 한 줄, 약어는 볼드
        extra += [[{"t": a, "bold": True}, {"t": ": " + e}] for a, e in notes]
    if descs:
        extra.append([{"t": "※ 용어 설명"}])
        extra += [[{"t": a, "bold": True}, {"t": f" ({h}): {d}"}] for a, h, d in descs]
    remarks = [x.strip() for x in doc.get("remarks") or [] if str(x).strip()]
    return {"tables": tables, "rows": out, "notes": notes, "descs": descs, "unknown": sorted(unknown), "extra": extra, "legend": legend,
            "remarks": remarks, "org": doc.get("org") or "", "title": tpl.get("title") or doc.get("title") or "", "guide": doc.get("guide") or []}


def split_label(name, tpl):
    """실 이름이 칸에 길면 두 줄로(양식의 '인공지능 / 응용연구실' 처럼): 공백이 있으면 공백에서, 없으면 가운데에서"""
    lim = float(tpl.get("label_em", 99))
    if rules.em(name) <= lim:
        return [name]
    if " " in name.strip():
        a, b = name.strip().split(" ", 1)
        return [a, b]
    k = len(name) // 2
    return [name[:k], name[k:]]


def para_text(p):
    return "".join(r["t"] for r in p["runs"])


def para_lines(p, width_em):
    """문단이 칸 폭(em)에서 차지할 줄 수(추정): 왼쪽 여백·내어쓰기를 뺀 폭으로"""
    avail = max(4.0, width_em - p.get("left", 0) - p.get("hang", 0))
    return max(1, -(-rules.em(para_text(p)) // avail))


def estimate(doc, gl, tpl, bbs=False):
    """1쪽 분량 추정 — 두 표(수행·계획) 합산, 과제명 줄·약어 풀이·특기사항 포함"""
    page = tpl["page"]
    lay = layout(doc, gl, tpl, bbs)
    lines = page["head_lines"] + len(lay["extra"]) + (1 + max(1, len(lay["remarks"])) if tpl.get("remarks", True) else 0)
    for t in lay["tables"].values():
        lines += page["table_head"]
        for r in t["rows"]:
            left = sum(para_lines(p, page["cpl"][0]) for p in r["label"])
            lines += max(left, sum(para_lines(p, page["cpl"][1]) for p in r["cell"]) or 1) + page["row_pad"]
    return rules.estimate_result(lines, page)


# ── HTML 미리보기 (공식 양식 모양: 제목 밑줄 → Ⅰ. 소본부 → 표 머리 '1. 수행업무 … 기간' → 실 | 내용) ─────
def _runs_html(p, col):
    o = []
    for r in p["runs"]:
        css = []
        if r.get("color"):
            css.append(f"color:{col[r['color']]}")
        if r.get("bold"):
            css.append("font-weight:700")
        if r.get("strike"):
            css.append("text-decoration:line-through")
        o.append(f'<span style="{";".join(css)}">{html.escape(r["t"])}</span>' if css else html.escape(r["t"]))
    return "".join(o)


def to_html(doc, gl, tpl, bbs=False, full=False):
    lay = layout(doc, gl, tpl, bbs)
    col, ratio = tpl["colors"], tpl["col_ratio"]
    cell = lambda ps: "".join(f'<p class="{p["kind"]}" style="padding-left:{p["left"] + p["hang"]:.2f}em;text-indent:-{p["hang"]:.2f}em">{_runs_html(p, col)}</p>'
                              for p in ps) or '<p class="none">-</p>'
    fill = tpl.get("header_fill", "#F2F2F2")
    h = [f'<div class="wk-doc"><h2 class="wk-title">{html.escape(lay["title"])}</h2>']
    if tpl.get("org_heading") and lay["org"]:
        h.append(f'<h3 class="wk-org">{html.escape(tpl["org_heading"].format(org=lay["org"]))}</h3>')
    elif doc.get("subtitle"):
        h.append(f'<p class="wk-sub">{html.escape(doc["subtitle"])}</p>')
    for k, t in lay["tables"].items():
        h.append(f'<div class="wk-tw"><table class="wk-table" data-kind="{k}"><colgroup>' + "".join(f'<col style="width:{w}%">' for w in ratio) + "</colgroup>"
                 f'<thead><tr class="wk-head" style="background:{fill}"><th class="t">{html.escape(t["title"])}</th><th class="pd">{html.escape(t["period"])}</th></tr></thead><tbody>')
        for r in t["rows"]:
            h.append(f'<tr><td class="lab">{cell(r["label"])}</td><td>{cell(r["cell"])}</td></tr>')
        h.append("</tbody></table></div>")
    if tpl.get("remarks", True):
        h.append(f'<p class="wk-rem-h"><b>{html.escape(tpl.get("remarks_title", "3. 특기 및 애로사항"))}</b> <span>(필요 시)</span></p>')
        h += [f'<p class="wk-rem">- {html.escape(x)}</p>' for x in lay["remarks"]] or ['<p class="wk-rem">-</p>']
    h += ['<p class="wk-notes">' + "".join(f"<b>{html.escape(r['t'])}</b>" if r.get("bold") else html.escape(r["t"]) for r in x) + "</p>" for x in lay["extra"]]
    if lay["legend"]:
        h.append(f'<p class="wk-legend">{html.escape(lay["legend"])}</p>')
    h += [f'<p class="wk-legend">{html.escape(g)}</p>' for g in lay["guide"]]
    h.append("</div>")
    body = "".join(h)
    for mk, tip in ((rules.UNKNOWN, "약어집에 없거나 뜻을 정하지 못한 약어 — 작성자 확인 필요"), (rules.NO_PROJECT, "과제명 없음 — 작성자 확인 필요")):
        body = body.replace(html.escape(mk), f'<mark class="wk-unk" title="{tip}">{html.escape(mk)}</mark>')
    if not full:
        return body
    return f'<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{html.escape(lay["title"] or "주간보고")}</title><style>{DOC_CSS}</style></head><body>{body}</body></html>'


DOC_CSS = ("body{font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;max-width:900px;margin:24px auto;padding:0 16px;color:#111}" + """
.wk-doc{font-family:'휴먼명조','HCR Batang','Batang',serif}.wk-title{text-align:center;font-size:20px;margin:0 0 10px;text-decoration:underline;font-family:'Malgun Gothic',sans-serif}
.wk-org{font-size:17px;margin:8px 0 6px;font-family:'Malgun Gothic',sans-serif}.wk-sub{text-align:center;color:#444;font-size:13px;margin:0 0 10px}
.wk-tw{overflow-x:auto;margin:0 0 12px}.wk-table{border-collapse:collapse;width:100%;font-size:14px;table-layout:fixed;min-width:320px;border:1px solid #333}
.wk-table td{border-left:1px solid #333;border-right:1px solid #333;border-top:1px dashed #666;padding:4px 7px;vertical-align:top;overflow-wrap:anywhere}
.wk-table tr.wk-head th{border-bottom:3px double #333;padding:3px 8px;font-family:'Malgun Gothic',sans-serif}.wk-table th.t{text-align:left;font-size:16px}.wk-table th.pd{text-align:right;font-weight:400;font-size:13.5px}
.wk-table td.lab{text-align:center;vertical-align:middle}.wk-table p{margin:1px 0;line-height:1.6}.wk-table p.abbr{font-size:12.5px;color:#333}
.wk-rem-h{margin:14px 0 2px;font-family:'Malgun Gothic',sans-serif}.wk-rem-h span{font-size:12px}.wk-rem{margin:2px 0 2px 1em}
.wk-notes,.wk-legend{font-size:12.5px;color:#333;margin:3px 0}mark.wk-unk{background:#ffe08a;color:#b42318}""")


# ── 텍스트 ──────────────────────────────────────────────────────────────
def to_text(doc, gl, tpl, bbs=False, marks=True):
    lay = layout(doc, gl, tpl, bbs)
    o = [lay["title"]]
    if tpl.get("org_heading") and lay["org"]:
        o.append(tpl["org_heading"].format(org=lay["org"]))
    elif doc.get("subtitle"):
        o.append(doc["subtitle"])
    o.append("")
    for k, t in lay["tables"].items():
        o.append(f"[{t['title'].strip()}]" + (f" {t['period']}" if t["period"] else ""))
        for r in t["rows"]:
            o.append("■ " + " ".join(para_text(p) for p in r["label"]))
            for p in r["cell"] or []:
                pad = " " * (2 + round(p.get("left", 0)))
                if p["kind"] == "item":
                    st = p["runs"][-1]
                    tag = ("[과기정통부]" if st.get("color") == "msit" else "") + ("[핵심]" if st.get("color") == "core" or st.get("bold") else "") + ("[비게시]" if st.get("strike") else "")
                    o.append(pad + para_text(p) + (" " + tag if marks and tag else ""))
                else:
                    o.append(pad + para_text(p))
            if not r["cell"]:
                o.append("  -")
        o.append("")
    if tpl.get("remarks", True):
        o.append(tpl.get("remarks_title", "3. 특기 및 애로사항") + " (필요 시)")
        o += [f" - {x}" for x in lay["remarks"]] or [" -"]
        o.append("")
    o += ["".join(r["t"] for r in x) for x in lay["extra"]]
    return "\n".join(o).rstrip() + "\n"


# ── DOCX (stdlib zipfile) ───────────────────────────────────────────────
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"""
DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"""


def _styles(pt, font):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
            f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{round(pt * 2)}"/><w:szCs w:val="{round(pt * 2)}"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>'
            '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕"/><w:b/><w:u w:val="single"/><w:sz w:val="32"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Org"><w:name w:val="Org"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>'
            '<w:rPr><w:rFonts w:ascii="Malgun Gothic" w:hAnsi="Malgun Gothic" w:eastAsia="맑은 고딕"/><w:b/><w:sz w:val="30"/></w:rPr></w:style></w:styles>')


def _wrun(r, colors, sz=None):
    rpr = ("<w:b/>" if r.get("bold") else "") + ("<w:strike/>" if r.get("strike") else "") + \
          (f'<w:color w:val="{colors[r["color"]].lstrip("#")}"/>' if r.get("color") else "") + (f'<w:sz w:val="{sz}"/>' if sz else "")
    return f'<w:r>{"<w:rPr>" + rpr + "</w:rPr>" if rpr else ""}<w:t xml:space="preserve">{escape(r["t"])}</w:t></w:r>'


def _wp(runs, colors, jc=None, left=0, hang=0, style=None, sz=None):
    """left·hang 은 twip. 내어쓰기: w:ind left = 여백+내어쓰기, hanging = 내어쓰기 → 첫 줄 글머리는 여백 위치, 둘째 줄은 글머리 뒤에 맞춤"""
    ind = f'<w:ind w:left="{left + hang}" w:hanging="{hang}"/>' if hang else (f'<w:ind w:left="{left}"/>' if left else "")
    ppr = (f'<w:pStyle w:val="{style}"/>' if style else "") + ind + (f'<w:jc w:val="{jc}"/>' if jc else "")
    return f'<w:p>{"<w:pPr>" + ppr + "</w:pPr>" if ppr else ""}{"".join(_wrun(r, colors, sz) for r in runs)}</w:p>'


def to_docx(doc, gl, tpl, bbs=False):
    lay = layout(doc, gl, tpl, bbs)
    colors, pt = tpl["colors"], float(tpl.get("font_pt", 11))
    tw = round(pt * 20)  # 1em(twip)
    total = 9638
    ws = [round(total * x / sum(tpl["col_ratio"])) for x in tpl["col_ratio"]]
    fill = tpl.get("header_fill", "#F2F2F2").lstrip("#")

    def tc(paras, w, head=False, center=False, last=False, jc=None):
        b = ('<w:tcBorders><w:bottom w:val="double" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>' if head else
             f'<w:tcBorders><w:top w:val="dashed" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="{"single" if last else "dashed"}" w:sz="4" w:space="0" w:color="000000"/></w:tcBorders>')
        pr = f'<w:tcW w:w="{w}" w:type="dxa"/>' + b + (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if head else "") + ('<w:vAlign w:val="center"/>' if head or center else "")
        body = "".join(_wp(p["runs"], colors, jc or ("center" if center else None), round(p.get("left", 0) * tw), round(p.get("hang", 0) * tw),
                           sz=round(pt * 2 - 2) if p.get("small") else None) for p in paras) or _wp([{"t": "-"}], colors)
        return f"<w:tc><w:tcPr>{pr}</w:tcPr>{body}</w:tc>"

    body = [_wp([{"t": lay["title"]}], colors, style="Title")]
    if tpl.get("org_heading") and lay["org"]:
        body.append(_wp([{"t": tpl["org_heading"].format(org=lay["org"])}], colors, style="Org"))
    elif doc.get("subtitle"):
        body.append(_wp([{"t": doc["subtitle"]}], colors, "center"))
    for k, t in lay["tables"].items():
        rows = ['<w:tr><w:trPr><w:tblHeader/></w:trPr>' + tc([{"runs": [{"t": t["title"], "bold": True}]}], ws[0], head=True)
                + tc([{"runs": [{"t": t["period"]}]}], ws[1], head=True, jc="right") + "</w:tr>"]
        rows += ["<w:tr>" + tc(r["label"], ws[0], center=True, last=n == len(t["rows"]) - 1) + tc(r["cell"], ws[1], last=n == len(t["rows"]) - 1) + "</w:tr>"
                 for n, r in enumerate(t["rows"])]
        body += [_wp([{"t": ""}], colors),
                 '<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:tblBorders>' % total
                 + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ("top", "left", "bottom", "right", "insideV"))
                 + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'
                 + "".join(f'<w:gridCol w:w="{w}"/>' for w in ws) + "</w:tblGrid>" + "".join(rows) + "</w:tbl>"]
    if tpl.get("remarks", True):
        body.append(_wp([{"t": ""}], colors))
        body.append(_wp([{"t": tpl.get("remarks_title", "3. 특기 및 애로사항") + " ", "bold": True}, {"t": "(필요 시)"}], colors))
        body += [_wp([{"t": f" - {x}"}], colors) for x in lay["remarks"]] or [_wp([{"t": " -"}], colors)]
    body += [_wp(x, colors) for x in lay["extra"]]
    if lay["legend"]:
        body.append(_wp([{"t": lay["legend"]}], colors))
    body += [_wp([{"t": g}], colors, sz=18) for g in lay["guide"]]  # 작성 양식의 안내 줄

    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{"".join(body)}'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:title>{escape(lay["title"])}</dc:title><dc:creator>weekly-local</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created></cp:coreProperties>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CT)
        z.writestr("_rels/.rels", RELS)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/document.xml", xml)
        z.writestr("word/styles.xml", _styles(pt, tpl.get("docx_font") or "Malgun Gothic"))
        z.writestr("docProps/core.xml", core)
    return buf.getvalue()


# ── HWPX (양식 + 자리표시자) ─────────────────────────────────────────────
def _para_at(sec, mark):
    """자리표시자가 든 <hp:p …>…</hp:p> 의 (시작, 끝). 표 안 문단이어도 가장 안쪽 문단."""
    i = sec.find(mark)
    if i < 0:
        return None
    s = sec.rfind("<hp:p ", 0, i)
    e = sec.find("</hp:p>", i) + len("</hp:p>")
    return s, e


class HwpxHeader:
    """header.xml 에 글자모양·문단모양·테두리를 더한다(기존 항목은 그대로, id 는 뒤에 이어서)."""

    def __init__(self, xml):
        self.xml = xml
        self.add = {"charPr": [], "paraPr": [], "borderFill": []}
        self.next = {k: max([int(x) for x in re.findall(rf'<hh:{k} id="(\d+)"', xml)] or [0]) + 1 for k in self.add}
        self.cache = {}

    def _new(self, kind, xml_fn):
        nid = self.next[kind]
        self.next[kind] += 1
        self.add[kind].append(xml_fn(nid))
        return nid

    def char(self, base_id, height, bold=False, color=None, strike=False):
        key = ("c", base_id, height, bold, color, strike)
        if key in self.cache:
            return self.cache[key]
        base = re.search(rf'<hh:charPr id="{base_id}"[ >].*?</hh:charPr>', self.xml, re.S).group(0)

        def mk(nid):
            x = re.sub(r'<hh:charPr id="\d+"', f'<hh:charPr id="{nid}"', base, 1)
            x = re.sub(r'(<hh:charPr [^>]*?)height="\d+"', rf'\g<1>height="{height}"', x, 1)
            x = re.sub(r'(<hh:charPr [^>]*?)textColor="[^"]*"', rf'\g<1>textColor="{color or "#000000"}"', x, 1)
            x = re.sub(r'\s(?:bold|italic)="1"', "", x.split(">", 1)[0]) + ">" + x.split(">", 1)[1]
            x = re.sub(r"<hh:bold/>", "", x)
            if bold:
                x = re.sub(r"(<hh:offset [^>]*/>)", r"\1<hh:bold/>", x, 1)
            stk = f'<hh:strikeout shape="{"SOLID" if strike else "NONE"}" color="#000000"/>'
            if "<hh:strikeout" in x:
                x = re.sub(r"<hh:strikeout [^>]*/>", stk, x, 1)
            elif strike:
                anchors = list(re.finditer(r"<hh:underline [^>]*/>|<hh:bold/>|<hh:italic/>|<hh:offset [^>]*/>", x))
                x = x[:anchors[-1].end()] + stk + x[anchors[-1].end():] if anchors else x.replace("</hh:charPr>", stk + "</hh:charPr>")
            return x
        self.cache[key] = self._new("charPr", mk)
        return self.cache[key]

    def para(self, base_id, align="LEFT", hang=0, spacing=130, left=0):
        """문단모양: 한글 내어쓰기는 intent 음수 — 첫 줄은 left, 둘째 줄부터 left + |intent| (공백으로 흉내 내지 않음)"""
        key = ("p", base_id, align, hang, spacing, left)
        if key in self.cache:
            return self.cache[key]
        base = re.search(rf'<hh:paraPr id="{base_id}"[ >].*?</hh:paraPr>', self.xml, re.S).group(0)

        def mk(nid):
            x = re.sub(r'<hh:paraPr id="\d+"', f'<hh:paraPr id="{nid}"', base, 1)
            x = re.sub(r'(<hh:align horizontal=")\w+', rf"\g<1>{align}", x)
            x = re.sub(r'<hc:intent value="-?\d+"', f'<hc:intent value="{-hang}"', x)
            x = re.sub(r'<hc:left value="-?\d+"', f'<hc:left value="{left}"', x)
            x = re.sub(r'<hc:(prev|next) value="-?\d+"', r'<hc:\1 value="0"', x)
            x = re.sub(r'(<hh:lineSpacing type=")\w+(" value=")\d+', rf"\g<1>PERCENT\g<2>{spacing}", x)
            return x
        self.cache[key] = self._new("paraPr", mk)
        return self.cache[key]

    def border(self, fill=None):
        key = ("b", fill)
        if key in self.cache:
            return self.cache[key]
        side = lambda s: f'<hh:{s}Border type="SOLID" width="0.12 mm" color="#000000"/>'
        brush = f'<hc:fillBrush><hc:winBrush faceColor="{fill}" hatchColor="#000000" alpha="0"/></hc:fillBrush>' if fill else ""
        self.cache[key] = self._new("borderFill", lambda nid: (
            f'<hh:borderFill id="{nid}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            + "".join(side(s) for s in ("left", "right", "top", "bottom")) + '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>' + brush + "</hh:borderFill>"))
        return self.cache[key]

    def result(self):
        x = self.xml
        for kind, group in (("charPr", "charProperties"), ("paraPr", "paraProperties"), ("borderFill", "borderFills")):
            if self.add[kind]:
                x = x.replace(f"</hh:{group}>", "".join(self.add[kind]) + f"</hh:{group}>", 1)
            n = len(re.findall(rf"<hh:{kind} id=", x))
            x = re.sub(rf'<hh:{group} itemCnt="\d+"', f'<hh:{group} itemCnt="{n}"', x, 1)
        return x


def _body_width(sec):
    m = re.search(r'<hp:pagePr [^>]*width="(\d+)"[^>]*height="(\d+)"', sec)
    mg = re.search(r'<hp:margin [^>]*left="(\d+)" right="(\d+)"', sec)
    landscape = re.search(r'<hp:pagePr [^>]*landscape="NARROWLY"', sec)
    w = int(m.group(2) if landscape else m.group(1)) if m else 59528
    return w - (int(mg.group(1)) + int(mg.group(2)) if mg else 11338)


def to_hwpx(doc, gl, tpl, bbs=False):
    if tpl.get("engine") == "form":
        return to_hwpx_form(doc, gl, tpl, bbs)
    lay = layout(doc, gl, tpl, bbs)
    src = os.path.join(TPL_DIR, tpl.get("hwpx") or "default.hwpx")
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i) for i in infos}
    sec_name = next(n for n in files if re.fullmatch(r"Contents/section\d+\.xml", n) and tpl["table"].encode() in files[n])
    sec = files[sec_name].decode("utf-8")
    hd = HwpxHeader(files["Contents/header.xml"].decode("utf-8"))
    at = _para_at(sec, tpl["table"])
    para = sec[at[0]:at[1]]
    base_c = re.search(r'charPrIDRef="(\d+)"', para).group(1)
    base_p = re.search(r'paraPrIDRef="(\d+)"', para).group(1)
    height = int(float(tpl.get("font_pt", 10)) * 100)
    colors = tpl["colors"]
    em = height  # 한글 1자 ≈ 글자 크기

    def run_xml(r):
        cid = hd.char(base_c, height, bool(r.get("bold")), colors.get(r.get("color")) if r.get("color") else None, bool(r.get("strike")))
        return f'<hp:run charPrIDRef="{cid}"><hp:t>{escape(r["t"])}</hp:t></hp:run>'

    def p_xml(p, align):
        pid = hd.para(base_p, align, round(p.get("hang", 0) * em), 130, round(p.get("left", 0) * em))
        return f'<hp:p id="0" paraPrIDRef="{pid}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{"".join(run_xml(r) for r in p["runs"])}</hp:p>'

    width = _body_width(sec)
    ratio = tpl["col_ratio"]
    ws = [int(width * x / sum(ratio)) for x in ratio]
    ws[-1] = width - sum(ws[:-1])
    b_body, b_head = hd.border(), hd.border(tpl.get("header_fill", "#E7EEF7"))
    line_h = round(height * 1.32)  # 줄간격 130%

    def cell_lines(paras, w):
        return max(1, sum(para_lines(p, (w - 566) / em) for p in paras))  # 한글은 행 높이를 내용에 맞춰 늘리므로 적게 잡는 쪽이 빈칸이 덜 생김

    def tc(paras, col, row, w, cspan=1, rspan=1, head=False, center=False, valign="TOP"):
        body = "".join(p_xml(p, "CENTER" if head or center else "JUSTIFY" if p.get("kind") == "item" else "LEFT") for p in paras) \
            or p_xml({"kind": "x", "runs": [{"t": ""}]}, "LEFT")
        h = cell_lines(paras, w) * line_h + 282
        return (f'<hp:tc name="" header="{1 if head else 0}" hasMargin="0" protect="0" editable="1" dirty="0" borderFillIDRef="{b_head if head else b_body}">'
                f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="{"CENTER" if head or center else valign}" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
                f'{body}</hp:subList><hp:cellAddr colAddr="{col}" rowAddr="{row}"/><hp:cellSpan colSpan="{cspan}" rowSpan="{rspan}"/>'
                f'<hp:cellSz width="{w}" height="{h}"/><hp:cellMargin left="283" right="283" top="141" bottom="141"/></hp:tc>'), h

    hp = lambda t: [{"kind": "head", "runs": [{"t": t, "bold": True}]}]
    cols = tpl["columns"]

    def table(t, tid):
        """수행·계획 표 하나: 제목 행(2칸 병합) → '부서 | 주요 내용' → 부서별 행(높이는 내용만큼)"""
        c0, h0 = tc(hp(t["title"] + (f" ({t['period']})" if t["period"] else "")), 0, 0, width, cspan=2, head=True)
        c1, h1 = tc(hp(cols[0]), 0, 1, ws[0], head=True)
        c2, _ = tc(hp(cols[1]), 1, 1, ws[1], head=True)
        trs, total_h = [f"<hp:tr>{c0}</hp:tr>", f"<hp:tr>{c1}{c2}</hp:tr>"], h0 + h1
        for i, r in enumerate(t["rows"], start=2):
            n = max(cell_lines(r["label"], ws[0]), cell_lines(r["cell"], ws[1]), 1)
            cells = []
            for col, (ps, center) in enumerate(((r["label"], True), (r["cell"], False))):
                x, _ = tc(ps or [{"kind": "x", "runs": [{"t": "-"}]}], col, i, ws[col], center=center)
                cells.append(re.sub(r'<hp:cellSz width="(\d+)" height="\d+"/>', rf'<hp:cellSz width="\1" height="{n * line_h + 282}"/>', x))
            total_h += n * line_h + 282
            trs.append("<hp:tr>" + "".join(cells) + "</hp:tr>")
        return (f'<hp:tbl id="{tid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" '
                f'pageBreak="CELL" repeatHeader="1" rowCnt="{len(trs)}" colCnt="2" cellSpacing="0" borderFillIDRef="{b_body}" noAdjust="0">'
                f'<hp:sz width="{width}" widthRelTo="ABSOLUTE" height="{total_h}" heightRelTo="ABSOLUTE" protect="0"/>'
                '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
                '<hp:outMargin left="0" right="0" top="0" bottom="0"/><hp:inMargin left="283" right="283" top="141" bottom="141"/>'
                + "".join(trs) + "</hp:tbl>")

    p_open = re.match(r"<hp:p [^>]*>", para).group(0)
    gap = f'{p_open}<hp:run charPrIDRef="{base_c}"><hp:t></hp:t></hp:run></hp:p>'  # 두 표 사이 빈 줄
    new_para = gap.join(f'{p_open}<hp:run charPrIDRef="{base_c}">{table(t, 1900001 + n)}</hp:run></hp:p>' for n, t in enumerate(lay["tables"].values()))
    sec = sec[:at[0]] + new_para + sec[at[1]:]

    # 약어 각주·범례 문단
    if tpl.get("notes") and tpl["notes"] in sec:
        at2 = _para_at(sec, tpl["notes"])
        extra = list(lay["extra"])
        if lay["legend"]:
            extra.append([{"t": lay["legend"]}])
        if extra:
            p2 = sec[at2[0]:at2[1]]
            c2_ = re.search(r'charPrIDRef="(\d+)"', p2).group(1)
            small = hd.char(c2_, max(800, height - 100))
            p_open2 = re.match(r"<hp:p [^>]*>", p2).group(0)
            sec = sec[:at2[0]] + "".join(p_open2 + "".join(f'<hp:run charPrIDRef="{hd.char(c2_, max(800, height - 100), True) if r.get("bold") else small}"><hp:t>{escape(r["t"])}</hp:t></hp:run>' for r in x) + "</hp:p>" for x in extra) + sec[at2[1]:]
        else:
            sec = sec[:at2[0]] + sec[at2[1]:]
    for mark, key in (tpl.get("fields") or {}).items():
        sec = sec.replace(escape(mark), escape(str(doc.get(key) or "")))
    files[sec_name] = sec.encode("utf-8")
    files["Contents/header.xml"] = hd.result().encode("utf-8")
    if "Preview/PrvText.txt" in files:
        files["Preview/PrvText.txt"] = to_text(doc, gl, tpl, bbs, marks=False)[:1000].encode("utf-8")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for info in infos:  # mimetype 무압축 첫 항목 등 원본 순서·압축 방식 유지
            z.writestr(info, files[info.filename], compress_type=info.compress_type)
    return buf.getvalue()


# ── HWPX (공식 양식 복제: templates/kaeri_weekly.hwpx) ─────────────────────────
def _ptext(p):
    return "".join(re.findall(r"<hp:t>([^<]*)</hp:t>", p))


def _paras(xml):
    return re.findall(r"<hp:p [^>]*>.*?</hp:p>", xml, re.S)


def _tcs(tr):
    return re.findall(r"<hp:tc [^>]*>.*?</hp:tc>", tr, re.S)


def to_hwpx_form(doc, gl, tpl, bbs=False):
    """공식 양식 HWPX 를 열어 그대로 채운다:
    - 표 1(수행업무)·표 2(향후 2주 계획): 머리행 오른쪽 기간을 바꾸고, 실 행(머리행 아래)을 양식 행 서식 그대로 복제해 실 수만큼(가운데 행 = 점선 위아래, 마지막 행 = 아래 실선).
    - 실 칸: 양식의 실 이름 문단(가운데·휴먼명조 12pt) 서식, 길면 두 줄. 내용 칸: 양식의 '∙ (사업)' 문단·' - (기간)' 글자모양을 복제하고
      들여쓰기는 문단모양(왼쪽 여백·내어쓰기)으로, 과제명 굵게, 핵심 주황·과기정통부 파랑·BBS 비게시 취소선은 글자모양 복제 후 색·취소선만 바꿈.
    - 'Ⅰ. 소본부' 제목, '3. 특기 및 애로사항' 아래 ' - ' 를 채우고, 양식 아래의 ※ 작성 지침은 기본으로 뺀다(keep_guide 로 유지).
    - 약어 목록은 맨 끝에 '※ 약어' + 한 줄에 하나(약어 굵게)."""
    lay = layout(doc, gl, tpl, bbs)
    src = os.path.join(TPL_DIR, tpl["hwpx"])
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
        files = {i.filename: z.read(i) for i in infos}
    sec = files["Contents/section0.xml"].decode("utf-8")
    sec = re.sub(r"<hp:linesegarray>.*?</hp:linesegarray>", "", sec, flags=re.S)  # 줄 배치 캐시는 지운다(한글이 다시 조판)
    hd = HwpxHeader(files["Contents/header.xml"].decode("utf-8"))
    colors = tpl["colors"]
    tbls = list(re.finditer(r"<hp:tbl [^>]*>.*?</hp:tbl>", sec, re.S))
    assert len(tbls) >= 2, "양식에서 표 2개(수행·계획)를 찾지 못함"
    # 양식 견본에서 서식 뽑기 (첫 표의 첫 실 행)
    t0 = tbls[0].group(0)
    trs0 = re.findall(r"<hp:tr>.*?</hp:tr>", t0, re.S)
    lab_tc, body_tc = _tcs(trs0[1])
    lab_p = _paras(lab_tc)[0]
    lv1_p = next(p for p in _paras(body_tc) if _ptext(p).strip().startswith("∙"))
    lv2_p = next(p for p in _paras(body_tc) if _ptext(p).strip().startswith("-"))
    LAB_PARA, LAB_CHAR = re.search(r'paraPrIDRef="(\d+)"', lab_p).group(1), re.search(r'charPrIDRef="(\d+)"', lab_p).group(1)
    BODY_PARA = re.search(r'paraPrIDRef="(\d+)"', lv1_p).group(1)
    BODY_CHAR = re.findall(r'charPrIDRef="(\d+)"', lv2_p)[-1]  # '(2.24~3.5)' 글자모양
    height = int(re.search(rf'<hh:charPr id="{BODY_CHAR}" height="(\d+)"', hd.xml).group(1))
    em = height
    small = max(800, height - 200)
    line_h = round(height * float(tpl.get("line_spacing", 1.6)))

    def run_xml(r, base=BODY_CHAR, h=None):
        cid = hd.char(base, h or height, bool(r.get("bold")), colors.get(r.get("color")) if r.get("color") else None, bool(r.get("strike")))
        return f'<hp:run charPrIDRef="{cid}"><hp:t>{escape(r["t"])}</hp:t></hp:run>'

    def p_xml(p, para=BODY_PARA, base=BODY_CHAR, align="JUSTIFY"):
        pid = hd.para(para, align, round(p.get("hang", 0) * em), int(tpl.get("line_pct", 160)), round(p.get("left", 0) * em))
        h = small if p.get("small") else None
        return (f'<hp:p id="0" paraPrIDRef="{pid}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                + "".join(run_xml(r, base, h) for r in p["runs"]) + "</hp:p>")

    def cell_w(tc):
        return int(re.search(r'<hp:cellSz width="(\d+)"', tc).group(1))

    def fill_tc(tc, paras_xml, row, height_):
        tc = re.sub(r"(<hp:subList [^>]*>).*?(</hp:subList>)", lambda m: m.group(1) + paras_xml + m.group(2), tc, count=1, flags=re.S)
        tc = re.sub(r'rowAddr="\d+"', f'rowAddr="{row}"', tc)
        return re.sub(r'(<hp:cellSz width="\d+") height="\d+"', rf'\1 height="{height_}"', tc)

    def lines_in(paras, w):
        return max(1, sum(para_lines(p, (w - 1020) / em) for p in paras))

    def build(tm, t):
        x = tm.group(0)
        trs = re.findall(r"<hp:tr>.*?</hp:tr>", x, re.S)
        head = trs[0]
        head = re.sub(r"<hp:t>[‘'’][^<]*</hp:t>", f"<hp:t>{escape(t['period'])}</hp:t>", head, count=1)
        mid_tr, last_tr = trs[1], trs[-1]
        new, total = [head], int(re.search(r'<hp:cellSz width="\d+" height="(\d+)"', head).group(1))
        for n, r in enumerate(t["rows"] or [{"label": [], "cell": []}]):
            tr = last_tr if n == len(t["rows"]) - 1 else mid_tr
            ltc, btc = _tcs(tr)
            lab_xml = "".join(f'<hp:p id="0" paraPrIDRef="{LAB_PARA}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                              f'<hp:run charPrIDRef="{LAB_CHAR}"><hp:t>{escape(para_text(p))}</hp:t></hp:run></hp:p>' for p in r["label"]) \
                or f'<hp:p id="0" paraPrIDRef="{LAB_PARA}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="{LAB_CHAR}"/></hp:p>'
            body_xml = "".join(p_xml(p) for p in r["cell"]) or p_xml({"runs": [{"t": "-"}]})
            hh = max(len(r["label"]), lines_in(r["cell"], cell_w(btc))) * line_h + 282
            new.append("<hp:tr>" + fill_tc(ltc, lab_xml, n + 1, hh) + fill_tc(btc, body_xml, n + 1, hh) + "</hp:tr>")
            total += hh
        a, b = x.index("<hp:tr>"), x.rindex("</hp:tr>") + len("</hp:tr>")
        x = x[:a] + "".join(new) + x[b:]
        x = re.sub(r'rowCnt="\d+"', f'rowCnt="{len(new)}"', x, count=1)
        return re.sub(r'(<hp:sz width="\d+" widthRelTo="\w+") height="\d+"', rf'\1 height="{total}"', x, count=1)

    parts, pos = [], 0
    for tm, k in zip(tbls[:2], rules.KINDS):
        parts += [sec[pos:tm.start()], build(tm, lay["tables"][k])]
        pos = tm.end()
    sec = "".join(parts) + sec[pos:]
    # 소본부 제목 'Ⅰ. ○○' — 표 1 이 떠 있는 문단의 글
    org = lay["org"] or (doc.get("rows") or [{}])[0].get("label", "")
    sec = re.sub(r"<hp:t>Ⅰ\.[^<]*</hp:t>", f"<hp:t>{escape((tpl.get('org_heading') or 'Ⅰ. {org}').format(org=org))}</hp:t>", sec, count=1)
    if tpl.get("split_title", True):
        # 양식은 'Ⅰ. 소본부' 글과 표 1 을 한 문단에 두고 표를 아래로 띄운다(vertOffset). 보이는 모양은 같게, 제목 문단 → 표 문단으로 나눠
        # 표가 길어져도 제목과 겹치지 않게 한다(표 2 와 같은 방식).
        sec = re.sub(r'(<hp:p [^>]*>)(<hp:run charPrIDRef="\d+">)(<hp:tbl .*?</hp:tbl>)<hp:t>(Ⅰ\.[^<]*)</hp:t></hp:run></hp:p>',
                     lambda m: m.group(1) + m.group(2) + f"<hp:t>{m.group(4)}</hp:t></hp:run></hp:p>" + m.group(1) + m.group(2)
                     + re.sub(r'vertOffset="\d+"', 'vertOffset="219"', m.group(3), count=1) + "<hp:t/></hp:run></hp:p>", sec, count=1, flags=re.S)
    # 3. 특기 및 애로사항 + 그 아래 ' - ' 문단 → 특기사항, 그 뒤(※ 작성 지침)는 기본으로 뺀다
    ps = list(re.finditer(r"<hp:p [^>]*>.*?</hp:p>", sec, re.S))
    i3 = next(n for n, m in enumerate(ps) if "특기 및 애로사항" in _ptext(m.group(0)))
    p3 = ps[i3].group(0)
    if not tpl.get("keep_page_break"):
        p3 = p3.replace('pageBreak="1"', 'pageBreak="0"', 1)  # 양식은 3. 앞에서 쪽을 넘기지만 1쪽 보고라 붙인다
    rem_tpl = ps[i3 + 1].group(0)
    rem_para = re.search(r'paraPrIDRef="(\d+)"', rem_tpl).group(1)
    rem = "".join(f'<hp:p id="0" paraPrIDRef="{rem_para}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{run_xml({"t": " - " + x})}</hp:p>'
                  for x in lay["remarks"]) or f'<hp:p id="0" paraPrIDRef="{rem_para}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">{run_xml({"t": " -"})}</hp:p>'
    tail = "".join(m.group(0) for m in ps[i3 + 2:]) if tpl.get("keep_guide") or doc.get("keep_guide") else ""
    note_base = re.search(r'charPrIDRef="(\d+)"', ps[-1].group(0)).group(1) if tail else BODY_CHAR
    extra = list(lay["extra"]) + ([[{"t": lay["legend"]}]] if lay["legend"] else [])
    # 약어 목록 글자모양: 양식의 '(필요 시)' 글자(맑은 고딕 10pt) — 휴먼명조 굵은 영문은 렌더러에서 폭이 어긋나 겹쳐 보임
    nm = re.search(r'<hp:run charPrIDRef="(\d+)"><hp:t>\(필요 시\)</hp:t>', sec)
    note_char = nm.group(1) if nm else BODY_CHAR
    note_h = int(re.search(rf'<hh:charPr id="{note_char}" height="(\d+)"', hd.xml).group(1)) if nm else small
    notes = "".join(f'<hp:p id="0" paraPrIDRef="{hd.para(BODY_PARA, "LEFT", 0, 140)}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
                    + "".join(run_xml(dict(r, color=None), note_char, note_h) for r in x) + "</hp:p>" for x in extra)
    if not tpl.get("remarks", True):
        p3, rem = "", ""
    sec = sec[:ps[i3].start()] + p3 + rem + tail + notes + sec[ps[-1].end():]
    files["Contents/section0.xml"] = sec.encode("utf-8")
    files["Contents/header.xml"] = hd.result().encode("utf-8")
    if "Preview/PrvText.txt" in files:
        files["Preview/PrvText.txt"] = to_text(doc, gl, tpl, bbs, marks=False)[:1000].encode("utf-8")
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        for info in infos:
            z.writestr(info, files[info.filename], compress_type=info.compress_type)
    return buf.getvalue()


# ── 생성 결과 XML 검사 (selftest·export 후 확인용) ─────────────────────────
def inspect_hwpx(data):
    """→ {"ok", "colors": {색: 쓰인 run 수}, "strike_runs", "merged": [(colSpan,rowSpan)], "rows", "texts"}"""
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()
    assert names[0] == "mimetype" and z.read("mimetype") == b"application/hwp+zip", "mimetype"
    hdr = z.read("Contents/header.xml").decode()
    ET.fromstring(hdr)
    ns = {"hh": "http://www.hancom.co.kr/hwpml/2011/head", "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
    H = ET.fromstring(hdr)
    chars = {}
    for c in H.iter("{%s}charPr" % ns["hh"]):
        st = c.find("hh:strikeout", ns)
        chars[c.get("id")] = {"color": c.get("textColor"), "strike": st is not None and st.get("shape") not in (None, "NONE"),
                              "bold": c.find("hh:bold", ns) is not None}
    for grp, kind in (("charProperties", "charPr"), ("paraProperties", "paraPr"), ("borderFills", "borderFill")):
        g = H.find(f".//hh:{grp}", ns)
        assert int(g.get("itemCnt")) == len(g.findall(f"hh:{kind}", ns)), f"{grp} itemCnt"
    paras = {p.get("id"): p for p in H.iter("{%s}paraPr" % ns["hh"])}
    out = {"colors": {}, "strike_runs": 0, "bold_runs": 0, "merged": [], "rows": 0, "tables": 0, "texts": [], "indents": [], "bold_texts": []}
    hc = "{http://www.hancom.co.kr/hwpml/2011/core}"
    for n in names:
        if not re.fullmatch(r"Contents/section\d+\.xml", n):
            continue
        S = ET.fromstring(z.read(n).decode())
        for r in S.iter("{%s}run" % ns["hp"]):
            t = "".join(x.text or "" for x in r.findall("hp:t", ns))
            if not t.strip():
                continue
            c = chars.get(r.get("charPrIDRef"), {})
            out["colors"][c.get("color")] = out["colors"].get(c.get("color"), 0) + 1
            out["strike_runs"] += c.get("strike", False)
            out["bold_runs"] += c.get("bold", False)
            out["texts"].append((t, c.get("color"), c.get("strike", False)))
            if c.get("bold"):
                out["bold_texts"].append(t)
        for p in S.iter("{%s}p" % ns["hp"]):  # 문단별 (첫 글, 왼쪽 여백, 내어쓰기) — 들여쓰기가 속성으로 들어갔는지
            t = "".join(x.text or "" for r in p.findall("hp:run", ns) for x in r.findall("hp:t", ns))
            pp = paras.get(p.get("paraPrIDRef"))
            if t.strip() and pp is not None and pp.find(f".//{hc}left") is not None:
                out["indents"].append((t[:30], int(pp.find(f".//{hc}left").get("value")), int(pp.find(f".//{hc}intent").get("value"))))
        for tbl in S.iter("{%s}tbl" % ns["hp"]):
            out["tables"] += 1
            out["rows"] += int(tbl.get("rowCnt"))
            trs = tbl.findall("hp:tr", ns)
            assert len(trs) == int(tbl.get("rowCnt")), "rowCnt"
            for tc in tbl.iter("{%s}tc" % ns["hp"]):
                sp = tc.find("hp:cellSpan", ns)
                cs, rs = int(sp.get("colSpan")), int(sp.get("rowSpan"))
                if cs > 1 or rs > 1:
                    out["merged"].append((cs, rs))
        assert "{{" not in z.read(n).decode(), "남은 자리표시자"
    out["ok"] = True
    return out

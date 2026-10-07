"""가져오기·내보내기 — 엑셀 입력 양식, 엑셀(데이터형·보고서형) 내보내기, 아무 형식(HWPX·DOCX·XLSX·HWP·PDF·TXT·MD) 가져오기.

가져오기 순서 (LLM 없음):
 1) 이 도구가 내보낸 파일이면 안에 넣어 둔 항목 JSON(weekly-local.json)을 그대로 → 손실 없는 왕복(과제명·기간·단계·색·취소선·장소·참석자).
 2) 엑셀 입력 양식(머리글: 부서 | 작성자 | 과제명 | 구분 | 시작일 | 종료일 | 내용 | 단계 | 핵심 | 과기정통부 | BBS 비게시 | 장소 | 참석자 | 비고) — 머리글 이름은 느슨하게.
 3) 공식 양식대로 손으로 쓴 문서(HWPX·DOCX·엑셀 보고서형·kordoc 으로 바꾼 HWP/PDF) — '수행업무'·'향후 2주 계획' 표를 찾아 실 칸 | 내용 칸,
    '∙ (과제명)' / '- (기간) 내용' / '· 하위', 들여쓰기 → 단계, 글자색 주황·파랑 → 핵심·과기정통부, 취소선 → BBS 비게시, '약어: 풀이' 줄 → 약어집 후보.
 4) 구조를 못 찾으면 글만 꺼내 '메모'로 돌려준다(화면에서 'LLM 으로 정리'를 눌러야만 LLM 사용).
"""
import base64
import datetime
import html
import io
import json
import os
import re
import subprocess
import tempfile
import zipfile
from xml.sax.saxutils import escape

import render
import rules
import xlsx

EMBED_NAME = "weekly-local.json"  # HWPX: Contents/weekly-local.json · DOCX: customXml/weekly-local.json · XLSX: 숨김 시트 _data
FIELDS = [  # (키, 머리글, 열 폭)
    ("dept", "부서(실)", 16), ("name", "작성자", 9), ("project", "과제명", 18), ("kind", "구분(수행/계획)", 11), ("start", "시작일", 11), ("end", "종료일", 11),
    ("text", "내용", 60), ("depth", "단계(1=상위,2=하위,3=하하위)", 10), ("core", "핵심(O)", 7), ("msit", "과기정통부 보고(O)", 9), ("nobbs", "BBS 비게시(O)", 9),
    ("place", "외부활동 장소", 16), ("people", "외부활동 참석자(우리 연구원)", 16), ("party", "상대 기관·인물", 16), ("note", "비고", 18)]
SYN = {  # 머리글 → 키 (공백·괄호 지우고 소문자, 앞부분 일치)
    "dept": ["부서", "실", "소속", "팀", "부서명", "실명"], "name": ["작성자", "이름", "성명", "담당자", "담당"],
    "project": ["과제명", "과제", "사업명", "사업", "과제사업명", "프로젝트"], "kind": ["구분", "수행계획", "유형", "종류"],
    "start": ["시작일", "시작", "시작일자", "기간시작", "착수일"], "end": ["종료일", "종료", "마감", "마감일", "기한", "종료일자", "완료일"],
    "period": ["기간", "일정", "날짜", "일자"], "text": ["내용", "업무내용", "업무", "실적", "항목", "주요내용", "세부내용"],
    "depth": ["단계", "수준", "레벨", "depth", "깊이"], "core": ["핵심", "핵심사항", "중요"], "msit": ["과기정통부보고", "과기정통부", "과기부", "msit"],
    "nobbs": ["bbs비게시", "비게시", "비공개", "bbs제외", "게시제외"], "place": ["외부활동장소", "장소"], "people": ["외부활동참석자", "참석자", "참석"],
    "party": ["상대기관인물", "상대기관", "상대방", "상대", "협의상대", "외부기관"],
    "note": ["비고", "메모"], "remark": ["특기사항", "특기및애로사항", "애로사항", "특기"]}
EXAMPLE = "예시"


def _norm_head(h):
    h = re.sub(r"\([^)]*\)|（[^）]*）|\[[^\]]*\]", "", str(h or ""))
    return re.sub(r"[\s·/_\-:]", "", h).lower()


def map_headers(row):
    """머리글 줄 → {키: 열 번호}. 긴 동의어부터 맞춘다(‘외부활동 장소’가 ‘장소’보다 먼저)."""
    out = {}
    pairs = sorted(((s, k) for k, ss in SYN.items() for s in ss), key=lambda x: -len(x[0]))
    for c, h in enumerate(row):
        n = _norm_head(h)
        if not n:
            continue
        for s, k in pairs:
            if k not in out and (n == s or n.startswith(s)):
                out[k] = c
                break
    return out


def guess_cat(t):
    """출력에는 안 쓰는 내부 분류(1쪽 압축 우선순위) — 가져온 글에 낱말로 대략"""
    if re.search(r"수상|개최|행사", t):
        return "event"
    if re.search(r"발표|학회|논문|특허|협의|참석|방문|회의|보고서|기술이전|언론", t):
        return "perf"
    return "goal"


# ── 날짜·기간 ───────────────────────────────────────────────────────────
def to_date(v, year):
    """엑셀 날짜·일련번호·글('26.3.3., 3/3, 2026-03-03, 3.3, 3월 3일) → date"""
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return xlsx.EPOCH + datetime.timedelta(days=int(v))
    s = str(v).strip().replace("’", "'").replace("‘", "'")
    for pat, fn in ((r"^(\d{4})[-./]\s?(\d{1,2})[-./]\s?(\d{1,2})\.?$", lambda m: (int(m[1]), int(m[2]), int(m[3]))),
                    (r"^'(\d{2})\.\s?(\d{1,2})\.\s?(\d{1,2})\.?$", lambda m: (2000 + int(m[1]), int(m[2]), int(m[3]))),
                    (r"^(\d{1,2})\s?[./월]\s?(\d{1,2})\s?일?\.?$", lambda m: (year, int(m[1]), int(m[2])))):
        m = re.match(pat, s)
        if m:
            try:
                return datetime.date(*fn(m))
            except ValueError:
                return None
    return None


def period_of(start, end):
    s = f"{start.month}.{start.day}" if start else ""
    e = f"{end.month}.{end.day}" if end else ""
    if s and e:
        return s if s == e else f"{s}~{e}"
    return f"~{e}" if e else s


def kind_of(v):
    s = str(v or "").strip()
    if re.search(r"수행|완료|실적|done|한일|금주|이번", s, re.I):
        return "done"
    if re.search(r"계획|예정|plan|향후|차주|다음", s, re.I):
        return "plan"
    return ""


def yes(v):
    return str(v or "").strip().lower() in ("o", "○", "◯", "y", "yes", "true", "1", "v", "✓", "✔", "예", "있음") or v is True


# ── 엑셀 입력 양식 ───────────────────────────────────────────────────────
GUIDE = [
    "주간보고 엑셀 입력 양식 — 작성법",
    "",
    "1. '작성' 시트에 한 줄에 한 항목을 적습니다. 회색 기울임 '예시' 줄은 지우고 쓰세요(지우지 않아도 올릴 때 건너뜁니다).",
    "2. 구분: '수행'(이번 주 한 일) 또는 '계획'(향후 2주 할 일). 보고서에서 '1. 수행업무' / '2. 향후 2주 계획' 표로 나뉩니다.",
    "3. 과제명: 같은 과제명끼리 자동으로 묶여 '∙ (과제명)' 한 줄 아래 '- (기간) 내용' 으로 나갑니다. '과제목록' 시트에 적어 두면 드롭다운으로 고를 수 있습니다.",
    "   과제명이 비어 있으면 올릴 때 묻습니다.",
    "4. 시작일·종료일: 날짜로 적으면 '(2.24~3.5)' 처럼 기간이 붙습니다. 종료일만 = 마감 '(~3.6)', 같은 날 = '(3.20)'.",
    "5. 단계: 1 = 상위 항목('- '), 2 = 하위('· ', 한 단계 안으로), 3 = 하하위. 하위 줄은 바로 위 상위 항목 아래로 들어갑니다.",
    "6. 표시(O 입력): 핵심 → 주황 글자 / 과기정통부 보고 → 파랑 글자 / BBS 비게시 → 취소선(게시용에서는 빠짐).",
    "7. 외부활동은 장소·참석자 칸에 쓰면 내용 뒤에 '(@장소, 참석자)' 가 자동으로 붙습니다. 참석자는 우리 연구원 사람만 —",
    "   상대 기관·인물(예: 과기정통부 김사무관, KINS 담당자)은 '상대' 칸에 쓰면 문장 안에 들어갑니다. 화상회의는 장소에 '온라인'.",
    "8. 약어는 풀이를 붙이지 않아도 됩니다 — 약어집에 있으면 자동으로 풀이가 붙고, 모르는 약어는 올릴 때 질의됩니다.",
    "",
    "※ 공식 양식 작성 지침",
    " - 소·본부별 1쪽 이내, 아래 사항 위주로: 중점목표 연계 실적·계획 / 주요 연구(경영) 성과 및 대외활동 / 주요 행사(연구원 주관) 및 각종 수상 실적 등",
    " - 핵심사항 주황색, 과기정통부 보고 사항 파란색, BBS 에 게시하지 않을 내용은 취소선",
    " - 모든 약어(코드 이름, 국제회의 명칭 등)는 전체 이름을 풀어서 표기 / 외부활동은 (@장소, 참석자명)",
]


def input_template(projects=None, dept="", name="", week_key=None, prefill=None):
    """빈 입력 양식(.xlsx) — 작성·작성법·과제목록 시트, 드롭다운, 날짜 서식, 틀 고정, 열 폭, 내용 줄바꿈"""
    wk = rules.week_of(week_key or None)
    mon = datetime.date.fromisoformat(wk["mon"])
    b = xlsx.Book()
    ws = b.sheet("작성")
    head = {"font": {"bold": True, "size": 10}, "fill": "#DDE6F2", "border": True, "align": {"horizontal": "center", "vertical": "center", "wrapText": "1"}}
    for c, (k, h, w) in enumerate(FIELDS):
        ws.set(0, c, h, head)
        ws.width(c, w)
    ws.heights[0] = 33
    ex = {"font": {"italic": True, "color": "#808080"}, "border": True, "align": {"vertical": "top", "wrapText": "1"}}
    exd = dict(ex, numfmt="yyyy-mm-dd")
    samples = [(dept or "○○연구실", name or "홍길동", "기본사업", "수행", mon, mon + datetime.timedelta(days=3), "(예시) 노심 해석 대리모델 학습 완료", 1, "O", "", "", "", "", "", "예시 — 지우세요"),
               (dept or "○○연구실", name or "홍길동", "기본사업", "수행", None, None, "(예시) 학습 데이터 1,200 케이스 생성", 2, "", "", "", "", "", "", "예시 — 지우세요"),
               (dept or "○○연구실", name or "홍길동", "전략개발단", "계획", None, mon + datetime.timedelta(days=11), "(예시) KINS 담당자와 인허가 일정 협의", 1, "", "O", "", "대전 KINS", "홍길동", "KINS 담당자", "예시 — 지우세요")]
    for r, row in enumerate(samples, start=1):
        for c, v in enumerate(row):
            ws.set(r, c, v, exd if FIELDS[c][0] in ("start", "end") else ex)
    body = {"border": True, "align": {"vertical": "top"}}
    pre = {4 + n: row for n, row in enumerate(prefill or [])}  # 과제 등록부: 사람·과제마다 한 줄 미리(내용·구분은 비워 둠)
    for r in range(4, 204):
        for c, (k, _, _) in enumerate(FIELDS):
            st = dict(body, numfmt="yyyy-mm-dd") if k in ("start", "end") else (dict(body, align={"vertical": "top", "wrapText": "1"}) if k == "text" else body)
            if r in pre and k in pre[r]:
                ws.set(r, c, pre[r][k], st)
            elif dept and k == "dept":
                ws.set(r, c, dept, st)
            elif name and k == "name":
                ws.set(r, c, name, st)
            else:
                ws.set(r, c, None, st)
    col = {k: c for c, (k, _, _) in enumerate(FIELDS)}
    ws.validate_list(1, col["kind"], 203, col["kind"], ["수행", "계획"])
    ws.validate_list(1, col["depth"], 203, col["depth"], ["1", "2", "3"])
    for k in ("core", "msit", "nobbs"):
        ws.validate_list(1, col[k], 203, col[k], ["O"])
    ws.validate_list(1, col["project"], 203, col["project"], formula="과제목록!$A$2:$A$200")
    ws.freeze = (1, 0)
    g = b.sheet("작성법")
    g.width(0, 120)
    for r, t in enumerate(GUIDE):
        g.set(r, 0, t, {"font": {"bold": r in (0, 12), "size": 12 if r == 0 else 10}, "align": {"wrapText": "1", "vertical": "top"}})
    p = b.sheet("과제목록")
    p.width(0, 30)
    p.set(0, 0, "과제명(여기 적으면 '작성' 시트에서 고를 수 있음)", head)
    for r, x in enumerate(projects or ["기본사업", "전략개발단"], start=1):
        p.set(r, 0, x)
    return b.save()


def _years_of(week_key):
    return datetime.date.fromisoformat(rules.week_of(week_key or None)["mon"]).year


def items_to_rows(groups, week_key):
    """[{dept, name, items}] → 입력 양식 행 목록(dict)"""
    year = _years_of(week_key)
    out = []
    for g in groups:
        for it in g.get("items") or []:
            per = it.get("period") or ""
            a, _, e = per.partition("~")
            st = rules_date(a, year)
            en = rules_date(e, year) if "~" in per else None
            if "~" not in per:
                en = st
            out.append({"dept": g.get("dept", ""), "name": g.get("name", ""), "project": it.get("project", ""),
                        "kind": "수행" if it.get("kind") == "done" else "계획", "start": st if (a or "~" not in per) else None, "end": en,
                        "text": it.get("text", ""), "depth": int(it.get("depth") or 0) + 1, "core": "O" if it.get("core") else "",
                        "msit": "O" if it.get("msit") else "", "nobbs": "O" if it.get("nobbs") else "", "place": it.get("place", ""),
                        "people": it.get("people", ""), "party": it.get("party", ""), "note": ""})
    return out


def rules_date(s, year):
    m = re.match(r"^(\d{1,2})\.(\d{1,2})$", (s or "").strip())
    if not m:
        return None
    try:
        return datetime.date(year, int(m[1]), int(m[2]))
    except ValueError:
        return None


def payload(groups, week_key, org="", remarks=None, kind="report"):
    return {"weekly_local": 1, "kind": kind, "week": week_key, "org": org, "remarks": remarks or [],
            "groups": [{"dept": g.get("dept", ""), "name": g.get("name", ""), "items": g.get("items") or []} for g in groups]}


def _embed_sheet(b, data):
    s = b.sheet("_data", hidden=True)
    js = json.dumps(data, ensure_ascii=False)
    for i in range(0, len(js), 30000):
        s.set(i // 30000, 0, js[i:i + 30000])


def data_xlsx(groups, week_key, org="", remarks=None, projects=None):
    """데이터형(입력 양식과 같은 모양) — 엑셀에서 고쳐 다시 올리기용. 숨김 시트에 항목 JSON."""
    b = xlsx.Book()
    ws = b.sheet("작성")
    head = {"font": {"bold": True}, "fill": "#DDE6F2", "border": True, "align": {"horizontal": "center", "vertical": "center", "wrapText": "1"}}
    for c, (k, h, w) in enumerate(FIELDS):
        ws.set(0, c, h, head)
        ws.width(c, w)
    for r, row in enumerate(items_to_rows(groups, week_key), start=1):
        for c, (k, _, _) in enumerate(FIELDS):
            st = {"border": True, "align": {"vertical": "top", "wrapText": "1"} if k == "text" else {"vertical": "top"}}
            if k in ("start", "end"):
                st["numfmt"] = "yyyy-mm-dd"
            ws.set(r, c, row.get(k), st)
    n = max(len(items_to_rows(groups, week_key)) + 50, 100)
    col = {k: c for c, (k, _, _) in enumerate(FIELDS)}
    ws.validate_list(1, col["kind"], n, col["kind"], ["수행", "계획"])
    ws.validate_list(1, col["depth"], n, col["depth"], ["1", "2", "3"])
    for k in ("core", "msit", "nobbs"):
        ws.validate_list(1, col[k], n, col[k], ["O"])
    ws.freeze = (1, 0)
    if remarks:
        rs = b.sheet("특기사항")
        rs.width(0, 100)
        rs.set(0, 0, "3. 특기 및 애로사항", head)
        for i, x in enumerate(remarks, start=1):
            rs.set(i, 0, x)
    g = b.sheet("작성법")
    g.width(0, 120)
    for r, t in enumerate(GUIDE):
        g.set(r, 0, t, {"align": {"wrapText": "1"}})
    pl = b.sheet("과제목록")
    pl.width(0, 30)
    pl.set(0, 0, "과제명", head)
    for r, x in enumerate(projects or sorted({it.get("project") for g in groups for it in g.get("items") or [] if it.get("project")}), start=1):
        pl.set(r, 0, x)
    ws.validate_list(1, col["project"], n, col["project"], formula="과제목록!$A$2:$A$200")
    _embed_sheet(b, payload(groups, week_key, org, remarks, "data"))
    return b.save()


def report_xlsx(doc, gl, tpl, bbs=False, embed=None):
    """보고서형 — 공식 양식과 같은 배치(제목·Ⅰ.소본부·1.수행업무/2.향후 2주 계획 표·실 칸|내용 칸 줄바꿈), 과제명 굵게·주황·파랑·취소선은 서식 있는 글로,
    약어 줄 굵게, A4 1쪽 맞춤 인쇄. 숨김 시트에 항목 JSON."""
    lay = render.layout(doc, gl, tpl, bbs)
    col = tpl["colors"]
    b = xlsx.Book()
    ws = b.sheet("보고서")
    ws.width(0, 15)
    ws.width(1, 86)
    ws.print_fit = True
    r = 0
    ws.set(r, 0, lay["title"], {"font": {"bold": True, "underline": True, "size": 16}, "align": {"horizontal": "center"}})
    ws.merge(r, 0, r, 1)
    ws.heights[r] = 26
    r += 1
    if tpl.get("org_heading") and lay["org"]:
        ws.set(r, 0, tpl["org_heading"].format(org=lay["org"]), {"font": {"bold": True, "size": 14}})
        ws.merge(r, 0, r, 1)
        r += 1
    dbl = {"left": "thin", "right": None, "top": "thin", "bottom": "double"}
    for k, t in lay["tables"].items():
        r += 1
        ws.set(r, 0, t["title"].strip(), {"font": {"bold": True, "size": 13}, "fill": "#F2F2F2", "border": dbl})
        ws.set(r, 1, t["period"], {"font": {"size": 11}, "fill": "#F2F2F2", "border": {"left": None, "right": "thin", "top": "thin", "bottom": "double"},
                                    "align": {"horizontal": "right"}})
        r += 1
        for n, row in enumerate(t["rows"]):
            last = n == len(t["rows"]) - 1
            bd = {"left": "thin", "right": "thin", "top": "dashed", "bottom": "thin" if last else "dashed"}
            ws.set(r, 0, "\n".join(render.para_text(p) for p in row["label"]), {"font": {"size": 11}, "border": bd, "align": {"horizontal": "center", "vertical": "center", "wrapText": "1"}})
            runs, lines = [], 0
            for i, p in enumerate(row["cell"]):
                pad = " " * round(p.get("left", 0) * 2)
                runs.append({"t": ("\n" if i else "") + pad, "size": 11})
                for x in p["runs"]:
                    runs.append({"t": x["t"], "bold": bool(x.get("bold")), "strike": bool(x.get("strike")), "size": 10 if p.get("small") else 11,
                                 "color": col.get(x.get("color")) if x.get("color") else None})
                lines += max(1, -(-rules.em(render.para_text(p)) // 60))
            ws.set(r, 1, runs or "-", {"border": bd, "align": {"vertical": "top", "wrapText": "1"}})
            ws.heights[r] = max(18, lines * 15 + 4)
            r += 1
    if lay["extra"]:  # ※ 약어 — 계획 표 아래, 특기사항 앞
        r += 1
    for x in lay["extra"]:
        ws.set(r, 0, [{"t": y["t"], "bold": bool(y.get("bold")), "size": 10} for y in x])
        ws.merge(r, 0, r, 1)
        r += 1
    if tpl.get("remarks", True):
        r += 1
        ws.set(r, 0, [{"t": tpl.get("remarks_title", "3. 특기 및 애로사항") + " ", "bold": True, "size": 13}, {"t": "(필요 시)", "size": 10}])
        ws.merge(r, 0, r, 1)
        r += 1
        for x in lay["remarks"] or [""]:
            ws.set(r, 0, f" - {x}" if x else " -", {"font": {"size": 11}})
            ws.merge(r, 0, r, 1)
            r += 1
    if embed:
        _embed_sheet(b, embed)
    return b.save()


# ── 넣어 둔 JSON (HWPX·DOCX) ─────────────────────────────────────────────
def embed_zip(data, kind, obj):
    """내보낸 HWPX/DOCX 에 항목 JSON 을 넣는다. HWPX: Contents/weekly-local.json (manifest 에는 안 올림 — 한글은 모르는 파일을 무시),
    DOCX: customXml/weekly-local.json + 문서 관계(customXml) — 워드는 사용자 데이터로 둔다."""
    zin = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    js = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    with zipfile.ZipFile(buf, "w") as z:
        for info in zin.infolist():
            b = zin.read(info)
            if kind == "docx" and info.filename == "word/_rels/document.xml.rels":
                b = b.decode().replace("</Relationships>", '<Relationship Id="rIdWeekly" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" '
                                       'Target="../customXml/item1.xml"/></Relationships>').encode()
            z.writestr(info, b, compress_type=info.compress_type)
        if kind == "hwpx":
            z.writestr("Contents/" + EMBED_NAME, js, compress_type=zipfile.ZIP_DEFLATED)
        else:
            z.writestr("customXml/item1.xml", ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><weekly xmlns="urn:weekly-local"><![CDATA['
                                               + js.decode().replace("]]>", "]]]]><![CDATA[>") + "]]></weekly>").encode(), compress_type=zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def find_embedded(name, data):
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        return None
    names = z.namelist()
    if "Contents/" + EMBED_NAME in names:
        return json.loads(z.read("Contents/" + EMBED_NAME).decode("utf-8"))
    if "customXml/item1.xml" in names:
        m = re.search(r"<!\[CDATA\[(.*)\]\]>", z.read("customXml/item1.xml").decode("utf-8"), re.S)
        if m and '"weekly_local"' in m.group(1):
            return json.loads(m.group(1).replace("]]]]><![CDATA[>", "]]>"))
    if "ppt/presentation.xml" in names:
        js = pptx.find_embedded(data)
        return json.loads(js) if js and '"weekly_local"' in js else None
    if "xl/workbook.xml" in names:
        for sh in xlsx.read(data):
            if sh["name"] == "_data":
                js = "".join(str(r[0] or "") for r in sh["rows"] if r)
                if '"weekly_local"' in js:
                    return json.loads(js)
    return None


# ── 줄 단위 해석 (공식 양식 칸 안) ───────────────────────────────────────────
ORANGE, BLUE = ("FF6600", "FF6633", "E36C09", "F79646", "FFC000", "ED7D31", "FF9900"), ("0000FF", "0070C0", "0066FF", "2F5597", "0000CC", "4472C4")


def color_flag(c):
    c = (c or "").lstrip("#").upper()
    if not c or c in ("000000", "AUTO"):
        return None
    if c in ORANGE:
        return "core"
    if c in BLUE:
        return "msit"
    try:
        rr, gg, bb = int(c[:2], 16), int(c[2:4], 16), int(c[4:], 16)
    except ValueError:
        return None
    if rr > 180 and 60 < gg < 200 and bb < 90:
        return "core"
    if bb > 150 and rr < 90 and gg < 140:
        return "msit"
    return None


PROJ_RE = re.compile(r"^[∙•●◦ㆍ·]?\s*[\(（\[]([^)）\]]+)[\)）\]]\s*$")
ITEM_RE = re.compile(r"^([-–―‐·ㆍ∙•○◦*]|\d+\))\s*(.*)$")
PER_RE = re.compile(r"^\(\s*(~?\s*\d{1,2}\s*[./]\s*\d{1,2}\.?(?:\s*~\s*(?:\d{1,2}\s*[./]\s*)?\d{1,2}\.?)?)\s*\)\s*")
ABBR_LINE = re.compile(r"^([A-Za-z][A-Za-z0-9&/\-]{1,15})\s*:\s*(.+)$")


def parse_lines(lines, kind, year):
    """칸 안 문단 [{"t", "left"(상대 들여쓰기 숫자), "flags": {core,msit,nobbs}, "bold"}] → (항목, 약어 후보)"""
    items, abbrs, project = [], [], ""
    lefts = sorted({round(x.get("left", 0)) for x in lines if x["t"].strip() and not PROJ_RE.match(x["t"].strip())})
    for x in lines:
        t = x["t"].strip()
        if not t or t == "-":
            continue
        pm = PROJ_RE.match(t)
        if pm and (t[:1] in "∙•●◦ㆍ" or x.get("bold")):  # '∙ (과제명)' 또는 굵은 '(과제명)' 한 줄
            project = pm.group(1).strip()
            if project == rules.NO_PROJECT.strip("[]"):
                project = ""
            continue
        am = ABBR_LINE.match(t)
        if am and not ITEM_RE.match(t) and (x.get("small") or x.get("abbr_bold")):
            abbrs.append((am.group(1), am.group(2).strip()))
            continue
        im = ITEM_RE.match(t)
        bullet = im.group(1) if im else ""
        body = im.group(2) if im else t
        depth = 0
        if bullet in ("·", "ㆍ", "◦"):
            depth = 1
        if lefts and len(lefts) > 1:
            lv = lefts.index(round(x.get("left", 0))) if round(x.get("left", 0)) in lefts else 0
            depth = max(depth, min(2, lv))
        per = ""
        pm2 = PER_RE.match(body)
        if pm2:
            per, body = rules.norm_period(pm2.group(1), year), body[pm2.end():]
        if body.strip() in (PLACE_TEXT, "내용", "(기간)", "") or re.fullmatch(r"\(기간\)\s*내용", body.strip()):
            continue  # 작성 양식의 빈 자리
        it = {"text": body.strip(), "kind": kind, "project": "" if project == PLACE_PROJ else project, "period": per, "depth": depth, "cat": guess_cat(body)}
        for k in (x.get("flags") or {}):
            it[k] = True
        if depth > 0 and not items:
            it["depth"] = 0
        items.append(it)
    return items, abbrs


def _kind_of_header(t):
    t = re.sub(r"\s+", "", t)
    if "계획" in t:
        return "plan"
    if "수행" in t or "실적" in t or "추진" in t:
        return "done"
    return ""


def _remarks_from(paras):
    out, on = [], False
    for t in paras:
        s = t.strip()
        if "특기" in s and "애로" in s:
            on = True
            continue
        if on:
            if s.startswith("※") or not s:
                if s.startswith("※"):
                    break
                continue
            s = re.sub(r"^[-–·]\s*", "", s).strip()
            if s:
                out.append(s)
    return out


def _org_from(paras):
    for t in paras:
        m = re.match(r"^[\s#*]*[ⅠⅡⅢⅣⅤI]+\.\s*(.+?)[\s*]*$", t)
        if m:
            return m.group(1)
    return ""


def _abbr_block(paras):
    out, on = [], False
    for t in paras:
        s = t.strip()
        if s.startswith("※") and "약어" in s:
            on = True
            continue
        if "특기" in s and "애로" in s:  # 약어 목록은 '3. 특기 및 애로사항' 앞에서 끝남(특기사항 줄을 약어로 읽지 않게)
            on = False
            continue
        if on:
            m = ABBR_LINE.match(s)
            if m:
                out.append((m.group(1), m.group(2).strip()))
            elif s.startswith("※"):
                on = s.find("용어") >= 0
    return out


# ── HWPX 해석 ──────────────────────────────────────────────────────────
def parse_hwpx(data, year):
    z = zipfile.ZipFile(io.BytesIO(data))
    hd = z.read("Contents/header.xml").decode("utf-8")
    chars, paras = {}, {}
    for m in re.finditer(r'<hh:charPr id="(\d+)"([^>]*)>(.*?)</hh:charPr>', hd, re.S):
        col = re.search(r'textColor="#?([0-9A-Fa-f]{6})"', m.group(2))
        st = re.search(r'<hh:strikeout shape="(\w+)"', m.group(3))
        chars[m.group(1)] = {"color": col.group(1) if col else None, "strike": bool(st and st.group(1) != "NONE"), "bold": "<hh:bold/>" in m.group(3),
                             "h": int((re.search(r'height="(\d+)"', m.group(2)) or [0, "1000"])[1])}
    for m in re.finditer(r'<hh:paraPr id="(\d+)"[^>]*>(.*?)</hh:paraPr>', hd, re.S):
        lf = re.search(r'<hc:left value="(-?\d+)"', m.group(2))
        paras[m.group(1)] = int(lf.group(1)) if lf else 0
    secs = sorted(n for n in z.namelist() if re.fullmatch(r"Contents/section\d+\.xml", n))
    groups, abbrs, outer_paras, found = {}, [], [], False

    def p_info(p):
        runs = re.findall(r'<hp:run charPrIDRef="(\d+)"[^>]*>(.*?)</hp:run>|<hp:run charPrIDRef="(\d+)"[^>]*/>', p, re.S)
        t, flags, bold_all, first_bold, small = "", {}, True, None, True
        for cid, body, _ in runs:
            s = html.unescape("".join(re.findall(r"<hp:t>([^<]*)</hp:t>", re.sub(r"<hp:tbl .*?</hp:tbl>", "", body, flags=re.S))))
            if not s:
                continue
            c = chars.get(cid, {})
            if s.strip() and s.strip() not in ("-", "·", "∙"):
                f = color_flag(c.get("color"))
                if f:
                    flags[f] = True
                    if f == "msit" and c.get("bold"):
                        flags["core"] = True  # 핵심 + 과기정통부 = 파랑 굵게
                if c.get("strike"):
                    flags["nobbs"] = True
                bold_all = bold_all and c.get("bold", False)
                if first_bold is None:
                    first_bold = c.get("bold", False)
                small = small and c.get("h", 1000) < 1050
            t += s
        pid = re.search(r'paraPrIDRef="(\d+)"', p).group(1)
        return {"t": t, "left": paras.get(pid, 0) / 100, "flags": flags, "bold": bold_all and bool(t.strip()), "abbr_bold": bool(first_bold), "small": small}

    for n in secs:
        sec = z.read(n).decode("utf-8")
        tbls = list(re.finditer(r"<hp:tbl .*?</hp:tbl>", sec, re.S))
        rest = re.sub(r"<hp:tbl .*?</hp:tbl>", "", sec, flags=re.S)
        outer_paras += [p_info(p)["t"] for p in re.findall(r"<hp:p [^>]*>.*?</hp:p>", rest, re.S)]
        for tm in tbls:
            trs = re.findall(r"<hp:tr>.*?</hp:tr>", tm.group(0), re.S)
            if not trs:
                continue
            head = " ".join(p_info(p)["t"] for p in re.findall(r"<hp:p [^>]*>.*?</hp:p>", trs[0], re.S))
            kind = _kind_of_header(head)
            if not kind:
                continue
            found = True
            for tr in trs[1:]:
                tcs = re.findall(r"<hp:tc .*?</hp:tc>", tr, re.S)
                if len(tcs) < 2:
                    continue
                lab = [p_info(p)["t"].strip() for p in re.findall(r"<hp:p [^>]*>.*?</hp:p>", tcs[0], re.S)]
                if "".join(lab) in ("부서", "실", "부서명") or "주요내용" in "".join(lab).replace(" ", ""):
                    continue
                label = re.sub(r"\s+", "", "".join(x for x in lab if not x.startswith("(")))
                who = next((x.strip("() ") for x in lab if x.startswith("(")), "")
                lines = [p_info(p) for p in re.findall(r"<hp:p [^>]*>.*?</hp:p>", tcs[1], re.S)]
                items, ab = parse_lines(lines, kind, year)
                abbrs += ab
                g = groups.setdefault((label, who), {"dept": label, "name": who, "items": []})
                g["items"] += [x for x in items if x["text"] and x["text"] != "-"]
    abbrs += _abbr_block(outer_paras)
    return found, list(groups.values()), abbrs, _remarks_from(outer_paras), _org_from(outer_paras), "\n".join(outer_paras)


# ── DOCX 해석 ──────────────────────────────────────────────────────────
def parse_docx(data, year):
    z = zipfile.ZipFile(io.BytesIO(data))
    doc = z.read("word/document.xml").decode("utf-8")

    def p_info(p):
        t, flags, bold_all, first_bold, small = "", {}, True, None, True
        for r in re.findall(r"<w:r>.*?</w:r>|<w:r .*?</w:r>", p, re.S):
            s = html.unescape("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", r)))
            if not s:
                continue
            pr = (re.search(r"<w:rPr>(.*?)</w:rPr>", r, re.S) or [None, ""])[1] if "<w:rPr>" in r else ""
            if s.strip() and s.strip() not in ("-", "·", "∙"):
                col = re.search(r'<w:color w:val="([0-9A-Fa-f]{6})"', pr)
                f = color_flag(col.group(1) if col else None)
                if f:
                    flags[f] = True
                    if f == "msit" and ("<w:b/>" in pr or '<w:b w:val="1"' in pr):
                        flags["core"] = True
                if "<w:strike/>" in pr or '<w:strike w:val="1"' in pr or '<w:strike w:val="true"' in pr:
                    flags["nobbs"] = True
                b = "<w:b/>" in pr or '<w:b w:val="1"' in pr
                bold_all = bold_all and b
                if first_bold is None:
                    first_bold = b
                sz = re.search(r'<w:sz w:val="(\d+)"', pr)
                small = small and bool(sz and int(sz.group(1)) < 21)
            t += s
        ind = re.search(r'<w:ind [^>]*w:left="(\d+)"', p)
        hang = re.search(r'<w:ind [^>]*w:hanging="(\d+)"', p)
        left = (int(ind.group(1)) - (int(hang.group(1)) if hang else 0)) / 20 if ind else 0
        return {"t": t, "left": left, "flags": flags, "bold": bold_all and bool(t.strip()), "abbr_bold": bool(first_bold), "small": small}

    groups, abbrs, found = {}, [], False
    tbls = re.findall(r"<w:tbl>.*?</w:tbl>", doc, re.S)
    outer = re.sub(r"<w:tbl>.*?</w:tbl>", "", doc, flags=re.S)
    outer_paras = [p_info(p)["t"] for p in re.findall(r"<w:p[ >].*?</w:p>", outer, re.S)]
    for tb in tbls:
        trs = re.findall(r"<w:tr[ >].*?</w:tr>", tb, re.S)
        if not trs:
            continue
        head = " ".join(p_info(p)["t"] for p in re.findall(r"<w:p[ >].*?</w:p>", trs[0], re.S))
        kind = _kind_of_header(head)
        if not kind:
            continue
        found = True
        for tr in trs[1:]:
            tcs = re.findall(r"<w:tc>.*?</w:tc>", tr, re.S)
            if len(tcs) < 2:
                continue
            lab = [p_info(p)["t"].strip() for p in re.findall(r"<w:p[ >].*?</w:p>", tcs[0], re.S)]
            if "".join(lab) in ("부서", "실") or "주요내용" in "".join(lab).replace(" ", ""):
                continue
            label = re.sub(r"\s+", "", "".join(x for x in lab if not x.startswith("(")))
            who = next((x.strip("() ") for x in lab if x.startswith("(")), "")
            items, ab = parse_lines([p_info(p) for p in re.findall(r"<w:p[ >].*?</w:p>", tcs[1], re.S)], kind, year)
            abbrs += ab
            g = groups.setdefault((label, who), {"dept": label, "name": who, "items": []})
            g["items"] += [x for x in items if x["text"] and x["text"] != "-"]
    abbrs += _abbr_block(outer_paras)
    return found, list(groups.values()), abbrs, _remarks_from(outer_paras), _org_from(outer_paras), "\n".join(outer_paras)


# ── XLSX 해석 (입력 양식 · 보고서형) ─────────────────────────────────────────
def parse_xlsx_table(sheets, year, week_key):
    """입력 양식 모양의 시트 → (찾음, 행 목록[{row, dept, name, item, errors, warnings}], 특기사항)"""
    for sh in sheets:
        if sh["hidden"] or sh["name"] in ("작성법", "과제목록", "_data"):
            continue
        rows = sh["rows"]
        for hi, row in enumerate(rows[:10]):
            hm = map_headers([str(x or "") for x in row])
            if "text" in hm and ("kind" in hm or "project" in hm):
                return True, _rows_from_table(rows[hi + 1:], hm, year, week_key, hi + 2), _remark_sheet(sheets)
    return False, [], []


def _remark_sheet(sheets):
    for sh in sheets:
        if sh["name"].startswith("특기"):
            return [str(r[0]).strip() for r in sh["rows"][1:] if r and r[0] and str(r[0]).strip()]
    return []


def _rows_from_table(rows, hm, year, week_key, first_line):
    out, parent = [], None
    nextmon = datetime.date.fromisoformat(rules.week_of(week_key or None)["mon"]) + datetime.timedelta(days=7)
    get = lambda row, k: row[hm[k]] if k in hm and hm[k] < len(row) else None
    for n, row in enumerate(rows):
        line = first_line + n
        if not any(get(row, k) not in (None, "") for k in ("text", "kind", "start", "end", "period")):
            continue  # 빈 줄(부서·작성자·병합된 과제명만 있는 줄 포함)
        note = str(get(row, "note") or "")
        text = str(get(row, "text") or "").strip()
        if note.startswith(EXAMPLE) or text.startswith("(예시)"):
            continue
        errors, warns = [], []
        if not text:
            errors.append("내용 없음")
        raw_s, raw_e = get(row, "start"), get(row, "end")
        st, en = to_date(raw_s, year), to_date(raw_e, year)
        if raw_s not in (None, "") and not st:
            errors.append(f"시작일 '{raw_s}' 을 날짜로 못 읽음")
        if raw_e not in (None, "") and not en:
            errors.append(f"종료일 '{raw_e}' 을 날짜로 못 읽음")
        per = period_of(st, en)
        if not per and get(row, "period"):
            per = rules.norm_period(str(get(row, "period")), year)
        kind = kind_of(get(row, "kind"))
        if not kind:
            d0 = st or en
            if get(row, "kind") not in (None, ""):
                errors.append(f"구분 '{get(row, 'kind')}' 을 모름(수행/계획)")
            elif d0:
                kind = "plan" if d0 >= nextmon else "done"
                warns.append("구분이 비어 날짜로 추정: " + ("계획" if kind == "plan" else "수행"))
            else:
                errors.append("구분 없음(수행/계획)")
        try:
            depth = max(0, min(2, int(float(str(get(row, "depth") or 1).strip() or 1)) - 1))
        except ValueError:
            depth = 1 if re.search(r"하위", str(get(row, "depth"))) else 0
        project = str(get(row, "project") or "").strip().strip("()[]")
        if depth > 0 and parent and not project:
            project, kind = parent["item"]["project"], kind or parent["item"]["kind"]
        if not project and depth == 0:
            warns.append("과제명 없음 — 반영 뒤 과제명 질의")
        if not per and depth == 0:
            warns.append("기간 없음")
        it = {"text": text, "kind": kind or "done", "project": project, "period": per, "depth": depth, "cat": guess_cat(text),
              "core": yes(get(row, "core")), "msit": yes(get(row, "msit")), "nobbs": yes(get(row, "nobbs")),
              "place": str(get(row, "place") or "").strip(), "people": str(get(row, "people") or "").strip(),
              "party": str(get(row, "party") or "").strip()}
        rec = {"row": line, "dept": str(get(row, "dept") or "").strip(), "name": str(get(row, "name") or "").strip(), "item": it, "errors": errors, "warnings": warns}
        if depth == 0:
            parent = rec
        out.append(rec)
    return out


def parse_xlsx_report(sheets, year):
    """보고서형 엑셀(이 도구가 낸 것 또는 손으로 같은 모양으로 쓴 것) — 'N. 수행업무' 머리 줄 아래 A=실, B=여러 줄 내용"""
    groups, abbrs, outer, found = {}, [], [], False
    for sh in sheets:
        if sh["hidden"]:
            continue
        kind = ""
        for r, row in enumerate(sh["rows"]):
            a = str(row[0] if row else "") if row and row[0] is not None else ""
            b = str(row[1]) if len(row) > 1 and row[1] is not None else ""
            k = _kind_of_header(a) if re.match(r"^\s*(\d\.\s*)?(수행|향후|금주|차주)", a) and len(a) < 30 else ""
            if k:
                kind, found = k, True
                continue
            if kind and a and b and a != b:
                rich = sh["rich"].get((r, 1)) or [{"t": b}]
                lines = _rich_lines(rich)
                items, ab = parse_lines(lines, kind, year)
                abbrs += ab
                lab = [x for x in a.split("\n") if x.strip()]
                label = re.sub(r"\s+", "", "".join(x for x in lab if not x.strip().startswith("(")))
                who = next((x.strip("() ") for x in lab if x.strip().startswith("(")), "")
                g = groups.setdefault((label, who), {"dept": label, "name": who, "items": []})
                g["items"] += [x for x in items if x["text"]]
            else:
                if a and not b:
                    kind = "" if a.strip().startswith(("3.", "※")) else kind
                    rich = sh["rich"].get((r, 0))
                    t = "".join(x["t"] for x in rich) if rich else a
                    outer.append(t)
    abbrs += _abbr_block(outer)
    return found, list(groups.values()), abbrs, _remarks_from(outer), _org_from(outer), "\n".join(outer)


def _rich_lines(rich):
    """서식 있는 글(조각) → 줄별 [{"t", "left", "flags", "bold"}] — 줄 앞 공백 수를 들여쓰기로"""
    lines, cur = [], []
    for x in rich:
        parts = x["t"].split("\n")
        for i, part in enumerate(parts):
            if i:
                lines.append(cur)
                cur = []
            if part:
                cur.append(dict(x, t=part))
    lines.append(cur)
    out = []
    for ln in lines:
        t = "".join(x["t"] for x in ln)
        if not t.strip():
            continue
        flags = {}
        body = [x for x in ln if x["t"].strip() and x["t"].strip() not in ("-", "·", "∙")]
        for x in body:
            f = color_flag(x.get("color"))
            if f:
                flags[f] = True
                if f == "msit" and x.get("bold"):
                    flags["core"] = True
            if x.get("strike"):
                flags["nobbs"] = True
        out.append({"t": t.strip(), "left": len(t) - len(t.lstrip(" ")), "flags": flags, "bold": bool(body) and all(x.get("bold") for x in body),
                    "abbr_bold": bool(body) and bool(body[0].get("bold"))})
    return out


# ── 그 밖(HWP·PDF·TXT·MD): kordoc 으로 글을 꺼내 표 모양이면 해석, 아니면 메모 ─────────────
def kordoc_text(name, data, cli):
    if not cli or not os.path.exists(cli):
        return None
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "in" + os.path.splitext(name)[1].lower())
        with open(p, "wb") as f:
            f.write(data)
        try:
            r = subprocess.run(["node", cli, p, "--silent"], capture_output=True, text=True, timeout=120)
        except (OSError, subprocess.TimeoutExpired):
            return None
        return r.stdout if r.returncode == 0 else None


def parse_markdown(md, year):
    """kordoc 마크다운(표는 HTML <table>, 칸 안 줄은 <br>) 또는 손으로 쓴 글 — 표를 찾으면 칸 해석"""
    groups, found = {}, False
    for tb in re.findall(r"<table>.*?</table>", md, re.S):
        trs = re.findall(r"<tr>(.*?)</tr>", tb, re.S)
        kind = ""
        for tr in trs:
            cells = [re.sub(r"<[^>]+>", "", html.unescape(c)).strip() for c in re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr.replace("<br>", "\n"), re.S)]
            k = _kind_of_header(" ".join(cells)) if len(cells) <= 2 and re.search(r"\d\.\s*(수행|향후)", " ".join(cells)) else ""
            if k:
                kind, found = k, True
                continue
            if kind and len(cells) >= 2 and cells[0] and cells[0] not in ("부서", "실"):
                lines = [{"t": x.strip(), "left": len(x) - len(x.lstrip()), "flags": {}, "bold": False} for x in cells[-1].split("\n")]
                items, _ = parse_lines(lines, kind, year)
                label = re.sub(r"\s+", "", cells[0])
                g = groups.setdefault(label, {"dept": label, "name": "", "items": []})
                g["items"] += [x for x in items if x["text"]]
    plain = re.sub(r"<table>.*?</table>", "", md, flags=re.S)
    paras = [re.sub(r"<[^>]+>", "", x) for x in plain.splitlines()]
    return found, list(groups.values()), _abbr_block(paras), _remarks_from(paras), _org_from(paras), plain


# ── 한 파일 가져오기 ────────────────────────────────────────────────────────
def import_file(name, data, week_key, kordoc_cli=None):
    """→ {"name", "format", "method": embedded|table|form|memo, "rows": [{row, dept, name, item, errors, warnings}], "abbrs", "remarks", "org", "memo", "notes"}"""
    year = _years_of(week_key)
    ext = os.path.splitext(name)[1].lower().lstrip(".")
    res = {"name": name, "format": ext, "method": "", "rows": [], "abbrs": [], "remarks": [], "org": "", "memo": "", "notes": []}
    emb = find_embedded(name, data) if ext in ("hwpx", "docx", "xlsx", "pptx") else None
    if emb and emb.get("kind") == "template":
        res["notes"].append("이 도구의 작성 양식 — 채운 내용을 양식 모양대로 읽음")
        emb = None
    if emb:
        res.update(method="embedded", org=emb.get("org", ""), remarks=emb.get("remarks") or [])
        for g in emb.get("groups") or []:
            for it in g.get("items") or []:
                res["rows"].append({"row": None, "dept": g.get("dept", ""), "name": g.get("name", ""), "item": it, "errors": [], "warnings": []})
        res["notes"].append("이 도구가 내보낸 파일 — 안에 넣어 둔 항목을 그대로 가져옴(손실 없음)")
        return res
    found, groups, abbrs, remarks, org, text = False, [], [], [], "", ""
    try:
        if ext == "xlsx":
            sheets = xlsx.read(data)
            ok, rows, rem = parse_xlsx_table(sheets, year, week_key)
            if ok:
                res.update(method="table", rows=rows, remarks=rem)
                res["notes"].append("엑셀 입력 양식(머리글)으로 읽음")
                return res
            found, groups, abbrs, remarks, org, text = parse_xlsx_report(sheets, year)
            if not found:
                text = "\n".join(" ".join(str(x) for x in r if x not in (None, "")) for sh in sheets if not sh["hidden"] for r in sh["rows"])
        elif ext == "hwpx":
            found, groups, abbrs, remarks, org, text = parse_hwpx(data, year)
        elif ext == "docx":
            found, groups, abbrs, remarks, org, text = parse_docx(data, year)
        elif ext == "pptx":
            found, groups, abbrs, remarks, org, text = parse_pptx(data, year)
        elif ext in ("txt", "md"):
            md = data.decode("utf-8", errors="replace")
            found, groups, abbrs, remarks, org, text = parse_markdown(md, year)
        elif ext in ("hwp", "pdf", "hml", "doc"):
            md = kordoc_text(name, data, kordoc_cli)
            if md is None:
                res["notes"].append("이 형식은 kordoc(문서 변환기)이 있어야 읽을 수 있습니다")
                return res
            found, groups, abbrs, remarks, org, text = parse_markdown(md, year)
            res["notes"].append("kordoc 으로 글을 꺼냄(글자색·취소선 정보는 없음)")
        else:
            res["notes"].append(f"모르는 형식: .{ext}")
            return res
    except (zipfile.BadZipFile, KeyError, ValueError) as e:
        res["notes"].append(f"파일을 읽지 못함: {e}")
        return res
    res.update(abbrs=[{"abbr": a, "body": b} for a, b in abbrs], remarks=remarks, org=org)
    if found:
        res["method"] = "form"
        res["notes"].append("공식 양식 표(수행업무·향후 2주 계획)를 찾아 칸 안 글머리로 읽음"
                            + ("" if any(g["items"] for g in groups) else " — 내용이 적힌 줄은 없음(빈 양식)"))
        res["groups_found"] = [{"dept": g["dept"], "projects": sorted({i.get("project") for i in g["items"]} - {""})} for g in groups]
        for g in groups:
            for it in g["items"]:
                w = [] if it.get("project") or int(it.get("depth") or 0) else ["과제명 없음 — 반영 뒤 과제명 질의"]
                if not it.get("period") and not int(it.get("depth") or 0):
                    w.append("기간 없음")
                res["rows"].append({"row": None, "dept": g["dept"], "name": g["name"], "item": it, "errors": [], "warnings": w})
        return res
    res.update(method="memo", memo=text.strip()[:20000])
    res["notes"].append("양식 구조를 찾지 못해 글만 꺼냄 — 'LLM 으로 정리'를 누르면 항목으로 정리합니다(누르기 전에는 LLM 을 쓰지 않음)")
    return res


def decode_files(files):
    """[{"name", "b64"}] → [(name, bytes)]"""
    out = []
    for f in files or []:
        b = f.get("b64") or ""
        if "," in b[:80] and b.startswith("data:"):
            b = b.split(",", 1)[1]
        out.append((os.path.basename(str(f.get("name") or "file")), base64.b64decode(b)))
    return out


# ── PPTX: 보고서 내보내기 · 작성 양식 · 읽기 ─────────────────────────────────────
import pptxw as pptx  # noqa: E402  (python-pptx 와 이름이 겹치지 않게 pptxw)

PLACE_PROJ, PLACE_TEXT = "과제명", "(기간) 내용"


def _pptx_lines(paras, col):
    """render.layout 문단 → 글상자 문단(과제명 = 수준 0 '∙' 굵게, 항목 = 수준 1 '-', 하위 = 수준 2·3 '·', 약어 줄 = 글머리 없음)"""
    out = []
    for p in paras:
        runs = [{"t": r["t"], "bold": bool(r.get("bold")), "strike": bool(r.get("strike")), "color": col.get(r.get("color")) if r.get("color") else None}
                for r in p["runs"]]
        if p["kind"] == "proj":
            out.append({"lvl": 0, "bullet": "∙", "runs": runs[1:]})
        elif p["kind"] == "item":
            d = int(p.get("depth") or 0)
            out.append({"lvl": d + 1, "bullet": "-" if d == 0 else "·", "runs": runs[1:]})
        elif p["kind"] == "abbr":
            out.append({"lvl": int(p.get("depth") or 0) + 2, "runs": [dict(r, size=11) for r in runs]})
    return out


def _deck_slides(lay, tpl, guide=None):
    col = tpl["colors"]
    E = 914400
    slides = []
    head = [pptx._shape(2, "제목", int(0.5 * E), int(1.6 * E), int(12.3 * E), int(1.0 * E), [{"align": "ctr", "runs": [{"t": lay["title"], "bold": True, "size": 32}]}]),
            pptx._shape(3, "소본부", int(0.5 * E), int(2.8 * E), int(12.3 * E), int(0.8 * E), [{"align": "ctr", "runs": [{"t": (tpl.get("org_heading") or "Ⅰ. {org}").format(org=lay["org"]), "bold": True, "size": 22}]}]),
            pptx._shape(4, "기간", int(0.5 * E), int(3.8 * E), int(12.3 * E), int(1.2 * E),
                        [{"align": "ctr", "runs": [{"t": f"{t['title'].strip()}  {t['period']}", "size": 16}]} for t in lay["tables"].values()])]
    slides.append(head)
    t1, t2 = lay["tables"]["done"], lay["tables"]["plan"]
    for n, (r1, r2) in enumerate(zip(t1["rows"], t2["rows"])):
        lab = " ".join(render.para_text(p) for p in r1["label"])
        sh = [pptx._shape(2, "실", int(0.4 * E), int(0.2 * E), int(12.5 * E), int(0.6 * E), [{"runs": [{"t": lab, "bold": True, "size": 22}]}])]
        y = 0.9
        for i, (t, row) in enumerate(((t1, r1), (t2, r2))):
            paras = [{"runs": [{"t": t["title"].strip() + "    ", "bold": True, "size": 16}, {"t": t["period"], "size": 12}]}] + (_pptx_lines(row["cell"], col) or [{"lvl": 1, "bullet": "-", "runs": [{"t": "-"}]}])
            h = 3.0 if i == 0 else 2.6
            sh.append(pptx._shape(3 + i, t["title"].strip(), int(0.4 * E), int(y * E), int(12.5 * E), int(h * E), paras, size=14, line="7F7F7F"))
            y += h + 0.15
        if guide:
            sh.append(pptx._shape(9, "작성법", int(0.4 * E), int(6.95 * E), int(12.5 * E), int(0.45 * E), [{"runs": [{"t": g, "size": 10, "color": "#7F7F7F"}]} for g in guide], size=10))
        slides.append(sh)
    if lay["extra"]:  # ※ 약어 — 계획 슬라이드들 다음, 특기 슬라이드 앞(따로 한 장)
        ab = [{"runs": [{"t": y["t"], "bold": bool(y.get("bold")), "size": 11} for y in x]} for x in lay["extra"]]
        slides.append([pptx._shape(2, "약어", int(0.4 * E), int(0.3 * E), int(12.5 * E), int(6.9 * E), ab, size=14)])
    if tpl.get("remarks", True):
        tail = [{"runs": [{"t": tpl.get("remarks_title", "3. 특기 및 애로사항") + " ", "bold": True, "size": 16}, {"t": "(필요 시)", "size": 11}]}]
        tail += [{"lvl": 1, "bullet": "-", "runs": [{"t": x}]} for x in lay["remarks"]] or [{"lvl": 1, "bullet": "-", "runs": [{"t": ""}]}]
        slides.append([pptx._shape(2, "특기", int(0.4 * E), int(0.3 * E), int(12.5 * E), int(6.9 * E), tail, size=14)])
    return slides


def report_pptx(doc, gl, tpl, bbs=False, embed=None):
    lay = render.layout(doc, gl, tpl, bbs)
    return pptx.deck(_deck_slides(lay, tpl), json.dumps(embed, ensure_ascii=False) if embed else None, lay["title"])


# ── 빈 작성 양식 (XLSX 는 input_template) — HWPX·DOCX·PPTX: 실마다 빈 행, '∙ (과제명)' / '- (기간) 내용' 자리 ─────────
TEMPLATE_GUIDE = ["작성법: '∙ (과제명)' 한 줄 아래 '- (기간) 내용'(기간: 2.24~3.5 / ~3.6 / 3.20). 하위 항목은 한 단계 안으로(·).",
                  "핵심 = 주황 글자, 과기정통부 보고 = 파랑 글자, BBS 비게시 = 취소선. 외부활동은 내용 뒤 (@장소, 참석자 — 우리 연구원 사람만), 상대 기관·인물은 문장 안에(예: 과기정통부 김사무관과 진도 점검 회의). 약어는 그대로 — 풀이는 자동."]


def registry_doc(units, week_key, org="", per_person=False):
    """과제 등록부로 채운 작성 양식 문서: units = [{"dept", "members", "projects": [{"name", "period", "person"}]}]
    실(또는 사람)마다 두 표 모두 활성 과제마다 '∙ (과제명)' + 빈 '- (기간) 내용' 한 줄. 과제가 없으면 '(과제명)' 빈 자리."""
    wk = rules.week_of(week_key or None)
    rows = []
    for u in units:
        people = (u.get("members") or [""]) if per_person else [""]
        for who in people:
            ps = [p for p in u.get("projects") or [] if not who or not p.get("person") or p["person"] == who]
            its = []
            for kind in rules.KINDS:
                for n, p in enumerate(ps or [{"name": PLACE_PROJ}]):
                    its.append({"id": f"ph{kind}{n}", "text": PLACE_TEXT if not p.get("period") else "내용", "kind": kind, "project": p["name"],
                                "period": p.get("period", "") if kind == "done" else "", "depth": 0, "cat": "goal"})
            rows.append({"label": u["dept"], "sub": who, "items": its})
    return {"title": "주간업무보고", "week": wk["key"], "org": org, "remarks": [], "keep_guide": True, "guide": TEMPLATE_GUIDE,
            "periods": {"done": wk["label"], "plan": rules.plan_period(wk["key"])}, "expand_mode": "off", "legend": False,
            "rows": rows or [{"label": "○○실", "items": []}], "project_order": {u["dept"]: [p["name"] for p in u.get("projects") or []] for u in units}}


def template_from(doc, fmt, gl, tpl, xlsx_rows=None, projects=None):
    """작성 양식 파일 — fmt: xlsx | hwpx | docx | pptx"""
    mark = payload([], doc.get("week"), doc.get("org", ""), [], "template")
    if fmt == "xlsx":
        one = doc["rows"][0] if len(doc["rows"]) == 1 else {}
        return input_template(projects, one.get("label", ""), one.get("sub", ""), doc.get("week"), prefill=xlsx_rows)
    if fmt == "hwpx":
        return embed_zip(render.to_hwpx(doc, gl, tpl), "hwpx", mark)
    if fmt == "docx":
        return embed_zip(render.to_docx(doc, gl, tpl), "docx", mark)
    if fmt == "pptx":
        lay = render.layout(doc, gl, tpl)
        return pptx.deck(_deck_slides(lay, tpl, TEMPLATE_GUIDE), json.dumps(mark, ensure_ascii=False), lay["title"])
    raise ValueError("형식은 xlsx|hwpx|docx|pptx")


def template_doc(depts, week_key, org="", name=""):
    wk = rules.week_of(week_key or None)
    ph = lambda kind: [{"id": f"ph{kind}", "text": PLACE_TEXT, "kind": kind, "project": PLACE_PROJ, "period": "", "depth": 0, "cat": "goal"}]
    return {"title": "주간업무보고", "week": wk["key"], "org": org or (depts[0] if depts else ""), "remarks": [], "keep_guide": True, "guide": TEMPLATE_GUIDE,
            "periods": {"done": wk["label"], "plan": rules.plan_period(wk["key"])}, "expand_mode": "off", "legend": False,
            "rows": [{"label": d, "sub": name, "items": ph("done") + ph("plan")} for d in depts or ["○○실"]]}


def form_template(fmt, depts, week_key, gl, tpl, org="", name="", projects=None):
    """빈 작성 양식 내려받기 — 넣어 둔 표시(kind=template)로 이 도구 양식임을 알아보되, 올리면 사람이 채운 내용을 양식 모양 그대로 읽는다"""
    if fmt == "xlsx":
        return input_template(projects, depts[0] if len(depts) == 1 else "", name, week_key)
    doc = template_doc(depts, week_key, org, name)
    mark = payload([], week_key, org, [], "template")
    if fmt == "hwpx":
        return embed_zip(render.to_hwpx(doc, gl, tpl), "hwpx", mark)
    if fmt == "docx":
        return embed_zip(render.to_docx(doc, gl, tpl), "docx", mark)
    if fmt == "pptx":
        lay = render.layout(doc, gl, tpl)
        return pptx.deck(_deck_slides(lay, tpl, TEMPLATE_GUIDE), json.dumps(mark, ensure_ascii=False), lay["title"])
    raise ValueError("형식은 xlsx|hwpx|docx|pptx")


def parse_pptx(data, year):
    """슬라이드 → 실(맨 위 짧은 글상자) · '1. 수행업무' / '2. 향후 2주 계획' 글상자의 목록 수준(0=과제명, 1=항목, 2·3=하위) · 글자색·취소선"""
    groups, abbrs, outer, found, org = {}, [], [], False, ""
    for sl in pptx.read(data):
        label = ""
        for sh in sl["shapes"]:
            ps = [p for p in sh["paras"] if p["t"].strip()]
            if not ps:
                continue
            hdr = lambda t: bool(re.match(r"^\s*(\d\.\s*)?(수행|향후|금주|차주)", t)) and len(t) < 40
            k = _kind_of_header(ps[0]["t"]) if hdr(ps[0]["t"]) and not (len(ps) > 1 and hdr(ps[1]["t"])) else ""  # 표지의 기간 요약 상자는 제외
            if not k:
                txt = [p["t"] for p in ps]
                if not label and len(ps) == 1 and len(ps[0]["t"]) < 30 and not re.match(r"^\s*[ⅠⅡⅢ]", ps[0]["t"]) and "주간업무" not in ps[0]["t"] and not ps[0]["t"].startswith("작성법"):
                    label = re.sub(r"\s+", "", re.sub(r"\(.*?\)", "", ps[0]["t"]))
                    who = (re.search(r"\(([^)]*)\)", ps[0]["t"]) or [None, ""])[1]
                else:
                    outer += txt
                continue
            found = True
            lines = []
            for p in ps[1:]:
                t = p["t"].strip()
                body = [r for r in p["runs"] if r["t"].strip()]
                flags = {}
                for r in body:
                    f = color_flag(r.get("color"))
                    if f:
                        flags[f] = True
                        if f == "msit" and r.get("bold"):
                            flags["core"] = True
                    if r.get("strike"):
                        flags["nobbs"] = True
                if p["lvl"] == 0 and not ABBR_LINE.match(t):
                    t = "∙ (" + t.strip("∙ ").strip("()[]") + ")"
                lines.append({"t": t, "left": p["lvl"], "flags": flags, "bold": bool(body) and all(r.get("bold") for r in body),
                              "abbr_bold": bool(body) and bool(body[0].get("bold")), "small": False})
            items, ab = parse_lines(lines, k, year)
            abbrs += ab
            g = groups.setdefault((label, who if label else ""), {"dept": label, "name": who if label else "", "items": []})
            g["items"] += [x for x in items if x["text"]]
    org = _org_from(outer)
    abbrs += _abbr_block(outer)
    return found, [g for g in groups.values() if g["items"] or g["dept"]], abbrs, _remarks_from(outer), org, "\n".join(outer)

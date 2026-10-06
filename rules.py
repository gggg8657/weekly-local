"""주간보고 규칙 — LLM 없이 결정론으로: 약어 검출·풀이, 외부활동 (@장소, 참석자) 누락, 분량(1쪽) 추정, 빈 칸, 날짜 표기 통일,
LLM 결과의 원문 대조(숫자·영문 표기·장소·참석자). app.py 와 render.py 가 같이 쓴다."""
import datetime
import json
import os
import re
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
CATS = {"goal": "중점목표 연계 실적·계획", "perf": "주요 연구(경영) 성과 및 대외활동", "event": "주요 행사(연구원 주관)·수상", "etc": "기타 업무"}
CAT_ORDER = list(CATS)
KINDS = {"done": "수행한 일", "plan": "향후 2주 계획"}
EXT_WORDS = ("참석", "발표", "출장", "회의", "방문", "워크숍", "워크샵", "학회", "국제", "세미나", "포럼", "심포지엄", "컨퍼런스", "콘퍼런스",
             "면담", "실사", "설명회", "간담회", "자문", "초청", "강연", "전시", "박람회")
INTERNAL = re.compile(r"내부|회의실|실\s?회의|팀\s?회의|부서\s?회의|주간\s?회의|실 내|원내|과제\s?회의|정기\s?회의|수상|"
                      r"(?:실장|부장|팀장|단장|소장|원장|본부장)\s?면담|채용|면접")
WEAK = ("회의",)  # LLM(또는 사용자)이 외부활동 아님(ext=false)으로 둔 항목에서는 이 낱말만으로 경고하지 않는다
# 풀이가 필요 없는 단위·파일 형식 (설정 탭에서 사용자 무시 목록을 더할 수 있음)
IGNORE = {"MW", "MWe", "MWt", "MWh", "kW", "kWe", "kWh", "GW", "GWe", "GB", "TB", "MB", "KB", "PDF", "HWP", "HWPX", "DOCX", "XLSX",
          "PPT", "PPTX", "CSV", "PhD", "Dr", "AM", "PM", "OK", "ID", "PC", "USB", "Q&A", "TF", "mSv", "uSv", "mGy", "GBq", "TBq", "MeV",
          "keV", "eV", "GHz", "MHz", "kHz", "Hz", "pH", "CO2"}
WD = "월화수목금토일"
LOCK = threading.Lock()


def load_json(p, default):
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, ValueError):
        return default


def save_json(p, obj):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with LOCK:
        with open(p + ".tmp", "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=1)
        os.replace(p + ".tmp", p)


# ── 주차 ────────────────────────────────────────────────────────────────
def week_of(d=None):
    """날짜 → {"key": "2026-W41", "mon", "fri", "label": "2026. 10. 5.(월) ~ 10. 9.(금)"}"""
    if isinstance(d, str):
        m = re.fullmatch(r"(\d{4})-W(\d{2})", d)
        d = datetime.date.fromisocalendar(int(m.group(1)), int(m.group(2)), 1) if m else datetime.date.fromisoformat(d)
    d = d or datetime.date.today()
    y, w, _ = d.isocalendar()
    mon = d - datetime.timedelta(days=d.weekday())
    fri = mon + datetime.timedelta(days=4)
    return {"key": f"{y}-W{w:02d}", "mon": mon.isoformat(), "fri": fri.isoformat(),
            "label": f"{mon.year}. {mon.month}. {mon.day}.(월) ~ {fri.month}. {fri.day}.(금)",
            "title": f"{mon.year}년 {mon.month}월 {(mon.day - 1) // 7 + 1}주차"}


# ── 약어집 ──────────────────────────────────────────────────────────────
FIELDS = ["원자력", "기관·정책", "AI", "컴퓨터", "일반"]
# 풀이 수준: 1 전체명만 / 2 전체명 + 한글 / 3 전체명 + 한글 + 쉬운 설명. 분야별로 따로.
LEVELS = {"원자력": 1, "기관·정책": 2, "AI": 3, "컴퓨터": 2, "일반": 2}
PRESETS = {"원자력 전문가": LEVELS,
           "과기정통부·경영진": {"원자력": 2, "기관·정책": 1, "AI": 3, "컴퓨터": 3, "일반": 2},
           "일반 공개": {"원자력": 3, "기관·정책": 2, "AI": 3, "컴퓨터": 3, "일반": 3}}
UNKNOWN = "[약어 확인 필요]"  # 미확인 약어 뒤에 붙는 표시(미리보기 강조, 내보낸 문서에도 그대로 — 모르고 나가지 않게)
SEED_SRC = {"kr_seed.json": "사내 시드(확인 필요)", "tech_seed.json": "큐레이션(weekly-local)"}
PUBLIC = (("nureg0544.json", "원자력"), ("nist_csrc.json", "컴퓨터"))


class Glossary:
    """우선순위: 사내(WORKSPACE/glossary.json — 작성자 답변·수정, 이력) > 시드(국내 기관·원자력, AI·컴퓨터 큐레이션) > 공개(NUREG-0544, NIST CSRC).
    - 사내 항목은 약어 하나에 풀이 하나(이력 보관).
    - 시드는 동음이의(DT, ROM, PIE …)면 항목마다 ctx(문맥 낱말)가 있어, 문장에 그 낱말이 있어야 확정한다. 확정 못 하면 작성자에게 묻는다.
    - 공개 약어집은 옛 뜻·다른 분야 뜻이 섞여(GPU → General Public Utilities 등) 자동으로 쓰지 않고, 작성자 질의 때 후보로만 보여 준다."""

    def __init__(self, ws):
        self.user_path = os.path.join(ws, "glossary.json")
        self.ignore_path = os.path.join(ws, "glossary_ignore.json")
        self.seed = []
        for fn, src in SEED_SRC.items():
            for r in load_json(os.path.join(HERE, "glossary", fn), []):
                self.seed.append({"ko": "", "desc": "", "field": "일반", **r, "src": r.get("src") or src})
        self.seed_idx = {}
        for r in self.seed:
            self.seed_idx.setdefault(r["abbr"], []).append(r)
        self.public, self.pub_count = [], {}
        for fn, field in PUBLIC:
            rows = load_json(os.path.join(HERE, "glossary", fn), [])
            self.pub_count[rows[0]["src"] if rows else fn] = len(rows)
            self.public += [{"field": field, **r} for r in rows]
        self.pub_idx, self.pub_ci = {}, {}
        for r in self.public:
            self.pub_idx.setdefault(r["abbr"], []).append(r)
        for k in self.pub_idx:
            self.pub_ci.setdefault(k.lower(), []).append(k)

    def user(self):
        return load_json(self.user_path, {})

    def ignored(self):
        return set(load_json(self.ignore_path, [])) | IGNORE

    def public_cands(self, abbr):
        rows = self.pub_idx.get(abbr) or [r for k in self.pub_ci.get(abbr.lower(), []) for r in self.pub_idx[k]]
        seen, out = set(), []
        for r in rows:
            if r["full"].lower() not in seen:
                seen.add(r["full"].lower())
                out.append(r)
        return out

    def resolve(self, abbr, ctx=""):
        """→ (확정 항목 or None, 후보 목록, 상태). 상태: user | seed | homonym(동음이의, 문맥으로 못 정함) | public(공개 약어집에만) | missing"""
        u = self.user().get(abbr)
        if u and not u.get("deleted"):
            return dict(u, abbr=abbr, src=u.get("src") or "사내"), [], "user"
        seeds = self.seed_idx.get(abbr) or []
        if seeds:
            if len(seeds) == 1 and not seeds[0].get("ctx"):
                return seeds[0], [], "seed"
            scored = sorted(((sum(w.lower() in (ctx or "").lower() for w in s.get("ctx") or []), i) for i, s in enumerate(seeds)), reverse=True)
            if scored[0][0] > 0 and (len(scored) == 1 or scored[0][0] > scored[1][0]):
                return seeds[scored[0][1]], [], "seed"
            return None, seeds + self.public_cands(abbr), "homonym"
        cands = self.public_cands(abbr)
        return None, cands, "public" if cands else "missing"

    def lookup(self, abbr, ctx=""):
        e, c, _ = self.resolve(abbr, ctx)
        return e, c

    def put(self, abbr, full, ko="", note="", editor="", src="사내", desc="", field="", ctx=""):
        abbr, full = abbr.strip(), full.strip()
        if not abbr or not (full or ko):
            raise ValueError("약어와 풀이를 적어 주세요")
        u = self.user()
        old = u.get(abbr) or {}
        hist = old.get("history") or []
        keys = ("full", "ko", "desc", "field", "note", "editor", "ts", "src")
        if old and not old.get("deleted"):
            hist.append({k: old.get(k, "") for k in keys})
        field = field if field in FIELDS else (old.get("field") or next((s["field"] for s in self.seed_idx.get(abbr, [])), "일반"))
        u[abbr] = {"full": full, "ko": ko.strip(), "desc": desc.strip(), "field": field, "note": note.strip(), "editor": editor.strip(), "src": src,
                   "ctx": ctx.strip()[:200], "ts": datetime.datetime.now().isoformat(timespec="seconds"), "history": hist[-20:]}
        save_json(self.user_path, u)
        return u[abbr]

    def delete(self, abbr, editor=""):
        u = self.user()
        if abbr in u:
            hist = (u[abbr].get("history") or []) + [{k: u[abbr].get(k, "") for k in ("full", "ko", "desc", "field", "note", "editor", "ts", "src")}]
            u[abbr] = {"deleted": True, "editor": editor, "ts": datetime.datetime.now().isoformat(timespec="seconds"), "history": hist[-20:]}
            save_json(self.user_path, u)

    def ignore(self, abbr, on=True):
        cur = set(load_json(self.ignore_path, []))
        (cur.add if on else cur.discard)(abbr)
        save_json(self.ignore_path, sorted(cur))

    def search(self, q, limit=200, field=""):
        q = (q or "").strip().lower()
        hit = lambda r: (not q or q in r["abbr"].lower() or q in ((r.get("full") or "") + (r.get("ko") or "") + (r.get("desc") or "")).lower()) \
            and (not field or r.get("field") == field)
        out, seen = [], set()
        for abbr, u in sorted(self.user().items()):
            r = {"abbr": abbr, **{k: u.get(k, "") for k in ("full", "ko", "desc", "field", "note", "editor", "ts", "src")}}
            if not u.get("deleted") and hit(r):
                out.append(dict(r, tier="사내", history=u.get("history") or []))
                seen.add(abbr)
        for s in self.seed:
            if s["abbr"] not in seen and hit(s):
                out.append(dict(s, tier="시드", homonym=len(self.seed_idx[s["abbr"]]) > 1 or bool(s.get("ctx"))))
        seen |= set(self.seed_idx)
        if q:
            for r in self.public:
                if len(out) >= limit:
                    break
                if (q == r["abbr"].lower() or (len(q) >= 3 and (r["abbr"].lower().startswith(q) or q in r["full"].lower()))) and (not field or r["field"] == field):
                    out.append(dict(r, tier="공개", shadowed=r["abbr"] in seen))
        return out[:limit]

    def stats(self):
        u = {k for k, v in self.user().items() if not v.get("deleted")}
        by = {}
        for s in self.seed:
            by[s["src"]] = by.get(s["src"], 0) + 1
        return {"user": len(u), "seed": len(self.seed), "seed_by_src": by, "public": len(self.public), "public_by_src": self.pub_count,
                "public_abbr": len(self.pub_idx)}


# ── 약어 검출·풀이 ───────────────────────────────────────────────────────
ABBR_RE = re.compile(r"(?<![A-Za-z0-9&])((?:[a-z]{1,2}-)?[A-Za-z][A-Za-z0-9&]*(?:[-/][A-Za-z0-9&]+)*)(?![A-Za-z0-9&])")


def is_abbr(tok):
    up = sum(c.isupper() for c in tok)
    return up >= 2 or (up >= 1 and any(c.isdigit() for c in tok) and tok[0].isupper())


def find_abbrs(text):
    """[(약어, 시작, 끝)] — 'OECD/NEA' 처럼 사전에 통째로 있을 수 있는 조합은 그대로 한 토큰."""
    out = []
    for m in ABBR_RE.finditer(text or ""):
        tok = m.group(1).rstrip("-/")
        if is_abbr(tok):
            out.append((tok, m.start(1), m.start(1) + len(tok)))
    return out


def already_expanded(text, s, e, english_only=False):
    """'MARS-KS(Multi-…)' 또는 '원자력안전위원회(NSSC)' 처럼 이미 풀어 쓴 자리인지.
    english_only: 영문 전체 이름이 붙어 있을 때만 — '이상탐지(VAD)', 'LLM(대규모 언어모델)' 처럼 한글 풀이만 있으면 아니라고 본다
    (모든 약어는 영문 전체 이름을 따로 적는 규칙 — ※ 약어 목록·풀이 줄 방식)"""
    if english_only:
        inner = re.match(r"\s?\(([^)]{3,})\)", text[e:])
        if inner and len(re.findall(r"[A-Za-z]{2,}", inner.group(1))) >= 2:
            return True
        if s > 1 and text[s - 1] == "(" and text[e:e + 1] == ")":
            return len(re.findall(r"[A-Za-z]{2,}", re.split(r"[,(]", text[max(0, s - 80):s - 1])[-1])) >= 2  # 'Large Language Model(LLM)'
        return False
    after = text[e:e + 3]
    if re.match(r"\s?\(", after):
        inner = re.match(r"\s?\(([^)]{3,})\)", text[e:])
        if inner and (re.search(r"[가-힣]", inner.group(1)) or len(re.findall(r"[A-Za-z]+", inner.group(1))) >= 2):
            return True  # 'sLLM(gemma)' 처럼 한 단어 영문 괄호는 풀이로 보지 않음
    return s > 1 and text[s - 1] == "(" and text[e:e + 1] == ")" and re.search(r"[가-힣A-Za-z]", text[max(0, s - 6):s - 1]) is not None


def level_of(entry, levels=None):
    lv = {**LEVELS, **(levels or {})}
    try:
        return max(1, min(3, int(lv.get(entry.get("field") or "일반", 2))))
    except (TypeError, ValueError):
        return 2


def expansion(entry, level=2, desc_block=False):
    """→ (괄호 안 글, 따로 뺄 설명 or None). 수준 1 도 전체명은 반드시(영문 전체명이 없으면 한글 명칭)."""
    full, ko, desc = (entry.get("full") or "").strip(), (entry.get("ko") or "").strip(), (entry.get("desc") or "").strip()
    head = full or ko
    if level >= 2 and full and ko and ko != full:
        head = f"{full}, {ko}"
    if level >= 3 and desc:
        if desc_block:
            return head, desc
        return f"{head}: {desc}", None
    return head, None


def _units(t):
    """'OECD/NEA' 가 통째로 없으면 나눠 본다"""
    return [t] if "/" not in t else [t] + [x for x in t.split("/") if is_abbr(x)]


def line_body(entry, level=2, desc_block=False):
    """약어 풀이 줄(lines 방식)의 ':' 뒤 — 1 'Small Modular Reactor' / 2 'Small Modular Reactor(소형모듈원자로)' / 3 … + ' — 쉬운 설명'"""
    full, ko, desc = (entry.get("full") or "").strip(), (entry.get("ko") or "").strip(), (entry.get("desc") or "").strip()
    body = full or ko
    if level >= 2 and full and ko and ko != full:
        body = f"{full}({ko})"
    if level >= 3 and desc and not desc_block:
        body += f" — {desc}"
    return body


def expand_items(texts, gl, mode="lines", levels=None, desc_block=False, first_only=True, contexts=None):
    """문서 순서대로 받은 항목 글 목록 → (글 목록, 약어 목록[(약어, 풀이)], 용어 설명[(약어, 이름, 설명)], 미확인 약어 집합, 항목별 풀이 줄[[(약어, 풀이)]]).
    mode: note = 본문은 그대로, 문서 끝 '※ 약어' 아래 한 줄에 하나 '**약어**: 풀이' (공식 양식 기본)
          lines = 약어가 처음 나온 항목 바로 아래에 한 줄씩 / paren = 첫 등장에 괄호 병기 / off.
    미확인 약어(약어집에 없거나 동음이의를 못 정함)는 지어내지 않고 '[약어 확인 필요]' 로 둔다(note·lines 는 풀이 줄에, paren 은 약어 뒤에)."""
    seen, notes, descs, unknown, out, lines = set(), [], [], set(), [], []
    ign = gl.ignored()
    full_ctx = " ".join(contexts or texts)
    for n, t in enumerate(texts):
        lines.append([])
        if mode == "off" or not t:
            out.append(t)
            continue
        ctx = (contexts[n] if contexts else t) + " " + full_ctx
        parts, pos = [], 0
        for tok, s, e in find_abbrs(t):
            if tok in ign or (first_only and tok in seen):
                continue
            entry, _, st = gl.resolve(tok, ctx)
            if not entry and "/" in tok and any(gl.resolve(x, ctx)[0] for x in tok.split("/") if is_abbr(x)):
                continue  # 'OECD/NEA' 의 부분이 따로 풀리는 경우 — abbr_warnings 가 부분별로 본다
            first = tok not in seen
            seen.add(tok)
            if already_expanded(t, s, e, english_only=mode in ("note", "lines")):
                continue
            if not entry:
                if first:
                    unknown.add(tok)
                    if mode == "lines":
                        lines[-1].append((tok, UNKNOWN))
                    elif mode == "note":
                        notes.append((tok, UNKNOWN))
                    else:
                        parts += [t[pos:e], UNKNOWN]
                        pos = e
                continue
            head, d = expansion(entry, level_of(entry, levels), desc_block)
            if d and first and mode != "note":  # note 는 '약어: 이름 — 설명' 한 줄로 한 목록(용어 설명 목록을 따로 두지 않음)
                descs.append((tok, head, d))
            if mode == "lines":
                lines[-1].append((tok, line_body(entry, level_of(entry, levels), desc_block)))
            elif mode == "paren":
                parts += [t[pos:e], f"({head})"]
                pos = e
            elif first:
                notes.append((tok, line_body(entry, level_of(entry, levels), False)))  # '약어: Full(한글) — 설명' 한 줄, 목록 하나
        out.append("".join(parts) + t[pos:])
    return out, notes, descs, unknown, lines


def abbr_warnings(items, gl, levels=None):
    """문서 전체 약어 점검 → [{kind:'abbr', abbr, item, status, sentence, candidates}]
    status: ok(자동 풀이) / expanded(본문에 풀이 있음) / homonym·public·missing(작성자에게 질의 — level 'ask') / nodesc(수준 3 인데 설명 없음)"""
    ign, seen, out = gl.ignored(), set(), []
    all_text = " ".join(i.get("text") or "" for i in items)
    for it in items:
        t = (it.get("text") or "") + ext_suffix(it)  # 장소·참석자 칸의 약어(KINS 대전 등)도
        for tok, s, e in find_abbrs(t):
            if tok in ign or tok in seen:
                continue
            seen.add(tok)
            exp = already_expanded(t, s, e, english_only=True)
            entry, cands, st = gl.resolve(tok, t + " " + all_text)
            if not entry and "/" in tok:
                subs = [x for x in tok.split("/") if is_abbr(x) and x not in seen and x not in ign]
                if any(gl.resolve(x, t)[0] for x in subs) or not cands:
                    for sub in subs:
                        seen.add(sub)
                        e2, c2, st2 = gl.resolve(sub, t + " " + all_text)
                        out.append(_abbr_row(sub, it, t, e2, c2, st2, exp, levels))
                    continue
            out.append(_abbr_row(tok, it, t, entry, cands, st, exp, levels))
    return out


def _abbr_row(tok, it, sentence, entry, cands, st, expanded, levels):
    base = {"kind": "abbr", "abbr": tok, "item": it.get("id"), "sentence": sentence[:160],
            "candidates": [{k: c.get(k, "") for k in ("full", "ko", "desc", "field", "src")} for c in cands[:12]]}
    if expanded:
        return dict(base, level="info", status="expanded", msg=f"{tok}: 본문에 이미 풀이가 있음")
    if entry:
        lv = level_of(entry, levels)
        head, d = expansion(entry, lv)
        if lv >= 3 and not (entry.get("desc") or "").strip():
            return dict(base, level="warn", status="nodesc", field=entry.get("field"),
                        msg=f"{tok}: 수준 3 인데 약어집에 쉬운 설명이 없어 '전체명 + 한글'로만 씀 — 약어집에서 설명을 넣으세요")
        return dict(base, level="info", status="ok", field=entry.get("field"), msg=f"{tok} → {head} [{entry.get('field', '')}·수준 {lv}] ({entry.get('src', '')})")
    why = {"homonym": "뜻이 여럿이라 문맥으로 정하지 못함", "public": f"공개 약어집에만 있음(후보 {len(cands)}개) — 맞는 뜻을 확인해야 함",
           "missing": "약어집에 없음"}[st]
    return dict(base, level="ask", status=st, msg=f"{tok}: {why} — 작성자 확인 필요")


def unresolved(items, gl):
    return [w for w in abbr_warnings(items, gl) if w["level"] == "ask"] + [w for w in project_warnings(items) if w["level"] == "ask"]


# ── 외부활동·분량·빈칸·날짜 ──────────────────────────────────────────────
INTERNAL_PLACE = re.compile(r"본관|회의실|연구동|본원|실험실|사무실|원내|내부|연구실|\bPC\b|자리|테스트\s?공간|온라인 내부")


def external_activity(it):
    """(@장소, 참석자) 를 붙일 외부활동인지: LLM·작성자가 외부활동으로 표시했거나 장소가 연구원 밖. 참석자(사람 이름)만 있으면 아니다 — 내부 동료."""
    place = (it.get("place") or "").strip()
    if place and INTERNAL_PLACE.search(place):
        return False
    if not place and re.search(r"전화|통화|메일|메신저", it.get("text") or ""):
        return False  # 장소 없는 전화·메일 협의는 외부활동 표기를 붙이지 않는다
    return bool(it.get("ext")) or bool(place)


def is_external(it):
    """LLM 이 외부활동으로 표시했거나, 외부활동 낱말이 있고 내부 회의 표시가 없을 때"""
    t = it.get("text") or ""
    if INTERNAL.search(t):
        return False
    if external_activity(it):
        return True
    words = [w for w in EXT_WORDS if w in t]
    return bool(words) and not (it.get("ext") is False and all(w in WEAK for w in words))


ATTENDEE_MODES = {"external": "외부활동만", "always": "항상", "never": "안 함"}


def ext_suffix(it, mode="external"):
    """'(@장소, 참석자)' — 공식 규칙은 외부활동만(기본). 문장에 이미 있는 이름은 다시 붙이지 않는다."""
    place, people = (it.get("place") or "").strip(), (it.get("people") or "").strip()
    t = it.get("text") or ""
    if mode == "never" or not (place or people) or "(@" in t or (mode != "always" and not external_activity(it)):
        return ""
    names = [x.strip() for x in re.split(r"[,·/]", people) if x.strip()]
    base = lambda x: re.sub(r"(님|박사님|박사|선임|책임|팀장님|팀장|실장님|실장)$", "", x)
    people = ", ".join(x for x in names if x not in t and not (len(base(x)) >= 2 and base(x) in t))
    if place and place in t:
        place = ""
    if not (place or people):
        return ""
    return " (" + ", ".join(([f"@{place}"] if place else []) + ([people] if people else [])) + ")"  # 빠진 칸은 규칙 점검이 경고


def ext_warnings(items):
    out = []
    for it in items:
        if not is_external(it) or "(@" in (it.get("text") or ""):
            continue
        miss = [n for n, k in (("장소", "place"), ("참석자", "people")) if not (it.get(k) or "").strip()]
        if miss:
            out.append({"kind": "ext", "level": "warn", "item": it.get("id"),
                        "msg": f"외부활동으로 보임 — (@장소, 참석자) 중 {'·'.join(miss)} 없음: {it.get('text', '')[:40]}"})
    return out


DATE_PATS = [
    ("iso", re.compile(r"(?<!\d)(20\d\d)[-./]\s?(\d{1,2})[-./]\s?(\d{1,2})\.?(?!\d)")),
    ("slash", re.compile(r"(?<![\d/])(\d{1,2})/(\d{1,2})(?![\d/])")),
    ("korean", re.compile(r"(?<!\d)(\d{1,2})월\s?(\d{1,2})일")),
    ("dot", re.compile(r"(?<![\d.])(\d{1,2})\.\s?(\d{1,2})\.(?=\s?[(~∼\-–)]|[\s,]|$)")),
]


def date_styles(text):
    return {k for k, p in DATE_PATS if p.search(text or "")}


def normalize_dates(text, year):
    """날짜 표기를 공식 양식의 'M.D' 하나로(예: 10/15·10월 15일·2026-10-15 → 10.15) — 연도가 다르면 "'27.1.5".
    뒤에 붙은 요일 괄호 '(목)' 는 지운다(양식은 요일을 쓰지 않음)."""
    def fmt(y, m, d, tail):
        try:
            datetime.date(y, m, d)
        except ValueError:
            return None
        return (f"'{y % 100:02d}." if y != year else "") + f"{m}.{d}"

    def sub(pat, fn):
        nonlocal text
        res, pos = [], 0
        for m in pat.finditer(text):
            new = fn(m, text[m.end():m.end() + 4])
            if new is None:
                continue
            res += [text[pos:m.start()], new]
            pos = m.end()
            wd = re.match(r"\s?\([월화수목금토일]\)", text[pos:])
            if wd:
                pos += wd.end()
        text = "".join(res) + text[pos:]

    sub(DATE_PATS[0][1], lambda m, t: fmt(int(m.group(1)), int(m.group(2)), int(m.group(3)), t))
    for k in ("slash", "korean"):
        sub(dict(DATE_PATS)[k], lambda m, t: fmt(year, int(m.group(1)), int(m.group(2)), t) if 1 <= int(m.group(1)) <= 12 else None)
    sub(dict(DATE_PATS)["dot"], lambda m, t: fmt(year, int(m.group(1)), int(m.group(2)), t) if 1 <= int(m.group(1)) <= 12 else None)
    return text


PERIOD_RE = re.compile(r"^(?:\d{1,2}\.\d{1,2})?(?:~(?:\d{1,2}\.\d{1,2})?)?$")


def norm_period(s, year=None):
    """항목 기간 → 공식 양식 표기 '2.24~3.5' / '~3.6' / '3.20' (괄호·공백·요일 제거, 날짜는 M.D)"""
    s = re.sub(r"\s?\([월화수목금토일]\)", "", str(s or "")).strip()
    if s[:1] in "(（" and s[-1:] in ")）":
        s = s[1:-1].strip()
    if not s:
        return ""
    s = normalize_dates(s, year or datetime.date.today().year)
    s = re.sub(r"\s*[~∼～\-–]\s*", "~", s).replace(" ", "")
    return s


def period_ok(p):
    return bool(p) and bool(PERIOD_RE.match(p)) and p != "~"


def date_warnings(items):
    styles = {}
    for it in items:
        for s in date_styles(it.get("text")):
            styles.setdefault(s, it.get("id"))
    if len(styles) > 1:
        return [{"kind": "date", "level": "warn", "msg": f"날짜 표기가 섞여 있음({', '.join(sorted(styles))}) — '날짜 통일'로 양식 표기 'M.D' 하나로 맞출 수 있습니다"}]
    return []


# 분량: A4 세로, 본문 폭 170mm, 표 글자 10pt·줄간격 130%, 기본 양식의 제목 머리(약 11줄) 기준 대략(kordoc 렌더로 맞춤). cpl = 칸별 한 줄 글자 폭(em, 한글 1·영문 0.6).
# 양식 매핑(templates/*.json)의 "page" 로 바꿀 수 있다.
PAGE = {"lines": 54, "cpl": [8, 36], "row_pad": 1, "head_lines": 12, "table_head": 3}


def em(s):
    """대략 글자 폭(em): 한글·전각·도형 기호(○□※ 등) 1, 그 밖(숫자·영문·공백·가운뎃점) 0.6"""
    return sum(1 if ord(ch) >= 0x2460 else 0.6 for ch in s or "")


def row_lines(row, kind, page=PAGE, bbs=False):
    """표(수행 또는 계획) 한 행이 차지할 줄 수(추정, 약어 풀이 줄 제외): 부서 칸과 내용 칸 중 긴 쪽. 정확한 값은 render.estimate."""
    n, cats = 0, set()
    for it in [i for i in sort_items(row.get("items") or []) if i.get("kind") == kind and not (bbs and i.get("nobbs"))]:
        d = int(it.get("depth") or 0)
        if d == 0 and it.get("cat") not in cats:
            cats.add(it.get("cat"))
            n += 1  # □ 분류 머리 줄
        n += max(1, -(-(em(it.get("text")) + em(ext_suffix(it)) + 1.6) // (page["cpl"][1] - 2 * d)))
    left = max(1, -(-em(row.get("label")) // page["cpl"][0])) + (1 if row.get("sub") else 0)
    return max(left, n or 1) + page["row_pad"]


def page_estimate(doc, page=PAGE, bbs=False):
    lines = page["head_lines"] + sum(page["table_head"] + sum(row_lines(r, k, page, bbs) for r in doc.get("rows") or []) for k in KINDS)
    return estimate_result(lines, page)


def estimate_result(lines, page=PAGE):
    return {"lines": lines, "limit": page["lines"], "pages": round(lines / page["lines"], 2), "over": lines > page["lines"]}


def length_warnings(doc, page=PAGE, est=None):
    out = []
    for r in doc.get("rows") or []:
        for it in r.get("items") or []:
            if len(it.get("text") or "") > 110:
                out.append({"kind": "length", "level": "warn", "item": it.get("id"), "msg": f"항목이 깁니다({len(it['text'])}자) — 한 항목은 두 줄 이내 권장"})
    est = est or page_estimate(doc, page)
    if est["over"]:
        out.append({"kind": "page", "level": "warn", "msg": f"1쪽 초과 추정 — 약 {est['lines']}줄 / 기준 {est['limit']}줄 ({est['pages']}쪽). '압축'을 쓰거나 항목을 줄이세요", "est": est})
    return out


def empty_warnings(doc):
    out = []
    for r in doc.get("rows") or []:
        its = r.get("items") or []
        for it in its:
            if not (it.get("text") or "").strip():
                out.append({"kind": "empty", "level": "warn", "item": it.get("id"), "msg": "빈 항목"})
        for k, lab in KINDS.items():
            if not any(i.get("kind") == k and (i.get("text") or "").strip() for i in its):
                out.append({"kind": "empty", "level": "warn", "msg": f"{r.get('label') or '(부서 없음)'}: '{lab}' 칸이 비어 있음"})
    return out


NO_PROJECT = "[과제명 확인 필요]"


def project_warnings(items, rows_of=None):
    """과제명(∙ (과제명) 묶음) 없음·추정 → 작성자 질의(level ask), 기간 괄호 형식 → 경고"""
    out = []
    for it in items:
        if int(it.get("depth") or 0) > 0:
            continue
        if not (it.get("project") or "").strip():
            out.append({"kind": "project", "level": "ask", "item": it.get("id"), "status": "missing", "sentence": (it.get("text") or "")[:160],
                        "msg": f"과제명 없음 — '∙ (과제명)' 아래 묶을 과제·사업 이름을 정해 주세요: {(it.get('text') or '')[:40]}"})
        elif it.get("project_guess"):
            out.append({"kind": "project", "level": "ask", "item": it.get("id"), "status": "guess", "sentence": (it.get("text") or "")[:160],
                        "suggest": it.get("project"), "msg": f"과제명 '{it['project']}' 은 메모에 없던 추정 — 맞는지 확인해 주세요: {(it.get('text') or '')[:40]}"})
        p = (it.get("period") or "").strip()
        if not p:
            out.append({"kind": "period", "level": "warn", "item": it.get("id"),
                        "msg": f"기간 없음 — 양식은 '- (2.24~3.5) 내용', 마감은 '(~3.6)', 하루는 '(3.20)': {(it.get('text') or '')[:40]}"})
        elif not period_ok(p):
            out.append({"kind": "period", "level": "warn", "item": it.get("id"), "msg": f"기간 표기 '{p}' — 'M.D~M.D', '~M.D', 'M.D' 중 하나로"})
    return out


def check_doc(doc, gl, page=PAGE, est=None):
    items = [it for r in doc.get("rows") or [] for it in sort_items(r.get("items") or [])]
    where = {it.get("id"): (r.get("label") or "", it.get("who") or ([r["sub"]] if r.get("sub") else r.get("people") or []))
             for r in doc.get("rows") or [] for it in r.get("items") or []}  # 작성자를 모르면(옛 취합본) 그 부서 제출자 전원
    ab = abbr_warnings(items, gl, doc.get("levels"))
    for w in ab:
        w["dept"], w["who"] = where.get(w.get("item"), ("", []))
    pj = project_warnings(items)
    for w in pj:
        w["dept"], w["who"] = where.get(w.get("item"), ("", []))
    return pj + empty_warnings(doc) + ext_warnings(items) + ab + date_warnings(items) + length_warnings(doc, page, est)


def blocks(items):
    """상위 항목(depth 0) + 바로 뒤 하위 항목(depth 1·2)을 한 덩어리로 — 하위는 상위의 수행/계획·분류를 따른다"""
    out = []
    for it in items:
        if int(it.get("depth") or 0) <= 0 or not out:
            out.append([it])
        else:
            out[-1].append(it)
    return out


def sort_items(items):
    """표에 실리는 순서: 수행 → 계획, 그 안에서는 입력 순서(과제명 묶음은 render 가 첫 등장 순으로). 하위 항목은 상위 바로 뒤.
    분류(cat: 중점목표·성과·행사·기타)는 출력 구조에 쓰지 않는다 — 1쪽 압축 우선순위·작성 안내용."""
    key = lambda b: 0 if b[0].get("kind") == "done" else 1
    return [it for b in sorted(blocks(items), key=key) for it in b]


def plan_period(week_key):
    """'향후 2주' = 다음 주 월요일 ~ 그다음 주 금요일"""
    mon = datetime.date.fromisoformat(week_of(week_key)["mon"]) + datetime.timedelta(days=7)
    fri = mon + datetime.timedelta(days=11)
    return f"{mon.month}. {mon.day}.({WD[0]}) ~ {fri.month}. {fri.day}.({WD[4]})"


# ── LLM 결과 원문 대조 (지어낸 숫자·영문 표기·장소·참석자) ──────────────────────
def numbers_of(text):
    return {re.sub(r"[^\d.]", "", m).strip(".") for m in re.findall(r"\d[\d,]*(?:\.\d+)?", text or "")} - {""}


def latin_of(text):
    return {t.lower() for t in re.findall(r"[A-Za-z][A-Za-z0-9&\-]*", text or "")}


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def date_numbers(text, year):
    """날짜 통일로 생긴 요일·연도는 지어낸 숫자로 치지 않게, 원문 날짜를 통일한 표기의 숫자도 허용 목록에 넣는다"""
    return numbers_of(normalize_dates(text or "", year))


def lost_numbers(it, source):
    """합친 글에서 원래 항목들의 숫자가 빠졌는지 (취합용 — 빠뜨려도 되는 경우가 있어 경고만)"""
    return sorted(numbers_of(source) - numbers_of(" ".join((it.get(k) or "").replace("~", " ") for k in ("text", "period", "place", "people"))))


def verify_item(it, source, year=None):
    """항목 하나를 원문(메모·가져온 기록·지난 계획 등을 이은 글)과 대조 → 경고 목록. 원문에 없는 장소·참석자는 지운다(it 수정)."""
    warns = []
    allow = numbers_of(source) | (date_numbers(source, year) if year else set())
    new_n = sorted(numbers_of((it.get("text") or "") + " " + (it.get("period") or "").replace("~", " ")) - allow)
    if new_n:
        warns.append(f"원문에 없는 숫자: {', '.join(new_n[:5])}")
    new_l = sorted(latin_of(it.get("text")) - latin_of(source))
    if new_l:
        warns.append(f"원문에 없는 영문 표기: {', '.join(new_l[:5])}")
    src = _norm(source)
    for k, lab in (("place", "장소"), ("people", "참석자")):
        v = (it.get(k) or "").strip()
        if not v:
            continue
        parts = [p for p in re.split(r"[,·/]|\s및\s|\s외\s", v) if p.strip()]
        bad = [p.strip() for p in parts if _norm(re.sub(r"\s*(외\s*\d+\s*명|등)$", "", p.strip())) not in src]
        if bad:
            keep = [p.strip() for p in parts if p.strip() not in bad]
            it[k] = ", ".join(keep)
            warns.append(f"원문에 없는 {lab} 지움: {', '.join(bad)}")
    return warns

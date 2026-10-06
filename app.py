#!/usr/bin/env python3
"""weekly local — 주간보고서 작성(개인)·취합(실/소본부). 로컬 LLM, 표준 라이브러리만, 외부 전송 없음.

  python3 app.py                         # http://localhost:8788
  python3 app.py --cli itemize < memo.txt

개인: 메모 → LLM 항목화(수행/계획·분류 a/b/c·외부활동 장소·참석자·표시) → 원문 대조(지어낸 숫자·영문·장소·참석자) → 지난 '향후 2주 계획' 대조 → 제출.
취합: 주차·소본부 → 부서별 제출 모아 LLM 이 중복 합치기·간결화(빠진 항목은 원문 복원) → 부서별 행의 표 → 1쪽 추정·압축 → HWPX/DOCX/HTML/텍스트(BBS 게시용 따로).
규칙 검사(약어·외부활동·분량·빈칸·날짜)는 rules.py 가 LLM 없이 한다. 출력은 render.py (HWPX 는 templates/ 양식 + 자리표시자).
"""
import datetime
import hashlib
import json
import os
import re
import secrets
import subprocess
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import exchange
import projects
import render
import rules

ROOT = os.path.dirname(os.path.abspath(__file__))
WS = os.environ.get("WORKSPACE") or os.path.join(ROOT, "_workspace")  # 포털이 AGENT_DATA/<도구> 로 모아 줌
DATA = os.environ.get("AGENT_DATA") or os.path.dirname(os.path.abspath(WS))  # 같은 서버 다른 도구 기록(읽기 전용)
LLM_API = os.environ.get("LLM_API", "ollama")
LLM_BASE = os.environ.get("LLM_BASE_URL", "http://localhost:8000/v1" if LLM_API == "openai" else "http://localhost:11434").rstrip("/")
MODEL = os.environ.get("LLM_MODEL", "qwen3:8b")
LLM_KEY = os.environ.get("LLM_API_KEY", "")
NUM_CTX = int(os.environ.get("NUM_CTX", "16384"))
PORT = int(os.environ.get("PORT", "8788"))
KORDOC = os.environ.get("KORDOC_CLI") or os.path.join(ROOT, "..", "kordoc-local", "node_modules", "kordoc", "dist", "cli.js")
MAX_MEMO = 8000
GL = rules.Glossary(WS)
REG = projects.Registry(WS)  # 과제 등록부(WORKSPACE/projects.json)

GUIDE = [  # '보통 들어가는 내용' — 입력란 옆 체크리스트
    ("과제 진도·마일스톤", "○○ 과제 2단계 설계 검토 완료(마일스톤 M3)"),
    ("내부·외부 회의", "KINS 와 i-SMR 인허가 사전 협의 (@KINS 대전, 김○○·이○○)"),
    ("출장·국제회의 발표", "IAEA 기술회의 발표 (@오스트리아 빈, 박○○)"),
    ("논문·특허·보고서", "SCI 논문 1편 투고, 특허 출원 1건, 연차보고서 제출"),
    ("시험·실험·해석", "ATLAS 시험 2회 수행, 계통 해석 결과 정리"),
    ("인허가·규제기관 대응", "원안위 질의 답변서 제출"),
    ("과기정통부 보고", "과기정통부 요청 실적 자료 제출 → '과기정통부' 표시"),
    ("행사 주관·참석", "연구원 주관 워크숍 개최 (@본원 국제원자력교육훈련센터, 실원 전원)"),
    ("수상·언론·홍보", "학회 우수논문상 수상, 언론 인터뷰"),
    ("인력·예산·안전", "신규 인력 채용 면접, 예산 집행 점검, 실험실 안전점검"),
]


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


def now():
    return datetime.datetime.now().isoformat(timespec="seconds")


# ── LLM ────────────────────────────────────────────────────────────────
def _clean(out):
    out = re.sub(r"<think>.*?</think>", "", out or "", flags=re.S).strip()
    out = re.sub(r"^```\w*\s*\n", "", out)
    return re.sub(r"\n?```\s*$", "", out).strip()


def llm(system, user, model=None, temperature=0.2, on_token=lambda t: None, json_mode=True):
    """스트리밍 채팅. Ollama /api/chat (format=json, think:false — 미지원이면 빼고 재시도) 또는 OpenAI 호환."""
    model = model or MODEL
    msgs = [{"role": "system", "content": system}, {"role": "user", "content": user}]
    try:
        if LLM_API == "openai":
            body = {"model": model, "stream": True, "temperature": temperature, "messages": msgs}
            if json_mode:
                body["response_format"] = {"type": "json_object"}
            hdr = {"Content-Type": "application/json", **({"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})}
            buf = []
            req = urllib.request.Request(LLM_BASE + "/chat/completions", json.dumps(body).encode(), hdr)
            with urllib.request.urlopen(req, timeout=1800) as r:
                for line in r:
                    line = line.decode().strip()
                    if not line.startswith("data:") or line == "data: [DONE]":
                        continue
                    tok = (json.loads(line[5:])["choices"][0].get("delta") or {}).get("content") or ""
                    if tok:
                        buf.append(tok)
                        on_token(tok)
            return "".join(buf)
        body = {"model": model, "stream": True, "think": False, "messages": msgs,
                "options": {"temperature": temperature, "num_ctx": NUM_CTX}}
        if json_mode:
            body["format"] = "json"
        for attempt in (0, 1):
            try:
                req = urllib.request.Request(LLM_BASE + "/api/chat", json.dumps(body).encode(), {"Content-Type": "application/json"})
                buf = []
                with urllib.request.urlopen(req, timeout=1800) as r:
                    for line in r:
                        if not line.strip():
                            continue
                        j = json.loads(line)
                        if "error" in j:
                            raise RuntimeError(j["error"])
                        tok = j.get("message", {}).get("content", "")
                        if tok:
                            buf.append(tok)
                            on_token(tok)
                        if j.get("done"):
                            break
                return "".join(buf)
            except urllib.error.HTTPError as e:
                msg = e.read().decode(errors="replace")
                if attempt == 0 and "think" in msg:
                    body.pop("think")
                    continue
                raise RuntimeError(f"LLM HTTP {e.code}: {msg[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"LLM 서버 연결 실패 ({LLM_BASE}): {e.reason}")


def llm_json(system, user, model, emit, label):
    """JSON 응답을 받는다. 깨지면 한 번 더(더 엄격히)."""
    for attempt in (0, 1):
        emit({"stage": "llm", "msg": f"{model} · {label}" + (" · 형식 재시도" if attempt else ""), "reset": bool(attempt)})
        raw = llm(system, user + ("\n\n반드시 설명 없이 JSON 객체 하나만 출력한다." if attempt else ""), model, 0.2 if not attempt else 0.1,
                  on_token=lambda t: emit({"token": t}))
        txt = _clean(raw)
        m = re.search(r"\{.*\}", txt, re.S)
        if m:
            try:
                return json.loads(m.group(0)), raw
            except ValueError:
                pass
    raise RuntimeError("모델 응답을 JSON 으로 읽지 못했습니다. 다시 시도해 주세요.")


def models():
    try:
        if LLM_API == "openai":
            req = urllib.request.Request(LLM_BASE + "/models", headers={"Authorization": f"Bearer {LLM_KEY}"} if LLM_KEY else {})
            with urllib.request.urlopen(req, timeout=3) as r:
                return [m["id"] for m in json.load(r)["data"]]
        with urllib.request.urlopen(LLM_BASE + "/api/tags", timeout=3) as r:
            return [m["name"] for m in json.load(r)["models"]]
    except Exception:
        return []


# ── 항목 정리 ───────────────────────────────────────────────────────────
def new_id():
    return secrets.token_hex(4)


def norm_item(x, **extra):
    """LLM·화면에서 온 항목을 정해진 모양으로"""
    t = re.sub(r"\s+", " ", str(x.get("text") or "")).strip().lstrip("-•○□·*└ ").strip()
    t = FLAG_TAG.sub(" ", t).strip()  # 취합·압축 입력의 참고 표시 ‹…› 를 LLM 이 글에 베낀 경우
    m = re.search(r"\s*\(@\s*([^,)]*?)\s*(?:,\s*([^)]*))?\)\s*$", t)  # 메모에 이미 (@장소, 참석자) 를 써 둔 경우 칸으로 옮긴다
    place, people = str(x.get("place") or "").strip(), str(x.get("people") or "").strip()
    if m:
        t = t[:m.start()].rstrip()
        place, people = place or (m.group(1) or "").strip(), people or (m.group(2) or "").strip()
    t = strip_dup_place(t, place, people)
    kind = x.get("kind") if x.get("kind") in rules.KINDS else ("plan" if re.search(r"예정|계획|할 것", t) else "done")
    cat = x.get("cat") if x.get("cat") in rules.CATS else "etc"
    tf = lambda v: v is True or str(v).lower() in ("true", "1", "yes", "y")
    try:
        depth = max(0, min(2, int(x.get("depth") or 0)))
    except (TypeError, ValueError):
        depth = 0
    proj = re.sub(r"\s+", " ", str(x.get("project") or "")).strip().strip("()[]（）").strip()
    period = str(x.get("period") or "")
    pm = re.match(r"^\(([\d.~∼\-–/\s월일]+)\)\s*", t)  # 글 앞에 '(2.24~3.5)' 를 써 둔 경우 기간 칸으로
    if pm and not period:
        period, t = pm.group(1), t[pm.end():]
    it = {"id": x.get("id") or new_id(), "text": t, "kind": kind, "cat": cat, "depth": depth, "project": proj, "period": rules.norm_period(period, x.get("_year")),
          "ext": tf(x.get("ext")), "place": place, "people": people, "core": tf(x.get("core")), "msit": tf(x.get("msit")), "nobbs": tf(x.get("nobbs"))}
    if tf(x.get("project_guess")) and proj:
        it["project_guess"] = True
    for k in ("warn", "src", "carry", "who"):
        if x.get(k):
            it[k] = x[k]
    it.update(extra)
    return it


FLAG_TAG = re.compile(r"\s*‹[^›]*›\s*|^(?:\s*\((?:핵심|과기정통부|비게시|수행한 일|향후 2주 계획|수행|계획|goal|perf|event|etc)\))+\s*")


def flatten(xs, depth=0, parent=None):
    """LLM 의 항목 트리 {"text", …, "children": […]} → depth 를 가진 평평한 목록(하위는 상위 바로 뒤, 수행/계획·분류는 상위를 따름)"""
    out = []
    for x in xs or []:
        if not isinstance(x, dict) or not str(x.get("text") or "").strip():
            continue
        x = dict(x, depth=depth)
        if parent:
            x.update(kind=parent.get("kind"), cat=parent.get("cat"), project=parent.get("project"))
        out.append(x)
        if depth < 2:
            out += flatten(x.get("children") or x.get("sub") or [], depth + 1, x)
    return out


def strip_dup_place(t, place, people):
    """LLM 이 장소·참석자를 글 안 괄호에도 쓴 경우 — '(10월 22일, 본원 ○○센터)' → '(10월 22일)', '(프랑스 파리)' → 삭제"""
    vals = [v.strip() for v in [place] + re.split(r"\s*[,·]\s*", people or "") if v and v.strip()]
    if not vals:
        return t

    def fix(m):
        lead, inner = m.group(1), m.group(2)
        if not any(v in inner for v in vals):
            return m.group(0)
        for v in sorted(vals, key=len, reverse=True):
            inner = inner.replace(v, "")
        inner = re.sub(r"^[\s,·@]+|[\s,·@]+$", "", re.sub(r"\s*,(\s*,)+\s*", ", ", inner))
        return f"{lead}({inner})" if inner else ""
    return re.sub(r"(\s*)\(([^()]*)\)", fix, t).strip()


ITEMIZE_SYS = """너는 한국원자력연구원(KAERI) 연구자의 주간보고 메모를 공식 주간보고 항목으로 정리하는 비서다.
공식 양식의 칸 안 구조: '∙ (과제명)' 한 줄 아래에 '- (기간) 내용' 들. 그래서 항목마다 과제명(project)과 기간(period)을 따로 뽑는다.
규칙:
- 메모와 [가져온 기록]에 적힌 사실만 쓴다. 적혀 있지 않은 숫자·날짜·사람 이름·장소·기관·성과·결과를 절대 만들지 않는다.
- project: 그 일이 속한 과제·사업 이름(예: 기본사업, 전략개발단, ○○ 과제). 메모에 적힌 표기 그대로. 메모의 '[과제명]'·'과제명:' 머리 아래 줄들은 그 과제.
  [내 과제 목록]이 있으면 메모 내용이 분명히 그 과제일 때만 그 이름을 쓴다. 모르면 "" (지어내지 않는다 — 프로그램이 작성자에게 묻는다).
- period: 그 일의 기간. 형식 'M.D~M.D'(기간), '~M.D'(마감), 'M.D'(하루). 메모에 날짜가 없으면 "". 기간은 text 에 다시 쓰지 않는다.
- text: 한 항목 = 한 가지 일. 공문서 개조식 명사형 종결(…완료, …발표, …제출, …개최, …착수, …검토, …예정). 70자 이내. 존댓말·마크다운 금지.
  뜻 없이 '수행'을 덧붙이지 않는다. 숫자·과제명·장비명·약어(MARS-KS, IAEA 등)는 메모 표기 그대로, 약어 풀이는 붙이지 않는다(프로그램이 붙인다).
- kind: "done" = 이번 주에 한 일, "plan" = 앞으로 2주 안에 할 일(예정·계획·다음 주·할 것).
- cat(출력에는 안 쓰는 내부 분류 — 1쪽으로 줄일 때 우선순위): "goal" 중점목표·주요 과제 진도·핵심 연구, "perf" 연구·경영 성과·대외활동,
  "event" 연구원 주관 행사·수상, "etc" 그 밖의 일반 업무.
- ext: 연구원 밖에서 하거나 외부 기관·외부인과 함께한 회의·발표·참석·방문·출장이면 true. 서류 제출·수상·내부 업무는 false.
- place(장소)·people(참석자): 메모에 적힌 그대로만, 없으면 "". 온라인이면 place "온라인". 장소·이름은 text 에 다시 쓰지 않는다(협의 상대 기관 이름은 남김).
- core: '핵심', '중요', '★' 표시가 있는 일만 true. msit: 과기정통부(과기부, MSIT) 보고·제출 관련만 true.
  nobbs: '비공개', '대외비', '게시 금지', 'BBS 제외' 표시가 있는 일만 true(이 낱말은 text 에서 뺀다).
- [지난 계획]이 주어지면 번호마다 상태: "완료" / "진행" / "미착수" / "미확인"(언급 없음), evidence 는 근거 메모 구절 그대로(없으면 ""). 이번 주에 한 지난 계획은 items 에도 done 으로.
- 한 일의 세부 사항(들여쓴 줄)은 그 항목의 children 으로(최대 2단계). children 은 project 를 쓰지 않아도 된다(상위를 따름).
- remarks: 메모에 '특기', '애로', '건의' 사항이 있으면 한 줄씩, 없으면 [].
출력: JSON 객체 하나 — {"items":[{"project":"","period":"","text":"","kind":"done","cat":"goal","ext":false,"place":"","people":"","core":false,"msit":false,"nobbs":false,
         "children":[{"period":"","text":"","ext":false,"place":"","people":"","core":false,"msit":false,"nobbs":false,"children":[]}]}],
       "carry":[{"n":1,"status":"완료","evidence":""}], "remarks":[]}"""


def known_projects(prof, week_key=None, weeks=8):
    """과제명 후보: 과제 등록부(등록 순서) 먼저, 그다음 같은 부서(실)의 최근 제출에서 쓴 과제명(많이 쓴 순). 질의 카드·LLM 참고용"""
    reg = [p["name"] for p in REG.active(prof.get("org"), prof.get("dept"), prof.get("name") or None)] if prof.get("dept") else []
    if not reg and prof.get("dept"):  # 소본부를 안 적었으면 같은 실 이름의 등록부
        reg = [p["name"] for u in REG.units() if u["dept"] == prof.get("dept") for p in REG.active(u["org"], u["dept"])]
    return list(dict.fromkeys(reg + _used_projects(prof, week_key, weeks)))[:40]


def _used_projects(prof, week_key=None, weeks=8):
    cnt = {}
    base = datetime.date.fromisoformat(rules.week_of(week_key or None)["mon"])
    for k in range(0, weeks + 1):
        wk = rules.week_of(base - datetime.timedelta(days=7 * k))["key"]
        for r in list_reports(wk):
            if r.get("dept") != prof.get("dept"):
                continue
            w = 3 if r.get("name") == prof.get("name") else 1
            for it in r.get("items") or []:
                pj = (it.get("project") or "").strip()
                if pj:
                    cnt[pj] = cnt.get(pj, 0) + w
    return [p for p, _ in sorted(cnt.items(), key=lambda x: -x[1])][:30]


def prev_report(profile, week_key, back=4):
    """같은 사람의 지난 제출(최근 back 주 안) — '향후 2주 계획' 대조용"""
    mon = datetime.date.fromisoformat(rules.week_of(week_key)["mon"])
    rid = report_id(profile)
    for k in range(1, back + 1):
        wk = rules.week_of(mon - datetime.timedelta(days=7 * k))["key"]
        p = os.path.join(WS, "reports", wk, rid + ".json")
        if os.path.exists(p):
            return rules.load_json(p, None)
    return None


def run_itemize(req, emit, model):
    prof = req.get("profile") or {}
    week = rules.week_of(req.get("week") or None)
    memo = (req.get("memo") or "").replace("\r\n", "\n").strip()
    imports = [str(x) for x in req.get("imports") or [] if str(x).strip()]
    if not memo and not imports:
        raise ValueError("메모를 적거나 기록을 가져와 주세요")
    if len(memo) > MAX_MEMO:
        raise ValueError(f"메모가 너무 깁니다 ({len(memo)}자 > {MAX_MEMO}자)")
    prev = prev_report(prof, week["key"]) if req.get("use_prev", True) else None
    prev_plans = [i for i in (prev or {}).get("items") or [] if i.get("kind") == "plan"]
    lines = [f"[주차] {week['label']}", f"[작성자] {prof.get('org', '')} {prof.get('dept', '')} {prof.get('name', '')}".strip(),
             f"[메모]\n{memo or '(없음)'}"]
    if imports:
        lines.append("[가져온 기록]\n" + "\n---\n".join(x[:3000] for x in imports))
    if prev_plans:
        lines.append("[지난 계획]\n" + "\n".join(f"{n}. {i['text']}" for n, i in enumerate(prev_plans, 1)))
    projects = known_projects(prof, week["key"])
    if projects:
        lines.append("[내 과제 목록] " + ", ".join(projects))
    data, raw = llm_json(ITEMIZE_SYS, "\n\n".join(lines), model, emit, "항목 정리")
    year = datetime.date.fromisoformat(week["mon"]).year
    source = "\n".join([memo] + imports + [i["text"] for i in prev_plans])
    items, notes = [], []
    src_norm = re.sub(r"\s+", "", memo + "".join(imports))
    for x in flatten(data.get("items")):
        x["_year"] = year
        pj = re.sub(r"\s+", "", str(x.get("project") or "")).strip("()[]（）")
        if pj and pj not in src_norm:
            x["project_guess"] = True  # 메모에 없는 과제명(목록에서 고른 추정) → 작성자 확인
        it = norm_item(x)
        it["text"] = rules.normalize_dates(it["text"], year)  # 날짜 표기 통일 'M.D'(양식)
        if it["ext"] and not (it["place"] or it["people"] or any(w in it["text"] for w in rules.EXT_WORDS)):
            it["ext"] = False  # 수상·서류 제출 같은 것을 외부활동으로 잘못 표시한 경우
        w = rules.verify_item(it, source, year)
        if w:
            it["warn"] = w
            notes.append(f"'{it['text'][:30]}…': " + " / ".join(w))
        items.append(it)
    lost = rules.lost_numbers({"text": " ".join(i["text"] + " " + i["period"].replace("~", " ") + " " + i["place"] + " " + i["people"] for i in items)},
                              rules.normalize_dates(memo, year))
    if lost:
        notes.append("메모의 숫자가 항목에 없음(빠뜨렸는지 확인): " + ", ".join(lost[:8]))
    carry = []
    for c in data.get("carry") or []:
        try:
            n = int(c.get("n")) - 1
        except (TypeError, ValueError):
            continue
        if 0 <= n < len(prev_plans):
            st = c.get("status") if c.get("status") in ("완료", "진행", "미착수", "미확인") else "미확인"
            ev = str(c.get("evidence") or "").strip()
            if ev and re.sub(r"\s+", "", ev) not in re.sub(r"\s+", "", memo + "".join(imports)):
                ev = ""  # 메모에 없는 근거는 버림
            if st == "완료" and not ev:
                st = "미확인"
            carry.append({"prev": prev_plans[n], "status": st, "evidence": ev})
    for n, p in enumerate(prev_plans):
        if not any(c["prev"] is p for c in carry):
            carry.append({"prev": p, "status": "미확인", "evidence": ""})
    questions = abbr_questions(items, model, emit, prof.get("dept", ""), projects)
    remarks = [str(x).strip() for x in data.get("remarks") or [] if str(x).strip()]
    return {"items": items, "carry": carry, "notes": notes, "questions": questions, "remarks": remarks, "projects": projects,
            "prev_week": (prev or {}).get("week"), "week": week, "raw": raw}


SUGGEST_SYS = """너는 약어 풀이 도우미다. 한국원자력연구원 연구자의 주간보고에 나온 약어의 영문 전체 이름 후보를 제안한다.
규칙:
- 문장 맥락과 일반적으로 널리 쓰이는 뜻으로 확실히 아는 것만 쓴다. 조금이라도 불확실하면 full 을 "" 로 둔다(지어내지 않는다).
- 이 제안은 작성자가 확인하기 전에는 문서에 쓰이지 않는다.
- field 는 "원자력" | "기관·정책" | "AI" | "컴퓨터" | "일반" 중 하나. desc 는 쉬운 한국어 한 줄(30자 이내, 모르면 "").
출력: JSON 객체 하나 — {"abbrs":[{"abbr":"","full":"","ko":"","field":"","desc":""}]}"""


def abbr_questions(items, model, emit, dept="", projects=None):
    """약어집으로 확정 못 한 약어 → 작성자에게 물을 질문 카드. 후보: 약어집(동음이의·공개) → LLM 제안(확인 필요) 순.
    LLM 제안은 후보로만 보여 주고, 작성자가 고르거나 직접 적기 전에는 문서에 넣지 않는다."""
    allq = rules.unresolved(items, GL)
    pq = [dict(q, candidates=[{"full": x, "src": "이 실에서 쓴 과제명"} for x in (projects or [])]) for q in allq if q["kind"] == "project"]
    qs = [q for q in allq if q["kind"] == "abbr"]
    if not qs:
        return pq
    try:
        data, _ = llm_json(SUGGEST_SYS, f"[부서] {dept}\n" + "\n".join(f"- {q['abbr']}: \"{q['sentence']}\"" for q in qs), model, emit, f"약어 후보 제안 {len(qs)}개")
        sug = {str(x.get("abbr")): x for x in data.get("abbrs") or [] if isinstance(x, dict) and str(x.get("full") or "").strip()}
    except RuntimeError:
        sug = {}
    for q in qs:
        x = sug.get(q["abbr"])
        if x and not any(c["full"].lower() == str(x["full"]).strip().lower() for c in q["candidates"]):
            q["candidates"].append({"full": str(x["full"]).strip(), "ko": str(x.get("ko") or "").strip(), "desc": str(x.get("desc") or "").strip(),
                                    "field": x.get("field") if x.get("field") in rules.FIELDS else "일반", "src": "LLM 제안(확인 필요)", "suggest": True})
    return qs + pq


# ── 저장 (WORKSPACE/reports/<주차>/<사람id>.json, agg/<주차>/<소본부id>.json) ─────
def slug(*parts):
    return hashlib.sha1("|".join(str(p or "").strip() for p in parts).encode()).hexdigest()[:12]


def report_id(prof):
    return slug(prof.get("org"), prof.get("dept"), prof.get("name"))


def _wk(key):
    if not re.fullmatch(r"\d{4}-W\d{2}", key or ""):
        raise ValueError("잘못된 주차")
    return key


class Unresolved(ValueError):
    """미확인 약어·과제명이 남아 제출을 막음 — 화면이 질문 카드를 다시 띄운다"""

    def __init__(self, rows):
        self.rows = rows
        ab = [w["abbr"] for w in rows if w.get("abbr")]
        pj = [w for w in rows if w.get("kind") == "project"]
        msg = []
        if ab:
            msg.append(f"확인되지 않은 약어 {len(ab)}개({', '.join(ab[:6])}) — 풀이를 답하거나 '약어 아님'으로 표시")
        if pj:
            msg.append(f"과제명이 없거나 확인 안 된 항목 {len(pj)}개 — 과제명을 정하세요")
        super().__init__(" / ".join(msg) + " 한 뒤 제출하세요")


def submit(req):
    prof = req.get("profile") or {}
    if not (prof.get("name") or "").strip() or not (prof.get("dept") or "").strip():
        raise ValueError("이름과 부서(실/팀)를 적어 주세요")
    week = rules.week_of(req.get("week") or None)
    rid = report_id(prof)
    p = os.path.join(WS, "reports", week["key"], rid + ".json")
    old = rules.load_json(p, {})
    items = [norm_item(x) for x in req.get("items") or [] if (x.get("text") or "").strip()]
    if not items:
        raise ValueError("항목이 없습니다")
    unk = rules.unresolved(items, GL)
    reason = (req.get("unresolved_reason") or "").strip()
    if unk and not reason:
        raise Unresolved(unk)
    rec = {"id": rid, "week": week["key"], "org": prof.get("org", "").strip(), "dept": prof["dept"].strip(), "name": prof["name"].strip(),
           "memo": req.get("memo") or "", "items": items, "carry": req.get("carry") or [], "ts": now(),
           "remarks": [str(x).strip() for x in req.get("remarks") or [] if str(x).strip()],
           "history": (old.get("history") or []) + [{"ts": now(), "n": len(items)}],
           "unresolved": [{"abbr": w.get("abbr") or "", "kind": w.get("kind"), "sentence": w.get("sentence", "")} for w in unk], "unresolved_reason": reason if unk else ""}
    rules.save_json(p, rec)
    return {"ok": True, "id": rid, "week": week["key"], "path": os.path.relpath(p, WS)}


def list_reports(week_key, org=None):
    d = os.path.join(WS, "reports", _wk(week_key))
    out = []
    if os.path.isdir(d):
        for fn in sorted(os.listdir(d)):
            if fn.endswith(".json"):
                r = rules.load_json(os.path.join(d, fn), None)
                if r and (org is None or r.get("org") == org):
                    out.append(r)
    return out


def orgs():
    return rules.load_json(os.path.join(WS, "orgs.json"), {})


def save_orgs(o):
    clean = {}
    for k, v in (o or {}).items():
        k = str(k).strip()
        if k:
            clean[k] = [s.strip() for s in (v if isinstance(v, list) else str(v).split(",")) if str(s).strip()]
    rules.save_json(os.path.join(WS, "orgs.json"), clean)
    return clean


def status(week_key, org):
    reps = list_reports(week_key, org)
    depts = list(orgs().get(org) or [])
    for r in reps:
        if r["dept"] not in depts:
            depts.append(r["dept"])
    by = {d: [{"id": r["id"], "name": r["name"], "ts": r["ts"], "n": len(r["items"])} for r in reps if r["dept"] == d] for d in depts}
    agg = rules.load_json(agg_path(week_key, org), None)
    return {"week": rules.week_of(week_key), "org": org, "depts": [{"dept": d, "reports": by[d], "missing": not by[d]} for d in depts],
            "agg": agg}


def agg_path(week_key, org):
    return os.path.join(WS, "agg", _wk(week_key), slug(org) + ".json")


# ── 취합 ────────────────────────────────────────────────────────────────
AGG_SYS = """너는 한국원자력연구원 소본부 주간보고 취합 담당자다. 한 부서의 여러 사람이 낸 주간보고 항목을 부서 하나의 항목 목록으로 합친다.
규칙:
- 입력 항목에 있는 사실만 쓴다. 숫자·날짜·이름·장소·과제명·약어는 원문 그대로. 새 사실·평가·수식어를 만들지 않는다.
- 원래 항목의 숫자·날짜(횟수·건수·일자)는 합친 글에도 남긴다. '(@장소, 참석자)' 는 text 에 쓰지 않는다(프로그램이 붙인다).
- 같은 일·같은 회의·같은 과제를 여러 사람이 쓴 항목은 하나로 합친다(src 에 원래 id 를 모두). 다른 일은 합치지 않는다.
- 수행(done)과 계획(plan)은 섞지 않는다. 표시(핵심/과기정통부/비게시)가 다른 항목끼리는 합치지 않는다. 과제명이 다른 항목끼리는 합치지 않는다.
- period: 합친 일의 기간('M.D~M.D', '~M.D', 'M.D'). 원래 항목들의 기간 안에서만(없으면 ""). text 에는 기간을 쓰지 않는다.
- 문장은 공문서 개조식(…완료, …발표, …제출, …예정)으로 간결하게, 한 항목 70자 이내. 원래 글에 없던 '수행'·'진행'을 덧붙이지 않는다. 합치지 않는 항목은 원래 글을 거의 그대로 둔다.
- 순서: cat 은 goal → perf → event → etc, 같은 cat 안에서는 핵심·과기정통부 표시 항목과 중요한 일을 앞에.
- 모든 입력 id 는 결과 어딘가의 src 에 꼭 들어가야 한다(빠뜨리지 않는다).
- 줄 끝 ‹…› 는 참고 정보(작성자·수행/계획·분류·표시)다. text 에 옮겨 쓰지 않는다.
출력: JSON 객체 하나 — {"items":[{"src":["a1","a2"],"period":"","text":"","kind":"done","cat":"goal"}]}"""


def _meta(it, *head):
    """LLM 입력 줄 끝의 참고 표시 ‹수행·goal·핵심› — 글과 섞이지 않게 꺾쇠로 따로 둔다(결과 글에 베끼면 norm_item 이 지움)"""
    kind = "수행" if it.get("kind") == "done" else "계획"
    return " ‹" + "·".join([*head, "과제 " + (it.get("project") or "미정"), kind, "기간 " + (it.get("period") or "-"), it.get("cat") or "etc"] + [x for k, x in (("core", "핵심"), ("msit", "과기정통부"), ("nobbs", "비게시")) if it.get(k)]) + "›"


def merge_dept(dept, src_items, model, emit, year=None):
    """부서 하나: 상위 항목만 LLM 으로 합치고, 표시·장소·참석자는 원래 항목에서 결정론으로 모은다. 빠진 항목은 원문 그대로 복원.
    하위 항목(depth ≥ 1)은 LLM 에 보내지 않고, 합친 상위 항목 아래에 원래 상위 순서대로 그대로 붙인다."""
    kids = {}
    tops = []
    for it in src_items:
        if int(it.get("depth") or 0) > 0 and tops:
            kids.setdefault(tops[-1]["sid"], []).append(it)
        else:
            tops.append(it)
    src_items = tops
    ids = {it["sid"]: it for it in src_items}
    year = year or datetime.date.today().year

    def with_kids(head, srcs):
        return [head] + [dict(norm_item(k), kind=head["kind"], cat=head["cat"], who=[k.get("who", "")]) for s in srcs for k in kids.get(s, [])]
    if len(src_items) <= 1:
        return [x for it in src_items for x in with_kids(dict(norm_item(it), src=[it["sid"]], who=[it.get("who", "")]), [it["sid"]])], []
    user = f"[부서] {dept}\n[항목]\n" + "\n".join(
        f"[{it['sid']}] {it['text']}"
        + (f" (@{it.get('place', '')}, {it.get('people', '')})" if it.get("place") or it.get("people") else "") + _meta(it, it.get("who", ""))
        for it in src_items)
    data, _ = llm_json(AGG_SYS, user, model, emit, f"{dept} 합치기 ({len(src_items)}항목)")
    out, notes, used = [], [], set()
    for x in data.get("items") or []:
        if not isinstance(x, dict) or not str(x.get("text") or "").strip():
            continue
        srcs = [s for s in (x.get("src") or []) if s in ids and s not in used]
        if not srcs or len({(ids[s].get("project") or "") for s in srcs}) > 1:
            continue  # 다른 과제끼리 합친 것은 받지 않음 → 아래에서 원문 복원
        used.update(srcs)
        group = [ids[s] for s in srcs]
        it = norm_item({"text": x["text"], "kind": group[0]["kind"], "cat": x.get("cat") if x.get("cat") in rules.CATS else group[0]["cat"]})
        it["kind"] = group[0]["kind"] if len({g["kind"] for g in group}) == 1 else x.get("kind", group[0]["kind"])
        for k in ("core", "msit", "nobbs", "ext"):
            it[k] = any(g.get(k) for g in group)
        if len({bool(g.get("nobbs")) for g in group}) > 1:
            notes.append(f"{dept}: 게시/비게시 항목이 합쳐져 비게시(취소선)로 둠 — '{it['text'][:30]}'")
        for k in ("place", "people"):
            vals = []
            for g in group:
                for v in re.split(r"\s*,\s*", g.get(k) or ""):
                    if v and v not in vals:
                        vals.append(v)
            it[k] = ", ".join(vals)
        it["src"] = srcs
        it["who"] = sorted({g.get("who", "") for g in group})
        it["project"] = group[0].get("project") or ""
        pers = {g.get("period") or "" for g in group}
        it["period"] = pers.pop() if len(pers) == 1 else rules.norm_period(x.get("period") or "", year)
        src_text = "\n".join(g["text"] + " " + (g.get("period") or "").replace("~", " ") + " " + g.get("place", "") + " " + g.get("people", "") for g in group)
        it["text"] = rules.normalize_dates(it["text"], year)
        w = rules.verify_item(it, src_text, year)
        lost = rules.lost_numbers(it, rules.normalize_dates(src_text, year))
        if lost:
            w.append(f"원래 항목의 숫자가 빠짐: {', '.join(lost[:5])}")
        if w:
            it["warn"] = w
            notes.append(f"{dept} '{it['text'][:30]}…': " + " / ".join(w))
        out += with_kids(it, srcs)
    missed = [s for s in ids if s not in used]
    for s in missed:  # LLM 이 빠뜨린 항목은 원문 그대로 되살린다
        it = norm_item(ids[s], src=[s], who=[ids[s].get("who", "")])
        it["id"] = new_id()
        it["warn"] = ["취합 과정에서 빠져 원문 그대로 복원"]
        out += with_kids(it, [s])
    if missed:
        notes.append(f"{dept}: LLM 이 빠뜨린 {len(missed)}개 항목을 원문 그대로 복원")
    return out, notes


def run_aggregate(req, emit, model):
    week = rules.week_of(req.get("week") or None)
    org = (req.get("org") or "").strip()
    if not org:
        raise ValueError("소본부를 고르세요")
    st = status(week["key"], org)
    rows, notes = [], []
    reps = list_reports(week["key"], org)
    for d in st["depts"]:
        src = []
        for r in reps:
            if r["dept"] != d["dept"]:
                continue
            for n, it in enumerate(r["items"]):
                src.append(dict(it, sid=f"{r['id'][:6]}{n}", who=r["name"]))
        if not src:
            notes.append(f"{d['dept']}: 제출 없음")
            rows.append({"label": d["dept"], "items": [], "missing": True})
            continue
        items, nn = merge_dept(d["dept"], src, model, emit, datetime.date.fromisoformat(week["mon"]).year)
        notes += nn
        rows.append({"label": d["dept"], "items": items, "people": sorted({s["who"] for s in src})})
    remarks = [f"({r['dept']}) {x}" for r in reps for x in r.get("remarks") or []]
    agg = {"week": week["key"], "org": org, "ts": now(), "rows": rows, "notes": notes, "model": model, "remarks": remarks}
    rules.save_json(agg_path(week["key"], org), agg)
    return agg


COMPRESS_SYS = """너는 주간보고 편집자다. 소본부 주간보고가 1쪽을 넘어 줄여야 한다. 목표 글자 수 안으로 줄인다.
규칙:
- ‹…› 에 핵심·과기정통부 표시가 있는 항목은 지우지 않는다. 문장만 짧게 다듬을 수 있다.
- 소본부 보고는 (a) goal 중점목표 (b) perf 성과·대외활동 (c) event 행사·수상 위주다. 줄일 때는 etc(기타) → event → perf 순으로 drop 하고,
  goal 은 지우기보다 문장을 짧게 쓴다. 목표에 못 미치면 drop 을 더 넣는다.
- 숫자·이름·과제명·약어는 원문 그대로. 새 사실을 만들지 않는다. 기간(‹…› 의 '기간')은 따로 나가므로 text 에 쓰지 않는다.
- 줄 끝 ‹…› 는 참고 정보(부서·과제·수행/계획·기간·분류·표시)다. text 에 옮겨 쓰지 않는다.
출력: JSON 객체 하나 — {"items":[{"id":"r0i3","text":"짧게 고친 글"}],"drop":["r1i5"]}  (바꾸지 않을 항목은 쓰지 않아도 된다)"""


def run_compress(req, emit, model):
    agg = req.get("agg") or {}
    tpl = render.template(req.get("template"))
    doc = agg_doc(agg)
    est = render.estimate(doc, GL, tpl)
    idx, lines = {}, []
    for ri, r in enumerate(agg.get("rows") or []):
        for ii, it in enumerate(r.get("items") or []):
            key = f"r{ri}i{ii}"
            idx[key] = it
            lines.append(f"[{key}] {'  ' * int(it.get('depth') or 0)}{it['text']}{_meta(it, r['label'])}")
    total = sum(len(it["text"]) for it in idx.values())
    target = int(total * tpl["page"]["lines"] / max(est["lines"], 1) * 0.92)
    user = f"[현재] 약 {est['lines']}줄 (1쪽 기준 {tpl['page']['lines']}줄), 항목 글자 합 {total}자 → 목표 {target}자 이하\n[항목]\n" + "\n".join(lines)
    data, _ = llm_json(COMPRESS_SYS, user, model, emit, "1쪽 압축")
    notes, keep_drop = [], set()
    drops = [k for k in data.get("drop") or [] if k in idx]
    prio = {"etc": 0, "event": 1, "perf": 2, "goal": 3}  # 1쪽으로 줄일 때 먼저 빼는 순서: 기타 → 행사·수상 → 성과·대외활동 → 중점목표
    droppable = lambda it: not (it.get("core") or it.get("msit"))
    for k in sorted(drops, key=lambda k: prio.get(idx[k].get("cat"), 0)):
        it = idx[k]
        if not droppable(it):
            notes.append(f"표시 항목은 지우지 않음: {it['text'][:30]}")
            continue
        lower = [x for x, v in idx.items() if droppable(v) and prio.get(v.get("cat"), 0) < prio.get(it.get("cat"), 0) and x not in keep_drop]
        if lower:  # 덜 중요한 분류가 남아 있으면 이 항목은 남긴다
            notes.append(f"'{rules.CATS.get(idx[lower[0]].get('cat'), '기타')}' 항목이 남아 있어 지우지 않음: {it['text'][:30]}")
            continue
        keep_drop.add(k)
    for x in data.get("items") or []:
        it = idx.get(x.get("id"))
        t = str(x.get("text") or "").strip()
        if it is None or not t or x.get("id") in keep_drop:
            continue
        t = rules.normalize_dates(FLAG_TAG.sub(" ", t).strip().lstrip("└- ").strip(), datetime.date.fromisoformat(rules.week_of(agg.get("week") or None)["mon"]).year)
        per = it.get("period") or ""
        if per:  # LLM 이 기간을 글에 베낀 경우 지운다(기간은 따로 '(기간)' 으로 나감)
            t = re.sub(r"^\(?" + re.escape(per) + r"\)?\s*|\s*\(" + re.escape(per) + r"\)", "", t).strip()
        cand = dict(it, text=t)
        w = rules.verify_item(cand, it["text"] + " " + per.replace("~", " ") + " " + it.get("place", "") + " " + it.get("people", ""))
        if w:
            notes.append(f"원문과 달라 고치지 않음: {it['text'][:30]} ({' / '.join(w)})")
            continue
        it["text"] = t
    for ri, r in enumerate(agg.get("rows") or []):  # 상위 항목을 빼면 그 하위도 같이
        head = None
        for ii, it in enumerate(r.get("items") or []):
            if int(it.get("depth") or 0) == 0:
                head = f"r{ri}i{ii}"
            elif head in keep_drop:
                keep_drop.add(f"r{ri}i{ii}")
    for ri, r in enumerate(agg.get("rows") or []):
        dropped = [it for ii, it in enumerate(r.get("items") or []) if f"r{ri}i{ii}" in keep_drop]
        r["items"] = [it for ii, it in enumerate(r.get("items") or []) if f"r{ri}i{ii}" not in keep_drop]
        if dropped:
            r.setdefault("dropped", []).extend(dropped)
            notes.append(f"{r['label']}: {len(dropped)}개 항목 뺌(되살리기 가능)")
    agg["notes"] = (agg.get("notes") or []) + notes
    agg["compressed"] = now()
    after = render.estimate(agg_doc(agg), GL, tpl)
    return {"agg": agg, "before": est, "after": after, "notes": notes}


# ── 문서 만들기 ──────────────────────────────────────────────────────────
def doc_opts(opts):
    """표기 옵션: 약어 표기(paren|note|off), 분야별 풀이 수준, 설명을 표 아래로, 첫 등장에만, 범례"""
    lv = {k: v for k, v in (opts.get("levels") or {}).items() if k in rules.FIELDS}
    return {"expand_mode": opts.get("expand_mode") or "note", "levels": {**rules.LEVELS, **lv}, "desc_block": bool(opts.get("desc_block")),
            "first_only": opts.get("first_only", True) is not False, "legend": opts.get("legend", True)}


def agg_doc(agg, opts=None):
    opts = opts or {}
    wk = rules.week_of(agg.get("week") or None)
    return {"title": opts.get("title") or f"{agg.get('org', '')} 주간업무 보고", "subtitle": opts.get("subtitle") or f"{wk['label']}",
            "periods": {"done": wk["label"], "plan": rules.plan_period(wk["key"])}, "week": wk["key"], "org": agg.get("org", ""),
            "remarks": agg.get("remarks") or [],
            "project_order": {r["label"]: [p["name"] for p in REG.active(agg.get("org", ""), r["label"])] for r in agg.get("rows") or []},
            "rows": [{"label": r["label"], "items": r.get("items") or [], "people": r.get("people") or []} for r in agg.get("rows") or []],
            "queries": agg.get("queries") or {},
            **doc_opts(opts)}


def personal_doc(rep, opts=None):
    opts = opts or {}
    wk = rules.week_of(rep.get("week") or None)
    return {"title": opts.get("title") or f"{rep.get('dept', '')} 주간업무 보고", "subtitle": opts.get("subtitle") or f"{wk['label']} | 작성: {rep.get('name', '')}",
            "periods": {"done": wk["label"], "plan": rules.plan_period(wk["key"])}, "week": wk["key"], "org": rep.get("org") or rep.get("dept", ""),
            "remarks": rep.get("remarks") or [],
            "project_order": {rep.get("dept", ""): [p["name"] for p in REG.active(rep.get("org", ""), rep.get("dept", ""))]},
            "rows": [{"label": rep.get("dept", ""), "sub": rep.get("name", ""), "items": rep.get("items") or []}],
            **doc_opts(opts)}


def make_doc(req):
    opts = req.get("options") or {}
    if req.get("agg"):
        return agg_doc(req["agg"], opts)
    return personal_doc(dict(req.get("report") or {}, items=[norm_item(x) for x in (req.get("report") or {}).get("items") or []]), opts)


def kordoc_check(data):
    """kordoc 이 있으면 validate + 다시 읽기(parse) 로 확인 — 없으면 None"""
    if not os.path.exists(KORDOC):
        return None
    import tempfile
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "w.hwpx")
        with open(p, "wb") as f:
            f.write(data)
        try:
            v = subprocess.run(["node", KORDOC, "validate", p], capture_output=True, text=True, timeout=60)
            t = subprocess.run(["node", KORDOC, p, "--silent"], capture_output=True, text=True, timeout=60)
        except (OSError, subprocess.TimeoutExpired):
            return None
        return {"valid": v.returncode == 0, "msg": (v.stdout + v.stderr).strip()[-300:], "text": t.stdout[:20000]}


XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def embed_of(doc, bbs=False):
    groups = []
    for r in doc.get("rows") or []:
        its = []
        for blk in rules.blocks(r.get("items") or []):
            if bbs and blk[0].get("nobbs"):
                continue
            its += [{k: v for k, v in i.items() if k not in ("warn", "src")} for i in blk if not (bbs and i.get("nobbs"))]
        groups.append({"dept": r.get("label", ""), "name": r.get("sub", ""), "items": its})
    return exchange.payload(groups, doc.get("week"), doc.get("org", ""), doc.get("remarks") or [], "bbs" if bbs else "report")


def registry_view(org, dept=None):
    """양식 만들기 화면: 소본부의 실 목록(설정의 소본부·부서 + 등록부) 과 실마다 과제·구성원·지운 과제"""
    depts = list(orgs().get(org) or [])
    for u in REG.units(org):
        if u["dept"] not in depts:
            depts.append(u["dept"])
    units = []
    for d in ([dept] if dept else depts):
        u = REG.unit(org, d)
        units.append({"org": org, "dept": d, "members": u.get("members") or [],
                      "projects": sorted([p for p in u["projects"] if not p.get("deleted")], key=lambda p: p.get("order", 0)),
                      "deleted": [p for p in u["projects"] if p.get("deleted")]})
    return {"org": org, "depts": depts, "units": units, "orgs": sorted(set(orgs()) | {u["org"] for u in REG.units() if u["org"]})}


def registry_op(req):
    if req.get("op") == "copy_prev":  # 지난주(최근 4주) 제출에서 쓴 과제명을 등록부에 더한다(이미 있는 이름은 건너뜀)
        wk = rules.week_of(req.get("week") or None)
        names = _used_projects({"org": req.get("org"), "dept": req.get("dept")}, wk["key"], 4)
        if not names:
            raise ValueError("최근 제출에서 쓴 과제명이 없습니다")
        req = dict(req, op="add", names=names, source="지난 제출에서 가져옴")
    REG.op(req)
    return registry_view(req.get("org", ""), None)


def registry_template(req):
    """등록부로 채운 작성 양식 — scope: org(소본부 전체) | dept(한 실) / per_person: 사람마다 한 블록·한 장 / prefill: 엑셀에 사람·과제마다 한 줄"""
    org, dept, fmt = (req.get("org") or "").strip(), (req.get("dept") or "").strip(), req.get("fmt") or "hwpx"
    view = registry_view(org, dept or None)
    units = [{"dept": u["dept"], "members": [m for m in u["members"] if not req.get("person") or m == req["person"]] or ([req["person"]] if req.get("person") else []),
              "projects": [p for p in u["projects"] if p.get("active", True)]} for u in view["units"]]
    per_person = bool(req.get("per_person") or req.get("person"))
    doc = exchange.registry_doc(units, req.get("week"), org, per_person)
    tpl = render.template(req.get("template") or None)
    if req.get("preview"):
        return {"html": render.to_html(doc, GL, tpl), "rows": len(doc["rows"])}
    rows = []
    if req.get("prefill"):
        for u in units:
            for m in (u["members"] or [""]):
                rows += [{"dept": u["dept"], "name": m, "project": p["name"]} for p in u["projects"] if not p.get("person") or p["person"] == m]
    names = list(dict.fromkeys(p["name"] for u in units for p in u["projects"]))
    data = exchange.template_from(doc, fmt, GL, tpl, rows, names or None)
    ct = {"xlsx": XLSX_CT, "hwpx": "application/hwp+zip", "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}[fmt]
    return data, ct, f"주간보고_작성양식_{dept or org or '실'}.{fmt}"


def import_files(req):
    """여러 파일(형식 섞여도 됨) → 파일별 해석 결과 + 행별 점검 + 모르는 약어(질의 흐름으로)"""
    week = rules.week_of(req.get("week") or None)
    out = []
    for name, data in exchange.decode_files(req.get("files")):
        if len(data) > 30 * 1024 * 1024:
            out.append({"name": name, "method": "", "rows": [], "notes": ["30MB 넘는 파일은 받지 않음"]})
            continue
        res = exchange.import_file(name, data, week["key"], KORDOC)
        for r in res["rows"]:
            r["item"] = norm_item(dict(r["item"], _year=datetime.date.fromisoformat(week["mon"]).year))
        out.append(res)
    for f in out:  # 등록 과제와 비슷한 이름(띄어쓰기·괄호·대소문자만 다르거나 아주 비슷) → 제안만, 바꾸지 않음
        for r in f["rows"]:
            pj = r["item"].get("project") or ""
            cands = [p["name"] for u in REG.units() if not r.get("dept") or u["dept"] == r["dept"] for p in REG.active(u["org"], u["dept"])]
            sug = projects.near(pj, cands) if pj else None
            if sug:
                r["suggest_project"] = sug
                r["warnings"].append(f"과제명 '{pj}' — 등록 과제 '{sug}' 와 같은 과제인지 확인(제안만, 자동으로 바꾸지 않음)")
    items = [r["item"] for f in out for r in f["rows"]]
    asks = [w for w in rules.abbr_warnings(items, GL) if w["level"] == "ask"]
    known = {a: GL.resolve(a)[0] for a in {x["abbr"] for f in out for x in f.get("abbrs") or []}}
    for f in out:  # 문서 안 '약어: 풀이' 줄 — 약어집에 없으면 저장 후보
        f["abbrs"] = [dict(x, known=bool(known.get(x["abbr"]))) for x in f.get("abbrs") or []]
    return {"files": out, "asks": asks, "week": week}


def export(req):
    doc = make_doc(req)
    tpl = render.template((req.get("options") or {}).get("template"))
    fmt, bbs = req.get("format") or "html", bool(req.get("bbs"))
    name = re.sub(r"[\\/:*?\"<>|\s]+", "_", doc["title"]) + ("_BBS게시용" if bbs else "")
    emb = embed_of(doc, bbs)  # 다시 올리면 손실 없이 돌아오게 항목 JSON 을 넣는다(BBS 게시용은 비게시 항목을 빼고)
    if fmt == "hwpx":
        return exchange.embed_zip(render.to_hwpx(doc, GL, tpl, bbs), "hwpx", emb), "application/hwp+zip", name + ".hwpx"
    if fmt == "docx":
        return exchange.embed_zip(render.to_docx(doc, GL, tpl, bbs), "docx", emb), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name + ".docx"
    if fmt == "pptx":  # 실마다 한 장, 수행·계획 글상자, 목록 수준 = 과제명/항목/하위
        return exchange.report_pptx(doc, GL, tpl, bbs, emb), "application/vnd.openxmlformats-officedocument.presentationml.presentation", name + ".pptx"
    if fmt == "xlsx":  # 보고서형(공식 양식 배치)
        return exchange.report_xlsx(doc, GL, tpl, bbs, emb), XLSX_CT, name + ".xlsx"
    if fmt == "xlsx-data":  # 데이터형(입력 양식) — 엑셀에서 고쳐 다시 올리기
        return exchange.data_xlsx(emb["groups"], doc.get("week"), doc.get("org", ""), emb["remarks"]), XLSX_CT, name + "_입력양식.xlsx"
    if fmt == "txt":
        return render.to_text(doc, GL, tpl, bbs, marks=not bbs).encode(), "text/plain; charset=utf-8", name + ".txt"
    return render.to_html(doc, GL, tpl, bbs, full=True).encode(), "text/html; charset=utf-8", name + ".html"


def preview(req):
    doc = make_doc(req)
    tpl = render.template((req.get("options") or {}).get("template"))
    est = render.estimate(doc, GL, tpl, bool(req.get("bbs")))
    return {"html": render.to_html(doc, GL, tpl, bool(req.get("bbs"))), "checks": rules.check_doc(doc, GL, tpl["page"], est),
            "estimate": est, "text": render.to_text(doc, GL, tpl, bool(req.get("bbs")))}


# ── 내용 수집 보조: 같은 서버 다른 도구의 기록 (읽기 전용) ───────────────────
def sources(week_key):
    wk = rules.week_of(week_key)
    lo, hi = wk["mon"], (datetime.date.fromisoformat(wk["mon"]) + datetime.timedelta(days=6)).isoformat()
    out = {"meetings": [], "mails": []}
    d = os.path.join(DATA, "meeting-local")
    if os.path.isdir(d):
        for name in sorted(os.listdir(d), reverse=True):
            st = rules.load_json(os.path.join(d, name, "state.json"), None)
            md = os.path.join(d, name, "minutes.md")
            if not st or not os.path.exists(md) or not lo <= (st.get("ts") or "")[:10] <= hi:
                continue
            text = read(md)
            title = next((l.lstrip("# ").strip() for l in text.splitlines() if l.startswith("#")), name)
            out["meetings"].append({"id": name, "title": title, "ts": st.get("ts"), "text": text[:4000]})
    d = os.path.join(DATA, "mail-local", "history")
    if os.path.isdir(d):
        for fn in sorted(os.listdir(d), reverse=True):
            r = rules.load_json(os.path.join(d, fn), None) if fn.endswith(".json") else None
            if not r or not lo <= (r.get("ts") or "")[:10] <= hi or r.get("kind") not in ("write", "reply"):
                continue
            v = (r.get("versions") or [{}])[-1]
            body = (v.get("body") or "")[:1500]
            out["mails"].append({"id": r["id"], "title": r.get("title") or "", "ts": r.get("ts"), "kind": r["kind"],
                                 "text": f"[메일 {'작성' if r['kind'] == 'write' else '회신'}] 제목: {r.get('title', '')}\n{body}"})
    return out


def meta():
    return {"model": MODEL, "week": rules.week_of(), "cats": rules.CATS, "kinds": rules.KINDS, "guide": GUIDE, "orgs": orgs(),
            "templates": render.templates(), "glossary": GL.stats(), "fields": rules.FIELDS, "levels": rules.LEVELS, "presets": rules.PRESETS, "colors": render.template()["colors"],
            "has_sources": any(os.path.isdir(os.path.join(DATA, t)) for t in ("meeting-local", "mail-local")), "kordoc": os.path.exists(KORDOC)}


# ── HTTP ───────────────────────────────────────────────────────────────
HTML = read(os.path.join(ROOT, "ui.html")) if os.path.exists(os.path.join(ROOT, "ui.html")) else "ui.html 없음"

# ── 저작권 표기 (LICENSE·NOTICE 참고) ─────────────────────────────────────
_SIG = __import__("base64").b64decode("wqkgMjAyNiBnZ2dnODY1NyDCtyBkb25nanVraW0uZGV2QGdtYWlsLmNvbQ==").decode()
_SIG_A = __import__("base64").b64decode("Z2dnZzg2NTcgPGRvbmdqdWtpbS5kZXZAZ21haWwuY29tPg==").decode()


def signed(html):
    """화면에 저작권 표기를 붙인다. ui.html 에서 지워져도 서버가 내보낼 때 다시 붙는다."""
    name, mail = _SIG.split(" · ")
    if 'name="author"' not in html:
        meta_ = f'<meta name="author" content="{name[7:]} <{mail}>">'
        html = html.replace("<head>", "<head>" + meta_, 1) if "<head>" in html else meta_ + html
    if "data-sig" not in html:
        tag = (f'<!-- {_SIG} --><div data-sig title="{mail}" style="text-align:center;font-size:11px;color:#9aa0a6;'
               f'opacity:.55;margin:28px 0 8px">{name}</div>')
        html = html.replace("</body>", tag + "</body>", 1) if "</body>" in html else html + tag
    return html


STREAM = {"/api/itemize": run_itemize, "/api/aggregate": run_aggregate, "/api/compress": run_compress}


class H(BaseHTTPRequestHandler):
    def log_message(self, fmt, *a):
        if any(p in (a[0] if a else "") for p in STREAM):
            super().log_message(fmt, *a)

    def _send(self, body, ctype="application/json", code=200, name=None, extra=None):
        b = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("X-Author", _SIG_A)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(b)))
        self.send_header("Cache-Control", "no-store")
        if name:
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + urllib.parse.quote(name))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        q = {k: v[0] for k, v in urllib.parse.parse_qs(u.query).items()}
        try:
            if u.path == "/api/health":
                return self._send({"ok": True})
            if u.path == "/api/meta":
                return self._send(meta())
            if u.path == "/api/models":
                return self._send(models())
            if u.path == "/api/week":
                return self._send(rules.week_of(q.get("d") or None))
            if u.path == "/api/prev":
                prof = {k: q.get(k, "") for k in ("org", "dept", "name")}
                return self._send(prev_report(prof, rules.week_of(q.get("week") or None)["key"]) or {})
            if u.path == "/api/registry":
                return self._send(registry_view(q.get("org", ""), q.get("dept") or None))
            if u.path == "/api/form-template":  # 빈 작성 양식 — xlsx | hwpx | docx | pptx
                fmt = q.get("fmt", "xlsx")
                prof = {k: q.get(k, "") for k in ("org", "dept", "name")}
                depts = [prof["dept"]] if prof["dept"] else (orgs().get(prof["org"]) or [])
                tpl = render.template(q.get("template") or None)
                data = exchange.form_template(fmt, depts, q.get("week") or None, GL, tpl, prof["org"], prof["name"], known_projects(prof, q.get("week") or None) or None)
                ct = {"xlsx": XLSX_CT, "hwpx": "application/hwp+zip", "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                      "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation"}[fmt]
                return self._send(data, ct, name=f"주간보고_작성양식.{fmt}")
            if u.path == "/api/xlsx-template":
                prof = {k: q.get(k, "") for k in ("org", "dept", "name")}
                data = exchange.input_template(known_projects(prof, q.get("week") or None) or None, prof["dept"], prof["name"], q.get("week") or None)
                return self._send(data, XLSX_CT, name="주간보고_입력양식.xlsx")
            if u.path == "/api/projects":
                return self._send(known_projects({k: q.get(k, "") for k in ("org", "dept", "name")}, q.get("week") or None))
            if u.path == "/api/report":
                prof = {k: q.get(k, "") for k in ("org", "dept", "name")}
                p = os.path.join(WS, "reports", _wk(q.get("week")), report_id(prof) + ".json")
                return self._send(rules.load_json(p, {}))
            if u.path == "/api/reports":
                return self._send(list_reports(q.get("week"), q.get("org")))
            if u.path == "/api/status":
                return self._send(status(q.get("week"), q.get("org", "")))
            if u.path == "/api/sources":
                return self._send(sources(q.get("week") or None))
            if u.path == "/api/glossary":
                return self._send({"rows": GL.search(q.get("q", ""), field=q.get("field", "")), "stats": GL.stats(), "ignored": sorted(rules.load_json(GL.ignore_path, []))})
            self._send(signed(HTML).encode(), "text/html; charset=utf-8")
        except (FileNotFoundError, ValueError) as e:
            self._send({"error": str(e) or "없음"}, code=404)
        except Exception as e:
            self._send({"error": f"{type(e).__name__}: {e}"}, code=500)

    def do_POST(self):
        try:
            req = json.loads(self.rfile.read(int(self.headers.get("Content-Length") or 0)) or b"{}")
        except ValueError:
            return self._send({"error": "잘못된 요청"}, code=400)
        path = self.path.split("?")[0]
        if path in STREAM:
            return self._stream(STREAM[path], req)
        try:
            if path == "/api/submit":
                try:
                    return self._send(submit(req))
                except Unresolved as e:
                    return self._send({"error": str(e), "unresolved": e.rows}, code=409)
            if path == "/api/preview":
                return self._send(preview(req))
            if path == "/api/export":
                data, ctype, name = export(req)
                return self._send(data, ctype, name=name, extra={"X-Unresolved": str(len(rules.unresolved(
                    [i for r in make_doc(req)["rows"] for i in r["items"]], GL)))})
            if path == "/api/registry/op":
                return self._send(registry_op(req))
            if path == "/api/registry/template":
                res = registry_template(req)
                if isinstance(res, dict):
                    return self._send(res)
                return self._send(res[0], res[1], name=res[2])
            if path == "/api/import":
                return self._send(import_files(req))
            if path == "/api/verify-hwpx":  # 내보낸 HWPX 를 XML·kordoc 으로 다시 읽어 확인
                data, _, _ = export(dict(req, format="hwpx"))
                return self._send({"inspect": {k: v for k, v in render.inspect_hwpx(data).items() if k != "texts"}, "kordoc": kordoc_check(data)})
            if path == "/api/check":
                doc = make_doc(req)
                tpl = render.template((req.get("options") or {}).get("template"))
                est = render.estimate(doc, GL, tpl)
                return self._send({"checks": rules.check_doc(doc, GL, tpl["page"], est), "estimate": est})
            if path == "/api/dates":  # 날짜 표기 통일
                year = datetime.date.fromisoformat(rules.week_of(req.get("week") or None)["mon"]).year
                return self._send({"items": [dict(it, text=rules.normalize_dates(it.get("text") or "", year)) for it in req.get("items") or []]})
            if path == "/api/agg/save":
                agg = req.get("agg") or {}
                agg["ts"] = now()
                rules.save_json(agg_path(agg.get("week"), agg.get("org", "")), agg)
                return self._send({"ok": True})
            if path == "/api/orgs":
                return self._send(save_orgs(req.get("orgs")))
            if path == "/api/glossary/save":
                return self._send(GL.put(req.get("abbr", ""), req.get("full", ""), req.get("ko", ""), req.get("note", ""), req.get("editor", ""),
                                         req.get("src") or "사내", req.get("desc", ""), req.get("field", ""), req.get("ctx", "")))
            if path == "/api/glossary/delete":
                GL.delete(req.get("abbr", ""), req.get("editor", ""))
                return self._send({"ok": True})
            if path == "/api/glossary/ignore":
                GL.ignore(req.get("abbr", ""), req.get("on", True))
                return self._send({"ok": True})
            return self._send({"error": "없는 경로"}, code=404)
        except ValueError as e:
            return self._send({"error": str(e)}, code=400)
        except Exception as e:
            return self._send({"error": f"{type(e).__name__}: {e}"}, code=500)

    def _stream(self, fn, req):
        self.send_response(200)
        self.send_header("X-Author", _SIG_A)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()

        def emit(ev):
            self.wfile.write(f"data: {json.dumps(ev, ensure_ascii=False)}\n\n".encode())
            self.wfile.flush()

        try:
            emit({"done": fn(req, emit, req.get("model") or MODEL)})
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as e:
            try:
                emit({"error": str(e) if isinstance(e, (ValueError, RuntimeError)) else f"{type(e).__name__}: {e}"})
            except OSError:
                pass


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        tok = lambda ev: print(ev["token"], end="", file=sys.stderr, flush=True) if "token" in ev else None
        r = run_itemize({"memo": sys.stdin.read(), "use_prev": False}, tok, MODEL)
        print()
        for it in r["items"]:
            print(f"[{it['kind']}/{it['cat']}] {it['text']}{rules.ext_suffix(it)}" + (f"  ⚠ {' / '.join(it['warn'])}" if it.get("warn") else ""))
        sys.exit(0)
    print(f"weekly local → http://localhost:{PORT}  (llm={LLM_API} {LLM_BASE} {MODEL}, workspace={WS})  {_SIG}")
    ThreadingHTTPServer(("0.0.0.0", PORT), H).serve_forever()

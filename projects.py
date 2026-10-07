"""과제 등록부 — WORKSPACE/projects.json (실 전체가 같이 씀, 로그인 없음: 고친 사람·때·이력만 남김)

구조: {"units": {"소본부|실": {"org", "dept", "members": [이름], "projects": [과제]}}}
과제: {"id", "name"(보고서에 쓰는 이름), "full"(정식 이름·코드, 선택), "period"(기본 기간, 선택), "person"(""=실 전체), "active", "order",
       "deleted", "editor", "ts", "history": [{"ts", "editor", "action", "before"}]}
"""
import datetime
import difflib
import os
import re
import secrets

import rules

KEYS = ("name", "full", "period", "person", "active", "aliases")


def _now():
    return datetime.datetime.now().isoformat(timespec="seconds")


class Registry:
    def __init__(self, ws):
        self.path = os.path.join(ws, "projects.json")

    def load(self):
        return rules.load_json(self.path, {"units": {}})

    def save(self, d):
        rules.save_json(self.path, d)

    @staticmethod
    def key(org, dept):
        return f"{(org or '').strip()}|{(dept or '').strip()}"

    def unit(self, org, dept, d=None):
        d = d or self.load()
        u = d["units"].get(self.key(org, dept))
        return u or {"org": (org or "").strip(), "dept": (dept or "").strip(), "members": [], "projects": []}

    def units(self, org=None):
        d = self.load()
        return [u for u in d["units"].values() if org is None or u["org"] == (org or "").strip()]

    def active(self, org, dept, person=None):
        """보고서에 쓸 과제 이름(등록 순서) — person 을 주면 그 사람 과제 + 실 전체 과제"""
        ps = [p for p in self.unit(org, dept)["projects"] if p.get("active", True) and not p.get("deleted")]
        if person is not None:
            ps = [p for p in ps if not p.get("person") or p["person"] == person]
        return sorted(ps, key=lambda p: p.get("order", 0))

    def units_for(self, org, dept):
        """소본부를 모르면 같은 실 이름의 등록부"""
        if org:
            return [self.unit(org, dept)]
        return [u for u in self.units() if u["dept"] == dept] or [self.unit("", dept)]

    def alias_map(self, org, dept):
        """{정규화한 이름·별칭: 등록 이름}"""
        out = {}
        for u in self.units_for(org, dept):
            for p in u["projects"]:
                if p.get("deleted"):
                    continue
                out.setdefault(norm_name(p["name"]), p["name"])
                for a in p.get("aliases") or []:
                    out.setdefault(norm_name(a), p["name"])
        return out

    def canonical(self, name, org, dept):
        """→ (등록 이름, '별칭') 또는 (None, '') — 등록 이름과 같으면 (name, '')"""
        if not name:
            return None, ""
        for u in self.units_for(org, dept):
            for p in u["projects"]:
                if p.get("deleted"):
                    continue
                if p["name"] == name:
                    return name, ""
                if any(norm_name(name) == norm_name(a) for a in p.get("aliases") or []):
                    return p["name"], "별칭"  # 등록한 별칭만 자동으로(띄어쓰기·괄호 차이는 무시)
        return None, ""  # 등록 이름과 표기만 다른 경우는 바꾸지 않고 제안(near)

    def names(self, org, dept=None):
        if dept:
            return [p["name"] for p in self.active(org, dept)]
        out = []
        for u in self.units(org):
            out += [p["name"] for p in self.active(u["org"], u["dept"]) if p["name"] not in out]
        return out

    def op(self, req):
        """add · update · delete · restore · reorder · members · copy(지난 제출에서 가져온 이름들)"""
        op, org, dept, editor = req.get("op"), (req.get("org") or "").strip(), (req.get("dept") or "").strip(), (req.get("editor") or "").strip()
        if not dept:
            raise ValueError("실(부서)을 고르세요")
        d = self.load()
        k = self.key(org, dept)
        u = d["units"].setdefault(k, {"org": org, "dept": dept, "members": [], "projects": []})
        ps = u["projects"]
        find = lambda pid: next((p for p in ps if p["id"] == pid), None)
        hist = lambda p, action: p.setdefault("history", []).append({"ts": _now(), "editor": editor, "action": action,
                                                                    "before": {x: p.get(x) for x in KEYS + ("deleted",)}})
        clean = lambda s: re.sub(r"\s+", " ", str(s or "")).strip().strip("()[]（）").strip()
        if op == "alias":  # 앞으로도 이 이름은 합치기: alias → name(등록 과제, 없으면 새로 등록)
            nm, al = clean(req.get("name")), clean(req.get("alias"))
            if not nm or not al or nm == al:
                raise ValueError("과제 이름과 별칭을 적어 주세요")
            p = next((p for p in ps if p["name"] == nm and not p.get("deleted")), None)
            if not p:
                p = {"id": secrets.token_hex(4), "name": nm, "full": "", "period": "", "person": "", "active": True, "aliases": [],
                     "order": max([x.get("order", 0) for x in ps] or [0]) + 1, "deleted": False, "editor": editor, "ts": _now(), "history": []}
                ps.append(p)
            if norm_name(al) not in [norm_name(a) for a in p.get("aliases") or []]:
                hist(p, f"별칭 추가: {al}")
                p.setdefault("aliases", []).append(al)
                p.update(editor=editor, ts=_now())
            for q in ps:  # 별칭이 따로 등록된 과제였으면 그 과제는 지운다(되살리기 가능)
                if q is not p and q["name"] == al and not q.get("deleted"):
                    hist(q, f"'{nm}' 의 별칭으로 합침")
                    q.update(deleted=True, editor=editor, ts=_now())
            self.save(d)
            return u
        if op == "add":
            names = req.get("names") or [req.get("name")]
            for nm in names:
                nm = clean(nm)
                if not nm:
                    continue
                if any(p["name"] == nm and not p.get("deleted") for p in ps):
                    continue
                ps.append({"id": secrets.token_hex(4), "name": nm, "full": clean(req.get("full")) if len(names) == 1 else "", "period": rules.norm_period(req.get("period") or "") if len(names) == 1 else "",
                           "person": (req.get("person") or "").strip() if len(names) == 1 else "", "active": True,
                           "order": max([p.get("order", 0) for p in ps] or [0]) + 1, "deleted": False, "editor": editor, "ts": _now(),
                           "history": [{"ts": _now(), "editor": editor, "action": "추가" if op == "add" and not req.get("source") else req.get("source")}]})
        elif op == "update":
            p = find(req.get("id"))
            if not p:
                raise ValueError("과제를 찾지 못함")
            ch = {}
            for x in KEYS:
                if x in req:
                    if x == "aliases":
                        v = [clean(a) for a in (req[x] if isinstance(req[x], list) else str(req[x] or "").split(",")) if clean(a)]
                    else:
                        v = bool(req[x]) if x == "active" else (rules.norm_period(req[x]) if x == "period" else (clean(req[x]) if x == "name" else str(req[x] or "").strip()))
                    if v != p.get(x):
                        ch[x] = v
            if "name" in ch and not ch["name"]:
                raise ValueError("과제 이름이 비었습니다")
            if ch:
                hist(p, "이름 바꿈" if "name" in ch else "수정")
                p.update(ch, editor=editor, ts=_now())
        elif op in ("delete", "restore"):
            p = find(req.get("id"))
            if not p:
                raise ValueError("과제를 찾지 못함")
            hist(p, "삭제" if op == "delete" else "되살림")
            p.update(deleted=op == "delete", editor=editor, ts=_now())
        elif op == "reorder":
            ids = req.get("ids") or []
            for n, pid in enumerate(ids):
                p = find(pid)
                if p:
                    p["order"] = n + 1
        elif op == "members":
            u["members"] = [m.strip() for m in (req.get("members") if isinstance(req.get("members"), list) else str(req.get("members") or "").replace("\n", ",").split(",")) if m.strip()]
        else:
            raise ValueError("op 은 add|update|delete|restore|reorder|members|alias")
        self.save(d)
        return u


def norm_name(s):
    return re.sub(r"[\s()\[\]（）·_\-]", "", str(s or "")).lower()


LETTERS = {"에이치": "H", "더블유": "W", "에이": "A", "비": "B", "씨": "C", "시": "C", "디": "D", "이": "E", "에프": "F", "지": "G", "쥐": "G",
           "아이": "I", "제이": "J", "케이": "K", "엘": "L", "엠": "M", "엔": "N", "오": "O", "피": "P", "큐": "Q", "알": "R", "아르": "R",
           "에스": "S", "티": "T", "유": "U", "브이": "V", "엑스": "X", "와이": "Y", "제트": "Z", "지드": "Z"}


def _spell(run):
    """한글로 읽은 영문 글자 이름 → 글자('아이에스엠알' → 'ISMR'). 끝까지 글자 이름으로 읽히지 않으면 None"""
    best = [None] * (len(run) + 1)
    best[0] = ""
    for i in range(len(run)):
        if best[i] is None:
            continue
        for k, v in LETTERS.items():
            if run.startswith(k, i) and best[i + len(k)] is None:
                best[i + len(k)] = best[i] + v
    return best[-1]


def translit(s):
    """과제명 비교용: 한글로 적은 영문 약자를 영문으로('아이에스엠알 과제' → 'ISMR 과제' ≈ 'i-SMR')"""
    def sub(m):
        r = _spell(m.group(0))
        return r if r and len(r) >= 2 else m.group(0)
    return re.sub(r"[가-힣]{2,}", sub, str(s or ""))


def near(name, candidates):
    """가져온 과제명이 등록 과제와 비슷하면 그 이름(띄어쓰기·괄호·대소문자 무시 같음 → 확실, 아니면 비슷한 정도 0.8 이상) — 바꾸지는 않고 제안만"""
    n = norm_name(name)
    if not n or name in candidates:
        return None
    for c in candidates:
        if norm_name(c) == n:
            return c
    best = max(candidates, key=lambda c: difflib.SequenceMatcher(None, n, norm_name(c)).ratio(), default=None)
    if best and difflib.SequenceMatcher(None, n, norm_name(best)).ratio() >= 0.8:
        return best
    return None

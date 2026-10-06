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

KEYS = ("name", "full", "period", "person", "active")


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
            raise ValueError("op 은 add|update|delete|restore|reorder|members")
        self.save(d)
        return u


def norm_name(s):
    return re.sub(r"[\s()\[\]（）·_\-]", "", str(s or "")).lower()


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

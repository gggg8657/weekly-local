#!/usr/bin/env python3
"""예전 제출을 포털 계정에 연결(한 번만 돌리는 정리 도구). 실행 전에 WORKSPACE 를 따로 백업할 것.

  WORKSPACE=<weekly 데이터 폴더> PORTAL_DB=<portal.db> python3 link_accounts.py [--dry-run]

- 이미 만든 계정(_by)이 기록된 제출은 그대로 둔다.
- 없으면 제출의 이름·부서가 포털 계정(표시 이름·승인 부서)과 정확히 한 건 일치할 때 그 계정에 연결한다.
- 그 밖(일치 없음·여러 건)은 '미연결'로 표시만 한다 — 주인이 없으므로 관리자만 고치거나 취소할 수 있다.
결과는 개수만 출력한다(보고 내용·이름은 출력하지 않음).
"""
import os
import sqlite3
import sys
import urllib.parse

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import app  # noqa: E402
import rules  # noqa: E402


def main(dry=False):
    c = sqlite3.connect(f"file:{urllib.parse.quote(app.PORTAL_DB)}?mode=ro", uri=True)
    users = [(u, (n or "").strip(), (d or "").strip()) for u, n, d, x in c.execute("SELECT id, name, dept, disabled FROM users") if not x]
    c.close()
    n = {"already": 0, "linked": 0, "unlinked": 0}
    root = os.path.join(app.WS, "reports")
    for wk in sorted(os.listdir(root)) if os.path.isdir(root) else []:
        for fn in sorted(os.listdir(os.path.join(root, wk))):
            if not fn.endswith(".json"):
                continue
            p = os.path.join(root, wk, fn)
            rec = rules.load_json(p, None)
            if not isinstance(rec, dict):
                continue
            if rec.get("_by"):
                n["already"] += 1
                continue
            hit = [u for u, nm, d in users if nm and d and nm == (rec.get("name") or "").strip() and d == (rec.get("dept") or "").strip()]
            if len(hit) == 1:
                rec["_by"], rec["_link"] = hit[0], "이름·부서 일치로 연결 " + app.now()
                n["linked"] += 1
            else:
                rec["_link"] = "미연결"
                n["unlinked"] += 1
            if not dry:
                rules.save_json(p, rec)
    print(("(시험) " if dry else "") + f"이미 계정 있음 {n['already']} · 연결 {n['linked']} · 미연결 {n['unlinked']}")


if __name__ == "__main__":
    main("--dry-run" in sys.argv)

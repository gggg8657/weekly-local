#!/usr/bin/env python3
"""NUREG-0544 Rev.4 HTML(웹 아카이브 사본) → nureg0544.json  [{"abbr", "full", "src"}]

  python3 glossary/build_nureg.py      # source/NUREG-0544r4_wayback20150219.html 을 다시 파싱

원본: U.S. NRC, NUREG-0544 Rev.4 "Collection of Abbreviations" (1998) — 미국 연방정부 저작물(public domain).
nrc.gov 는 자동 내려받기를 막아(403) Internet Archive 사본을 받았다. URL·해시는 NOTICE 참고.
'A' 장 머리(<h2 id="03">) 앞의 머리말 예시(<dd>DBAs for …)는 건너뛴다. 한 약어에 풀이가 여럿이면 각각 한 항목.
"""
import html
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "source", "NUREG-0544r4_wayback20150219.html")
OUT = os.path.join(HERE, "nureg0544.json")
TAG = "NUREG-0544 Rev.4"


def text(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def build():
    s = open(SRC, encoding="utf-8", errors="replace").read()
    s = s[s.index('<h2 id="03">'):]  # 머리말 예시 제외 — 본문 'A' 부터
    s = s[:s.index("</dl>", s.rindex("<dt"))]
    out, cur = [], None
    for tag, body in re.findall(r"<(dt|dd)[^>]*>(.*?)</\1>", s, re.S | re.I):
        t = text(body)
        if not t:
            continue
        if tag.lower() == "dt":
            cur = t
        elif cur:
            out.append({"abbr": cur, "full": t, "src": TAG})
    return out


if __name__ == "__main__":
    rows = build()
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=0)
    print(f"{len(rows)} 항목, 약어 {len({r['abbr'] for r in rows})}개 → {OUT}")

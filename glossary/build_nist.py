#!/usr/bin/env python3
"""NIST CSRC Glossary 내보내기(glossary-export.zip) → nist_csrc.json  [{"abbr", "full", "src", "field"}]

  python3 glossary/build_nist.py

원본: https://csrc.nist.gov/csrc/media/glossary/glossary-export.zip (NIST Computer Security Resource Center Glossary,
NISTIR 7298 Rev.3 의 데이터베이스) — 미국 정부 저작물. 약어 → 전체 이름(abbrSyn)만 쓴다(정의문은 쓰지 않음).
용어가 약어 모양(대문자 2자 이상 등)이고 abbrSyn 이 있는 것만, 대소문자만 다른 같은 풀이는 하나로.
"""
import html
import json
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))
from rules import is_abbr  # noqa: E402

SRC = os.path.join(HERE, "source", "NIST-CSRC-glossary-export_20261006.zip")
OUT = os.path.join(HERE, "nist_csrc.json")
TAG = "NIST CSRC Glossary"


def build():
    with zipfile.ZipFile(SRC) as z:
        name = next(n for n in z.namelist() if n.endswith(".json"))
        data = json.loads(z.read(name).decode("utf-8-sig"))
    out = []
    for t in data["parentTerms"]:
        abbr = html.unescape(re.sub(r"<[^>]+>", "", t.get("term") or "")).strip()
        if not t.get("abbrSyn") or not re.fullmatch(r"[A-Za-z][A-Za-z0-9&/\-]{1,15}", abbr) or not is_abbr(abbr):
            continue
        seen = set()
        for s in t["abbrSyn"]:
            full = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", "", s.get("text") or ""))).strip()
            if full and full.lower() not in seen and full != abbr:
                seen.add(full.lower())
                out.append({"abbr": abbr, "full": full, "src": TAG, "field": "컴퓨터"})
    return out


if __name__ == "__main__":
    rows = build()
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=0)
    print(f"{len(rows)} 항목, 약어 {len({r['abbr'] for r in rows})}개 → {OUT}")

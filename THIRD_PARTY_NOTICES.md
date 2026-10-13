# 제3자 구성요소 고지 (THIRD_PARTY_NOTICES)

weekly-local 의 원저작물은 DongJu Kim(gggg8657)이 작성했습니다(`NOTICE`, `LICENSE`). 아래 제3자 구성요소는 각자의 라이선스를 따르며, 이 저장소의 라이선스(PolyForm Noncommercial 1.0.0)로 바뀌지 않습니다.
조사일: 2026-10-12 (실제 파일·import·setup.sh 기준). "상업 이용" 칸은 확인한 라이선스 문구를 요약한 것이며 법률 판단이 아닙니다.

코드는 파이썬 표준 라이브러리만 씁니다(XLSX·PPTX·DOCX·HWPX 는 `zipfile` + XML 로 직접 작성). 제3자 **데이터**를 일부 동봉합니다(아래 표).

| 이름 | 버전 | 쓰임 | 라이선스 | 링크 | 상업 이용 | 비고 |
|---|---|---|---|---|---|---|
| NUREG-0544 Rev.4 "Collection of Abbreviations" (U.S. NRC, 1998) | Rev.4 | **동봉 데이터** `glossary/source/NUREG-0544r4_wayback20150219.html` → `glossary/nureg0544.json` | 미국 정부 저작물 (17 U.S.C. §105, 미국 내 저작권 없음) | [NRC 원 URL](http://www.nrc.gov/reading-rm/doc-collections/nuregs/staff/sr0544/) · [Internet Archive 사본](http://web.archive.org/web/20150219155523id_/http://www.nrc.gov/reading-rm/doc-collections/nuregs/staff/sr0544/) | 조건부 — [#1](https://github.com/gggg8657/weekly-local/issues/1) | 수신일·SHA-256 은 `NOTICE`. 내용 수정 없이 파싱 |
| NIST CSRC Glossary 내보내기 (NISTIR 7298 Rev.3) | 2026-10-06 수신 | **동봉 데이터** `glossary/source/NIST-CSRC-glossary-export_20261006.zip` → `glossary/nist_csrc.json` (약어·전체 이름만) | 미국 정부 저작물 (17 U.S.C. §105). 해외 권리는 NIST 가 출처 표시 조건으로 무상 허락 | [csrc.nist.gov/glossary](https://csrc.nist.gov/glossary) · [NIST 이용 조건](https://www.nist.gov/open/copyright-fair-use-and-licensing-statements-srd-data-software-and-technical-series-publications) | 조건부 — [#1](https://github.com/gggg8657/weekly-local/issues/1) | 출처: National Institute of Standards and Technology. 가공: 약어 모양 용어와 전체 이름만 추출, 정의문 제외 (2026-10-06) |
| kordoc | kordoc-local 설치본 | 외부 프로그램 (HWPX 검증·다시 읽기, HWP/PDF 가져오기, `node` 로 호출) | MIT (Copyright (c) 2026 chrisryugj) | [chrisryugj/kordoc](https://github.com/chrisryugj/kordoc) | 가능 | 동봉 안 함. `templates/default.hwpx` 는 kordoc `generate --preset 통지` 출력에 자리표시자를 넣은 문서 |
| Node.js | 시스템 설치본 | 외부 프로그램 (kordoc 실행) | MIT | [nodejs.org](https://github.com/nodejs/node/blob/main/LICENSE) | 가능 | |
| openpyxl / python-pptx | 설치돼 있을 때만 | `selftest.py` 의 선택 검사 (실행 코드는 쓰지 않음) | MIT / MIT | [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) · [python-pptx](https://github.com/scanny/python-pptx) | 가능 | `setup.sh` 가 설치하지 않음 |
| Python 표준 라이브러리 | 3.x | 실행 환경 | PSF-2.0 | [python.org](https://docs.python.org/3/license.html) | 가능 | |
| Ollama | — | 외부 서비스 (LLM 서버) | MIT | [ollama/ollama](https://github.com/ollama/ollama) | 가능 | 동봉 안 함 |
| Google Gemma 4 `gemma4:31b` | — | 외부 서비스가 돌리는 모델 (기본 배포) | Apache-2.0 | [google/gemma-4-31B-it](https://huggingface.co/google/gemma-4-31B-it) | 가능 | 가중치 동봉 안 함. `LLM_MODEL` 로 다른 모델 지정 가능 |

## 이 저작자가 만든 데이터

- `glossary/tech_seed.json` — AI·컴퓨터·일반 약어 66개, 이 저장소 저작자가 정리(출처 표시 "큐레이션(weekly-local)"). PolyForm Noncommercial 1.0.0.
- `glossary/kr_seed.json` — 국내 기관·원자력 약어 31개, 공개적으로 알려진 명칭을 이 저장소 저작자가 적은 것(출처 표시 "사내 시드(확인 필요)"). 사내 확인 항목은 [#1](https://github.com/gggg8657/weekly-local/issues/1).
- `templates/default.json` — 기본 양식 매핑. MIT.

## 저장소에 넣지 않은 것

- 기관 공식 양식 `format.hwpx`, `templates/kaeri_weekly.hwpx`, `templates/kaeri_weekly.json` — 기관 내부 양식으로 이 저장소의 라이선스 대상이 아니며 `.gitignore` 로 제외했습니다(git 이력에도 없음, 2026-10-12 확인). 쓰려면 받은 양식을 로컬 `templates/` 에 넣으세요.
- IAEA Nuclear Safety and Security Glossary, NIST AI RMF Glossary — 저작권 자료·제3자 정의 인용이 있어 넣지 않았습니다.

## 확인 필요

- [#1 약어집 데이터 NUREG-0544·NIST CSRC Glossary — 미국 정부 저작물, 해외 권리·원본 동봉](https://github.com/gggg8657/weekly-local/issues/1)
  - 미국 정부 저작물은 미국 안에서 저작권이 없지만 해외 권리는 별도. NIST 는 출처 표시 조건으로 전 세계 이용을 허락, NRC 는 미확인.
  - 원본 zip 에는 정의문 전체(일부 비-NIST 인용 포함)가 들어 있음 — 앱은 `glossary/*.json` 만 읽으므로 **원본 zip/html 을 저장소에서 빼고 `glossary/build_*.py` 로 사용자가 직접 받은 파일에서 다시 만들게 하는 방법**이 회피책(제안, 미적용).
  - `kr_seed.json` 이 공개 명칭만 담았는지 사내 확인.

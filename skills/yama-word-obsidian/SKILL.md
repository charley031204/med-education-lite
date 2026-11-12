---
name: yama-word-obsidian
description: 문항 전용 Yama Word를 강의록과 녹음정제 Script를 근거로 해설이 붙은 Obsidian 문항집으로 바꾼다. 단일 기출 문항 질문에는 yama-question-answer를 사용한다.
---

# Yama Word → Obsidian

문항 전용 Word가 문두·선지·연도·회차·번호·공통지문·그림의 원본이다. 강의록과 Script로 해설을 작성한다. Lite는 Word→Obsidian 정방향을 지원한다.

1. 사용자가 지정한 Word, 대응 강의록과 Script를 확인한다. 작업 폴더에 인덱스를 한 번 생성한다. `$WordSkill`과 `$QaSkill`은 현재 패키지에 포함된 각 스킬의 디렉터리다.
2. `list`로 최대 10개씩 ID를 열고 `lookup`으로 한 문항의 원문과 필요한 그림을 조회한다. Word 문항 수, 섹션, 공통지문과 그림을 확인한다. 문항이 누락되거나 필요한 지문·그림을 못 찾으면 해당 원본 위치를 먼저 확인한다.
3. [답안 원칙](references/answer-principles.md)을 읽는다. [단일 문항 스킬](../yama-question-answer/SKILL.md)의 로컬 자료 검색 명령으로 필요한 근거만 연다. [Obsidian 답안 형식](references/obsidian-format.md)에 따라 ID별로 해설을 작성한다. 강의록과 Script 모두에 필요한 설명이 없을 때 표준교과서로 해당 부분만 보충할 수 있다. 이 해설집에는 웹 검색 결과를 사용하지 않는다.
4. `emit`으로 Markdown과 실제 그림을 저장한다. `qc`로 저장 내용을 확인한다. 각 문항의 번호·선지·지문·그림 연결, 선지별 판단 이유와 근거를 원본에 대조한다. 형식 검증과 내용 검토 결과를 구분해서 보고한다.

```powershell
python "$WordSkill/scripts/question_index.py" build --source BASE.docx --out WORK/questions.json
python "$WordSkill/scripts/obsidian_bank.py" list --index WORK/questions.json --offset 0 --limit 5
python "$WordSkill/scripts/question_index.py" lookup --index WORK/questions.json --qid QID --figures WORK/figures
python "$WordSkill/scripts/obsidian_bank.py" emit --index WORK/questions.json --answers WORK/answers.json --out OUTPUT.md
python "$WordSkill/scripts/obsidian_bank.py" qc --index WORK/questions.json --answers WORK/answers.json --out OUTPUT.md
```

기본 배열은 Word 섹션 안에서 국내 시험의 최신 연도 내림차순이다. `--sort word`로 Word 배열을 선택할 수 있다. 최근 시험만 요청하면 `list`, `emit`, `qc`에 같은 `--min-year YY`를 적용한다. 기준은 국내 시험 연도다. 건양대 이력만 있는 문항은 범위에서 제외한다. 포함된 문항의 출제 이력은 모두 유지한다. Word에 공통사진 결정이 기록되어 있으면 그 연결을 사용한다. 기록이 없는 Word는 각 문항의 실제 그림 배치를 유지한다.

---
name: yama-question-answer
description: 사용자가 지정한 기출 한 문항의 정답이나 특정 선지 질문에 답한다. 강의록과 녹음정제 Script를 먼저 검색하고 공식 답안·교수 답변·자료 기반 풀이·불확실성을 구분한다. 전체 해설집 작성은 yama-word-obsidian을 사용한다.
---

# 기출 한 문항 답변

[공통 답안 원칙](../yama-word-obsidian/references/answer-principles.md)을 읽는다. 지정한 한 문항과 사용자가 묻는 범위만 다룬다.

1. 사용자가 붙인 문항을 그대로 사용한다. 자료에서 찾아야 하면 문항 전용 Word, 전연도 통합 Markdown 또는 Obsidian 문항집 한 파일을 선택하고 로컬 문항 인덱스를 만든다. 같은 원본은 인덱스를 재사용한다. 문항 번호가 중복되면 연도·차수·섹션·문두로 좁힌다. 여전히 모호하면 후보를 제시하고 확인한다.
2. 강의록과 Script에서 문두·선지의 핵심 용어와 변형, Word에 기록된 페이지로 먼저 검색한다. 자료가 한쪽만 있으면 제공된 자료에서 검색하고 빠진 자료를 밝힌다. 첫 검색에 근거가 부족하면 인접 구간, 동의어, 관련 페이지로 좁혀 다시 검색한다. PART 2 요약만 제공되면 교수의 원발언을 확인할 수 없다고 밝힌다.
3. 그림으로 판단해야 하면 문항에 연결된 그림과 근거 페이지의 그림을 실제로 확인한다. PDF가 이미지 기반이거나 텍스트가 빠져 있으면 후보 페이지의 원본 이미지를 연다. 파일이나 판독 정보가 부족하면 한계를 남긴다.
4. 정답 또는 판정, 그 이유, 필요한 선지 설명과 개념을 답한다. 근거마다 파일명과 실제 페이지·타임스탬프·문단·줄을 명시한다. 공식 답안이나 직접 답변이 확인됐는지 표시한다. 근거가 상충하거나 필요한 자료가 없으면 불확실한 지점을 밝힌다. 두 자료에서 확인하지 못한 꼭 필요한 내용만 외부 근거로 보충한다. 보충 자료와 강의 근거를 구분한다.

## 범위를 좁혀 읽기

`$WordSkill`과 `$QaSkill`은 현재 패키지에 포함된 각 스킬의 디렉터리다. `WORK`의 인덱스에는 자료 텍스트가 있으므로 로컬에 둔다. 인덱스 전체를 열지 말고 조회 결과를 읽는다.

```powershell
python "$WordSkill/scripts/question_index.py" build --source YAMA.docx --out WORK/questions.json
python "$WordSkill/scripts/question_index.py" lookup --index WORK/questions.json --year 25 --round 1차 --number 88 --figures WORK/figures
python "$QaSkill/scripts/local_evidence.py" build --lecture RAW.md --script SCRIPT.md --out WORK/evidence.json
python "$QaSkill/scripts/local_evidence.py" search --index WORK/evidence.json --term "핵심어" --term "영문 용어"
python "$QaSkill/scripts/local_evidence.py" search --index WORK/evidence.json --file SCRIPT.md --at 142 --context 3
```

- 문항 조회는 `--qid`, `--number`, `--text` 중 하나가 필요하다. 필요하면 `--year`, `--round`, `--section`을 추가한다. Word에 번호 이력이 없으면 ID나 문두를 사용한다. 한 문항이 선택되면 `--figures`에 연결된 그림만 꺼낸다.
- 자료 인덱스는 Markdown·TXT·DOCX 및 강의 PDF의 텍스트층을 지원한다. 강의록 파일이나 Script 파일이 여러 개면 해당 옵션을 반복한다. 검색 결과는 강의록과 Script 각각의 일치 여부, 발췌, 위치를 반환한다. 기본은 자료 역할마다 최대 2곳과 앞뒤 한 줄이다. 발췌 일부를 생략하면 표시한다.
- `--page N`으로 페이지를 좁히고 `--file`과 `--at`으로 한 위치를 연다. PDF의 줄 번호는 페이지마다 다시 시작하므로 `--at`과 함께 `--page`를 지정한다. 인덱스가 원본 변경을 보고하면 해당 인덱스만 다시 만든다. 파일에 텍스트가 있지만 일치 결과가 없으면 근거 공백을 선언하기 전에 검색어와 구간을 바꿔 확인한다.

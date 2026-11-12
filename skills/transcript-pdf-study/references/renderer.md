# 구조화 입력과 실행

필수 Python 패키지는 `python-docx`, `Pillow`, `PyMuPDF`다. Windows Word 경로는 `pywin32`도 필요하다. 외부 변환기는 Windows Word 또는 headless LibreOffice다. 변환기는 기본적으로 LibreOffice이며 `--converter word`로 Word를 선택한다. Word는 별도 비가시 인스턴스를 확인하고 파일 API로 변환한다. 기본 `font_family`는 Windows에서 확인한 `Malgun Gothic`이다. 다른 환경에서는 설치된 한국어 글꼴의 family name을 명시하고 PDF 검사로 확인한다. 폰트 설치·변형을 자동 수행하지 않는다. PDF에는 실제 사용된 글꼴이 포함되어야 하며 DOCX 편집 환경에도 해당 글꼴이 필요하다. 기존 RIDIBatang-9 OTF/라이선스는 선택 자산으로 보존하되 이 Word 환경에서는 대체되므로 기본 적용했다고 주장하지 않는다.

## JSON

경로는 `study.json`의 위치를 기준으로 해석한다. 로컬 절대경로도 허용하지만 공개 예시에는 개인 경로를 넣지 않는다. `transcript`, `lecture`, `question_sources`에는 검색 가능한 UTF-8 텍스트를 지정한다. 원본 PDF가 있으면 먼저 필요한 강의/기출 텍스트를 정확히 추출하고 원본 페이지를 확인한다.

```json
{
  "title": "검체 비교 학습 정리",
  "transcript": "transcript.md",
  "lecture": "lecture.md",
  "question_sources": ["questions.md"],
  "originals": ["lecture.pdf"],
  "font_family": "Malgun Gothic",
  "include_images": false,
  "keywords": [{
    "runs": [{"text": "검체 A"}],
    "grounding": [{"source": "transcript.md", "quote": "검체 A는 조건 1에서 관찰한다.", "locator": "PART 2 검체 비교"}]
  }],
  "key_summary": [{
    "runs": [{"text": "검체 A는 조건 1에서 관찰한다."}],
    "grounding": [{"source": "lecture.md", "quote": "검체 A는 조건 1에서 관찰한다.", "locator": "p. 2"}]
  }],
  "sections": [{
    "title": "검체 관찰 조건",
    "blocks": [{
      "type": "paragraph",
      "runs": [{
        "text": "검체 A는 조건 1에서 관찰한다.",
        "evidence": [{"kind": "question", "source": "questions.md", "quote": "검체 A의 관찰 조건은?", "locator": "24 Q1", "match_reason": "문항이 검체 A의 관찰 조건을 묻는다."}]
      }]
    }]
  }]
}
```

- `sections[].blocks`는 `paragraph`, `bullet`, `heading`, `table`, `image`를 지원한다. 문단에는 `runs`를 둔다. 각 run은 일반 문자열 `text`, 선택형 `bold`, 선택형 `evidence`를 가진다. HTML/Markdown을 해석하지 않는다.
- 문장 내부의 강제 개행은 거부한다. 문단으로 나누고 실제 줄바꿈은 Word가 계산한다. 글자 간격을 음수로 줄이거나 본문을 작게 줄여 억지로 맞추지 않는다.
- 표는 `{"type":"table","rows":[[[{"text":"검체"}],[{"text":"조건"}]],[[{"text":"A"}],[{"text":"1"}]]]}` 형식이다. 첫 행은 반복 헤더다. 셀마다 run 목록을 사용해 근거별 색을 유지한다. 열 수는 2–4개이며 모든 행의 열 수가 같아야 한다. 큰 셀은 의미 단위별 여러 표/본문으로 재편집한다.
- 이미지는 `{"type":"image","path":"figure.png","caption":"그림 1. 검체 비교","alt":"왼쪽 A와 오른쪽 B","source":{"source":"lecture.md","quote":"그림 1. 검체 비교","locator":"p. 3"}}` 형식이다. `include_images=true`는 사용자가 그림을 요청한 경우에만 지정한다. 그림 종횡비를 유지하고 캡션과 함께 배치한다.
- 두 번째 큰 주제부터 새 페이지를 시작한다. 첫 주제를 별도 페이지로 시작하려면 해당 section에 `new_page:true`를 지정한다. 밀접한 소절은 한 section 안의 heading으로 묶는다.

## 근거 규칙

`grounding`과 `evidence`의 `quote`는 지정 파일에 실제 존재해야 한다. 키워드·핵심 정리는 `transcript`의 PT2 또는 `lecture`만 근거로 쓸 수 있다. 근거 인용·위치는 작업 JSON에 남으며 납품 PDF/DOCX에 개인 파일 경로를 출력하지 않는다.

노랑은 `kind:question`이며 `source`가 `question_sources`에 포함되어야 한다. 초록은 `kind:professor_exam`, `explicit:true`이며 `source`가 `transcript`여야 하고 인용은 PART 1에 있어야 한다. 교수 발언 예: `{"kind":"professor_exam","source":"transcript.md","quote":"이 비교는 시험에 출제합니다.","locator":"PART 1 비교 설명","explicit":true,"match_reason":"교수가 같은 비교를 시험 대상으로 직접 지목했다."}`.

겹침은 같은 run의 evidence 목록에 두 근거를 넣는다. 본문 노랑과 인접 초록 표지로 둘을 보존한다. 기계 검사는 실제 인용 존재·파일 종류·명시 플래그를 확인한다. 인용이 표시한 내용과 의미상 맞는지, 부정 발언인지, 실제 교수 발언인지는 에이전트가 원문 문맥으로 검토해야 한다. `match_reason`은 검토를 대신하지 않는다.

## 검사와 실패

`render_study.py`는 DOCX 작성 후 새 작업 하위 폴더에서 PDF를 만든다. 새 PDF가 생성되지 않으면 기존 PDF를 성공 결과로 사용하지 않는다. `originals`에는 참조한 원본 PDF/PPTX를 모두 지정한다. 입력 JSON·전사·강의·기출·그림·원본과 겹치는 출력 및 이 파일들을 포함하는 작업 폴더는 `--overwrite` 여부와 관계없이 거부한다. `qc.json`과 전 페이지 `page-N.png`는 QA 근거다. 자동 검사는 빈 페이지, 페이지 바깥 텍스트, 선택한 글꼴의 PDF 내 존재를 확인한다. `visual_review`, `semantic_review`는 항상 `pending`으로 시작한다. 실제 검토 전 합격으로 바꾸지 않는다.

문자 누락, 의미 변경, 표/이미지 겹침, 모든 글자에 정확한 글꼴 적용, 긴 셀의 가독성은 자동 검사만으로 보증하지 않는다. 최종 PDF/DOCX SHA-256에 연결한 별도 검수 기록을 남긴다. 변환기/폰트가 없으면 설치 또는 이용 가능한 경로를 해결하고 재검수한다. DOCX만 생긴 상태를 두 형식 완료로 보고하지 않는다.

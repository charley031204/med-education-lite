# Med Education Lite

강의 전·후 학습을 위한 독립 경량화 플러그인입니다. 기존 Med Education 및 개인 설치 Skill을 변경하지 않습니다. 친구에게 저장소 링크와 [파이프라인 안내](docs/pipeline.md)를 함께 전달하면 됩니다.

## 포함된 Skill

| Skill | 입력 → 결과 | 토큰을 줄이는 처리 |
|---|---|---|
| `pdf-to-rawmd` | PDF/PPTX → 원문 RAW MD + 실제 그림 | 코드 초안, 최대 15페이지 작업, 필요한 영역 확대 |
| `yama-chapter-word` | 전체 기출 + 강의 범위 → 문제 전용 Word | 문항 색인·후보 조회, 원문 직접 조립 |
| `med-transcript-polish` | Whisper TXT + 강의록 → 페이지 태깅된 녹음정제 | 구간 분할, 관련 페이지 조회, 로컬 결합 |
| `transcript-pdf-study` | PART 2 + 강의록·기출 → PDF + 편집 가능한 Word | 구조화한 편집 원본 재사용, 코드 렌더링 |
| `yama-word-obsidian` | 문제 Word + 강의 근거 → Obsidian 해설 | 문항별 추출·근거 조회, 원문 직접 재사용 |
| `yama-question-answer` | 한 문항 + 질문 → 근거가 있는 답변 | 해당 문항과 강의록·Script의 관련 구간만 조회 |

`transcript-pdf-study`는 기존 `transcript-study-pdf`를 바탕으로 새 요구사항을 반영한 이름입니다. Notion의 `/Med-transcript-polish`는 이 패키지에서 소문자 `med-transcript-polish`로 호출합니다. 플러그인으로 설치한 호스트에서는 `med-education-lite:` 접두사가 표시될 수 있으므로 검색되는 정식 명령을 선택합니다.

RAW 변환과 기출 Word 제작은 **서로 다른 새 채팅**에서 진행합니다. RAW 출력은 강의록마다 새 폴더를 만들어 MD와 전용 `동봉 이미지/`를 함께 둡니다. 기존 폴더와 충돌하면 `_2`, `_3`을 붙입니다.

정리본은 HTML을 납품하지 않습니다. 기출 대응 내용은 매우 연한 노란색, 교수의 실제 시험 언급은 매우 연한 초록색으로 구분하며 근거가 없는 강조를 만들지 않습니다. 이미지 포함은 요청에 따릅니다. 사용자 환경에 정확한 Google Drive 목적지가 설정되어 있으면 완성된 PDF·Word를 업로드하고 확인합니다. 실제 목적지나 인증 정보는 패키지에 포함하지 않습니다.

## 설치와 사용

Claude Code에서 별도 플러그인으로 설치할 때:

```text
claude plugin install med-education-lite --marketplace charley031204/med-education-lite
```

Python 도구를 사용할 환경에 저장소의 의존성을 설치합니다.

```text
python -m pip install -r requirements.txt
```

정리본 PDF 변환에는 Microsoft Word 또는 LibreOffice가 필요하며, 이번 검증은 Windows Word로 수행했습니다. 기출 Word의 페이지 실측에는 Windows Word와 `pywin32`가 필요합니다. 정리본 기본 글꼴은 검증한 `Malgun Gothic`이며 다른 환경에서는 설치된 한국어 글꼴을 지정하고 실제 PDF를 확인합니다. Word 파일이 만들어졌다는 사실만으로 PDF 변환이나 인쇄 검수가 완료됐다고 보고하지 않습니다. 사용할 수 없는 외부 프로그램이나 업로드 연결은 별도로 표시합니다.

루트 `plugin.json`은 portable plugin 형식이고 `.claude-plugin/`에는 호환 카탈로그가 있습니다. 설치 없이 저장소의 원하는 `skills/<이름>/SKILL.md`를 직접 지정해도 됩니다. 버전 번호는 지정하지 않았으며 개발 상태는 Git 커밋으로 식별합니다.

## 품질과 검증

원문·선지·출제 이력·실제 그림·불확실성을 보존합니다. source-map, 중앙 실행 레지스트리, 전역 hook은 사용하지 않습니다. 추출·검색·기계 검사는 모델의 내용 판단과 시각 검토를 보조합니다. 실제 실행 토큰 절감률은 입력 자료와 모델에 따라 달라지며 아직 일반화할 수 없습니다.

```text
python -m pytest -q tests
claude plugin validate .claude-plugin/plugin.json
claude plugin validate .claude-plugin/marketplace.json
```

검증 범위와 한계: [초기 RAW 이전](docs/validation/pdf-to-rawmd.md), [학습 파이프라인](docs/validation/pipeline.md).

원본 구현: [Med Education Plugin](https://github.com/charley031204/med-education-plugin). MIT License 및 별도 자산 라이선스를 보존합니다.

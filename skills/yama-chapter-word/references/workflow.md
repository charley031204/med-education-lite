# 문항 색인과 Word 생성

## 입력과 후보 패킷

Python에서 `python-docx`와 `Pillow`를 사용할 수 있어야 한다. 실제 페이지 실측과 검수 PDF 출력은 Windows, Microsoft Word, `pywin32`가 있는 환경에서 실행한다.

```sh
python scripts/yama_bank.py index --rawmd "bank.md" --csv "bank_questions.csv" --out "work/index.json"
python scripts/yama_bank.py packet --index "work/index.json" --topic "대상 주제" --alias "시간표 별칭" --term "교안 고유어" --out "work/packet-0.json"
```

- Index는 기존 인라인 `**Q. 9. 발문** [25본1_1차]`와 organizer의 `<!-- question qid=... -->` 문항 경계를 모두 읽는다. Index는 이 작업의 문항 검색용 파일이다.
- Packet은 기본 20문항이다. 같은 검색 조건과 stdout의 `next_offset`으로 다음 묶음을 만든다. 중간 본문을 읽지 않고 후보 수를 확인할 수 있다. 패킷의 `source_lines`는 organizer 마커형의 시작·끝 또는 인라인형의 문두 줄이다.
- `--alias`와 `--term`은 반복할 수 있다. 자료로 시험지 범위가 확인된 경우 `--number-range 2025:11-18`로 확인한다. 번호 범위 자체는 소속 확정 근거가 아니다. 서술형 누락 가능성이 있으면 `--include-written`을 함께 쓴다.
- `csv_links`와 `csv_has_candidates`를 확인한다. 미연결·모호한 행이나 비어 있는 후보 집합은 미확인으로 유지한다. 이미 배정된 대상 문항은 주제 후보가 비어 있어도 패킷에 포함된다.
- 한 문항이 여러 챕터에 배치되었다면 서로 다른 시험지의 문항인지 같은 원문의 반복 배치인지 확인한다. 원문 ID와 시험지·번호·본문을 대조한 뒤 선택한다.

## 확정 문항 복사

패킷에서 확인한 ID만 선택 파일에 넣는다. 두 해시는 index·packet 요약값을 복사한다. CSV 없이 만들었다면 `csv_sha256`은 `null`이다. 여러 주제를 묶을 때는 각 절에 ID 목록을 둔다.

```json
{
  "raw_sha256": "RAW SHA-256",
  "csv_sha256": null,
  "chapters": [
    {"topic": "대상 주제", "current": "담당 교수", "source_ids": ["source-id-1", "source-id-2"]}
  ]
}
```

```sh
python scripts/build_yama_docx.py --rawmd "bank.md" --index "work/index.json" --selection "work/selection.json" --subject "과목" --exam "차수" --out "output" --outname "강의 기출 야마.docx" --report "work/report.json" --photo-work "work/photo" --preview-pdf "work/preview.pdf"
```

변경된 RAW는 다시 색인하고 선별을 확인한다. CSV가 변경되면 다시 색인한 뒤 후보와 선택을 확인한다. CSV가 RAW와 다른 폴더에 있으면 빌드에도 `--csv`를 전달한다. 사진 검토가 필요한 첫 실행은 보고서에 `review_file`과 contact sheet를 남긴다. 동일한 선택으로 검토를 완료한 뒤 `--photo-review`를 추가한다.

검토한 강의 페이지가 있으면 `--page-map`을 전달한다. 기본값은 선택한 ID 순서이며, 페이지 매핑을 전달하면 교안 페이지순으로 배치한다. 연도 우선은 `--sort year`를 함께 지정한다. 동일 번호가 여러 시험지나 교수 구간에 걸치면 매핑 근거로 정확한 문두 또는 식별 가능한 멤버를 쓴다. 공통지문은 원문에서 관계가 확인된 문항 앞에 둔다. 기존 빌더의 `--help`에서 페이지·공통지문·연도별 교수 매핑 형식을 확인할 수 있다.

이미 확정된 챕터 전체만 필요한 경우 `--topic`으로 바로 빌드할 수 있다. 강의별 소속 검증과 다른 절에서 누락된 후보 검색은 앞 단계에서 수행한다.

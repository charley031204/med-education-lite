# 공통사진 판정과 원본 연결

`shared_photos.py`는 Word와 Obsidian이 함께 읽는 `yama-shared-photos/1`
manifest를 만든다. 사진의 동일성 판정과 문항의 소속은 이 파일 하나에서 가져온다.
새 형식으로 출력할 때 빌더는 현재 원문, 문항 목록, 이미지 파일을 다시 검증한다.
검토가 끝나지 않았거나 원본이 달라졌으면 `CHECK`로 종료한다.

## 1. 입력과 전체 조사

각 문항은 아래 정보를 제공한다. 문항 순서, 절 경계, 원래 이미지 출현 위치를
보존한다. 사진이 없는 문항도 `images: []`로 포함한다. Indication이나 해설이
없는 사진도 입력에서 빼지 않는다.

```json
{
  "id": "q018",
  "section": "신경계 실습",
  "images": [
    {"path": "동봉 이미지/A.jpg", "occurrence_id": "q018:A", "body_index": 2, "caption": "A"},
    {"path": "동봉 이미지/B.jpg", "occurrence_id": "q018:B", "body_index": 5, "caption": "B"}
  ],
  "payload": {"source_question": "빌더가 보존하는 원문 구조"}
}
```

- `id`는 문서 안에서 고유하다. `section`은 같은 이름의 별도 절까지 구별하는
  빌더의 절 식별자다. 같은 사진을 쓰더라도 다른 절의 문항은 합치지 않는다.
- `images`는 문자열 경로도 받는다. dict의 추가 키와 `payload`는 그대로 보존한다.
  `occurrence_id`를 생략하면 문항 ID와 이미지 순서로 만든다.
- inventory는 모든 이미지 출현과 참조 파일을 기록한다. 파일별 SHA-256,
  폭·높이, 디코딩한 RGBA 픽셀 해시를 포함한다. EXIF 방향을 적용하며 ICC profile이
  다른 파일을 자동으로 같은 픽셀 사진으로 처리하지 않는다.
- 바이트 또는 표시 픽셀이 정확히 같은 파일만 자동으로 같은 집단에 둔다.
  다중 프레임 파일은 첫 장만 가져오지 않고 명시적인 추출을 요구한다.
- **모든 고유 바이트 대표**를 전체 contact sheet에 수록한다. 픽셀은 같지만
  인코딩이 다른 대표도 조사 대상에 남긴다. 파일명이 같거나 질환명이 같다는
  이유로 사진을 생략하지 않는다.
- 서로 다른 픽셀 집단의 **모든 조합**을 pair queue에 만든다. 정규화한 색상 차이는
  검토 순서만 정한다. 값이 크다는 이유로 pair를 제외하거나 `different`로
  자동 판정하지 않는다.

## 2. 시각 검토와 완료 조건

contact sheet 생성은 열람의 증거가 아니다. 주 검토자와 독립 검토자는 전체
대표를 실제로 보고, 동일할 가능성이 있는 pair는 원본 해상도로 확대 대조한다.
전체 목록을 보기 전에 이미 알려진 동일 사진 두 쌍만 골라 승인하면 안 된다.

`discovery_review`에는 두 검토자의 이름, 실제로 본 대표 파일 ID 목록,
검토 근거와 `evidence_digest`를 기록한다. 목록은 각각 모든 고유 바이트 대표를
빠짐없이 한 번씩 포함해야 한다. 독립 검토자는 첫 검토자와 다른 사람 또는
독립적인 검토 에이전트다. 단순히 이름 두 개를 적는 것으로 실제 검토를 대체하지 않는다.

각 pair에는 다음 필드를 채운다.

| 필드 | 조건 |
|---|---|
| `verdict` | `same`, `different`, `uncertain` 중 하나 |
| `rationale`, `reviewer` | 실제 시야·패널·구조 배열을 근거로 판정 |
| `independent_verdict`, `independent_rationale`, `independent_reviewer` | 독립 검토 결과와 근거 |
| `evidence_digest` | 현재 `inventory_digest`와 일치 |
| `identity_basis` | `same`이면 `same_visual_field` |
| `source_visual_information_preserved` | `same`이면 `true` |
| `preservation_rationale` | `same`이면 패널·화살표·문자·스케일바·전체 시야의 보존 근거 |

`same`은 같은 질환이나 비슷한 조직 형태를 뜻하지 않는다. 실제 같은 시야와
구조 배열을 확인해야 한다. 같은 시야라도 원본마다 다른 화살표나 캡션을 대표
한 장에 보존할 수 없다면 교체를 승인하지 않는다. 미확정 상태를 `different`로
바꿔 검사를 통과시키지 않는다. `pending`, `uncertain`, 검토자 간 불일치는
완료를 막는다. A=B, B=C인데 A≠C라는 판정도 완료를 막는다.

바이트·픽셀이 다른 사진을 묶는 경우에는 대표도 명시적으로 검토한다.
`representatives[photo_id]`에 아래 객체를 기록한다.

```json
{
  "file_id": "f0022",
  "suitable": true,
  "source_visual_information_preserved": true,
  "display_notes": ["필요한 경우 원본별 패널 표지를 그대로 설명하는 문구"],
  "reviewer": "실제 주 검토자",
  "independent_reviewer": "실제 독립 검토자",
  "rationale": "같은 전체 시야이며 A 표지와 원래 화살표를 보존한 고해상도 파일이다.",
  "independent_rationale": "원본들을 확대 대조하여 시야, 표지, 화살표의 보존을 확인했다.",
  "evidence_digest": "현재 inventory_digest"
}
```

`photo_id`는 component에서 가장 작은 exact-group ID에 `photo:`를 붙인 값이다.
예를 들어 `e0008`과 `e0014`를 묶으면 `photo:e0008`이다. 대표는 반드시 그
component의 파일이어야 한다. 같은 시야와 정보를 보존하는 파일 중 적절한
최고 해상도 파일을 선택한다. 더 큰 파일의 흐림·잘림 등으로 다른 대표를
선택하면 두 검토자가 이유를 기록한다. 정확한 픽셀 동일성만으로 이뤄진
집단은 첫 동률 파일을 자동 선택할 수 있다.

선택적 `display_notes`는 원본별 표지 차이를 보존하는 표시 문구다. 대표
파일을 고칠 필요 없이 캡션으로 원래 표지를 설명할 수 있을 때 사용한다.
`grouping_plan()`은 이 목록을 사진과 부모의 이미지 기록 모두에 그대로
전달한다. Word와 Obsidian 렌더러는 같은 문구를 표시하고 QC에서 확인한다.
시야·구조 자체의 누락을 문구로 대신 보존했다고 승인해서는 안 된다.

## 3. 사진 집단과 표시 부모

사진 집단과 문항의 표시 부모는 다르다. `grouping_plan()`은 같은 절에서
사진을 공유하는 문항들을 연결해 부모를 만든다. 원문에서 처음 나온 문항의
순서로 부모를 배치한다. 사진 없는 문항은 개별 부모로 남긴다.

q010과 q018이 A를 공유하고 q018에만 B가 있으면 부모는 두 문항을 포함한다.
부모의 `images`에는 A와 B가 모두 있지만, A는 `shared: true`, B는
`shared: false`이며 `question_ids: ["q018"]`이다. 렌더러는 A를 공통 영역에
한 번 놓고 B는 q018의 위치에 둔다. 각 문항의 `image_memberships`에는
모든 원본 `occurrence_id`, `original_path`, 위치, 캡션을 남긴다. 원래
문두·선지·출제 이력·해설은 `payload`와 문항 매핑으로 보존한다.

같은 사진이라는 판정은 실제 시험에서 같은 사진을 재사용했다는 증거가 아니다.
상류 추출 과정의 복제 가능성과 출처 불확실성은 기존 공시 규칙에 따라 남긴다.

## 4. API와 CLI

```python
prepare_manifest(questions, source_digest, image_root) -> dict
validate_manifest(manifest, questions=None, source_digest=None, image_root=None,
                  *, require_complete=True) -> list[str]
grouping_plan(manifest) -> dict
inventory_digest(manifest) -> str
manifest_digest(manifest) -> str
digest_json(value) -> str
source_digest(path) -> str
canonical_json(value) -> bytes
export_manifest(manifest, path) -> None
contact_sheets(manifest, image_root, output_dir, *, pairs=False, pair_ids=None) -> dict
embed_docx_manifest(docx_path, manifest, payload, *, image_root=None) -> None
read_docx_manifest(docx_path) -> dict | None
extract_docx_images(docx_path, output_dir) -> dict[str, str]
```

`validate_manifest()`는 오류 목록이 비어 있을 때만 통과한다. `image_root` 없이
실행하면 구조와 검토 증거를 검사하는 순수 함수다. **새 문서 빌더는 현재
questions, source_digest, image_root를 모두 전달**하여 원문 변경, 누락된
문항/이미지 매핑, 수정·삭제된 파일을 함께 검사한다. `require_complete=False`는
검토용 contact sheet 생성에만 쓴다. 최종 출력 승인에는 사용할 수 없다.

`image_root`는 기본 폴더 또는 `{원래 경로: 실제 파일 경로}` alias map을 받는다.
alias map을 사용하면 추출 위치가 달라져도 manifest의 원래 경로를 수정할 필요가 없다.
JSON export는 정렬된 키와 UTF-8로 결정적으로 저장한다. 검토 완료 뒤 manifest를
고치면 digest가 달라지므로 이전 QC와 DOCX 연결을 재사용하지 않는다.

```powershell
python shared_photos.py inventory --questions questions.json --source source.md --image-root source-folder --out shared-photos.json
python shared_photos.py contact-sheets --manifest shared-photos.json --image-root source-folder --out review-folder
python shared_photos.py contact-sheets --manifest shared-photos.json --image-root source-folder --out pair-folder --pairs --pair-id e0008:e0014
python shared_photos.py validate --manifest shared-photos.json --questions questions.json --source source.md --image-root source-folder
```

`--pairs`만 지정하면 모든 pair sheet를 만든다. `--pair-id`는 확대 검토 자료를
선택하는 옵션이며, manifest의 전체 pair coverage를 줄이지 않는다. 실패한
validation CLI는 종료 코드 1을 반환한다. inventory와 contact sheet의
`REVIEW_REQUIRED`는 완료 상태가 아니다.

## 5. DOCX와 Obsidian 사이의 전달

DOCX는 `customXml/yama-shared-photos.json`에 검증한 manifest와 빌더 payload를
넣는다. 원본 사진 파일도 별도의 custom part에 바이트 그대로 보존한다.
사진마다 원래 경로, file ID, 저장 part, SHA-256을 기록한다. 화면에는 대표만
나타나더라도 Word→Obsidian은 모든 원래 alias를 되살릴 수 있다.

envelope는 manifest/payload digest와 `word/document.xml`, Word 관계 파일,
전체 `word/media/` 파일의 SHA-256에 연결된다. `[Content_Types].xml`,
패키지 루트 관계, 원본 alias 관계 파일도 모든 part를 완성한 뒤 같은 binding에
넣는다. 읽을 때는 각 alias와 manifest를 실제로 연결하는 내부 관계와
ContentType도 별도로 대조한다. 관계 파일이 빠졌거나 target·종류·소속이
달라지면 원본 bytes가 ZIP에 남아 있어도 읽기와 추출을 거부한다.
custom part 자체를 DOCX
해시로 다시 묶는 순환은 만들지 않는다. 전체 DOCX 해시는 외부 layout QC가
담당한다. 문서 본문·관계·사진·payload·alias가 달라지면 읽기를 거부한다.
일반 DOCX에는 `read_docx_manifest()`가 `None`을 반환하지만, 공유사진 관계가
남아 있는데 manifest part만 없어진 문서는 오류다.

`extract_docx_images()`는 검증 후 별도 출력 폴더에 `f0001.png` 같은 안전한
이름으로 추출한다. 원본 경로를 출력 경로로 직접 사용하지 않는다. 같은
출력 이름에 다른 파일이 있으면 덮어쓰지 않는다. 반환한 alias map으로
Obsidian IR을 검증하고 Word 재출력에도 같은 manifest를 전달한다.

이 계약의 검증은 사진 판정과 원본 연결을 확인한다. 기존 `yama-print/1.0.7`의
이미지폭, Word 실측, 전체 페이지 시각 검토, 원문/문항 수 검사를 대체하지 않는다.

회귀 테스트는 `tests/test_shared_photos.py`에 있다. 모든 fixture는 코드로 만든
도형이며 실제 의료 사진을 저장소에 넣지 않는다.

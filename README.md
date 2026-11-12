# Med Education Lite

Med Education 스킬을 하나씩 이전하는 독립 경량화 플러그인입니다. 현재 포함된 스킬은 **`pdf-to-rawmd` 하나**입니다. 기존 `med-education-plugin` 및 개인 설치 Skill과 별도로 관리합니다.

## 첫 번째 이전: PDF / PPTX → RAW MD

원문·표·수식·강조·실제 그림과 기존 RAW MD의 보이는 양식을 보존합니다. 최대 15페이지의 새 변환 작업자와 독립 전수 검토자는 유지합니다. 영구 source-map, 숨은 블록 ID, 블록별 원본 연결 JSON, stale 연결 검사는 제거했습니다. 따라서 이 플러그인은 최종 문단에서 원본 객체 좌표를 조회하는 기능을 제공하지 않습니다.

출력 상위 폴더 아래에 강의록마다 새 폴더를 만들고, 그 안에 `[원본파일명(확장자 제외)]_rawmd.md`와 전용 `동봉 이미지/`를 둡니다. 같은 이름의 강의 폴더가 있으면 `_2`, `_3`을 붙여 새로 만들며 기존 자료는 이동하거나 합치지 않습니다.

토큰 입력과 생성을 줄이기 위해 다음 흐름을 사용합니다.

1. 코드가 PDF/PPTX의 native 텍스트로 **미검수 초안**을 파일에 만듭니다.
2. 작업자가 담당 범위의 텍스트·서식·검토 후보를 묶어서 읽고, 원본 화면과 대조해 필요한 부분을 수정합니다.
3. 그림은 배치를 확인한 뒤 필요한 영역을 확대합니다. 표 셀과 그림 설명은 생략하지 않습니다.
4. 독립 검토자가 전 페이지의 내용·표·학습 Figure를 대조하고 문제 위치를 보고합니다. 수정은 해당 범위에 적용합니다.
5. 코드는 chunk 결합, 페이지 순서와 이미지 경로 검사를 수행하고 수치만 보고합니다.

자동 추출과 기계 검사는 내용 검수를 대체하지 않습니다. 스캔·복잡한 읽기 순서·수식·SmartArt·상속 서식은 시각 확인과 보정이 필요합니다. 실행 토큰 절감률은 아직 측정하지 않았습니다.

## 사용

Python 의존성을 설치합니다.

```text
python -m pip install -r requirements.txt
```

설치 없이 저장소의 `skills/pdf-to-rawmd/SKILL.md`를 지정해 사용할 수 있습니다. 플러그인은 루트 `plugin.json`의 portable 형식과 `.claude-plugin/` 호환 카탈로그를 제공합니다. 사용자가 설치를 요청할 때 이 저장소를 별도 marketplace로 등록하면 됩니다. 이번 이전 작업은 기존 Skill이나 플러그인 설치를 변경하지 않습니다.

버전 번호는 지정되지 않았습니다. 현재 개발 상태는 Git 커밋으로 식별하며 버전 태그를 만들지 않습니다. Claude의 strict validator는 버전 미지정을 경고할 수 있습니다.

## 검증과 확장

```text
python -m pytest -q tests
claude plugin validate .claude-plugin/plugin.json
claude plugin validate .claude-plugin/marketplace.json
git diff --check
```

검증 범위와 한계는 [초기 이전 검증 보고서](docs/validation/pdf-to-rawmd.md)에 기록합니다. 사용자가 다음 스킬을 지정하면 같은 `skills/` 아래에 추가합니다. 중앙 실행 레지스트리, 배포 동기화 도구, 전역 hook은 포함하지 않습니다.

원본 구현: [Med Education Plugin](https://github.com/charley031204/med-education-plugin). MIT License와 기존 저작권 고지를 보존합니다.

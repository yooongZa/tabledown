# Tabledown — 작업 지침 (AI/개발자용)

Tabledown 은 macOS 메뉴바 앱으로, 클립보드를 감시하며 **Excel/Sheets ↔ Markdown 표**를
양방향 변환한다. 핵심 로직은 `tablemark/app.py` 의 `TabledownApp._converted_clipboard` 와
`tablemark/converter/` 에 있다.

---

## ⚠️ 클립보드 변환 불변식 (회귀 금지)

아래는 **실제 사용자가 깨진 걸 겪은 회귀**에서 도출한 규칙이다. 변환 로직을 수정할 때
반드시 지키고, 끝에 적힌 회귀 방지 테스트를 통과시킬 것. "코드가 더 깔끔해 보여서"
이 규칙을 건드리면 과거 회귀가 그대로 재발한다.

### 0. 멀티포맷 클립보드 — 도착지가 슬롯을 고른다
macOS 클립보드는 **text(일반 텍스트) 슬롯과 html 슬롯을 동시에** 가진다. 붙여넣는
앱(도착지)이 자기에게 맞는 슬롯을 고른다 — Excel·Word·메모는 html `<table>`, 마크다운
에디터(Obsidian 등)는 text. **TableDown 은 복사 시점에 도착지를 알 수 없다.** 따라서
"한 슬롯만 맞추는" 변환은 틀린 접근이고, 가능하면 **두 슬롯을 공존**시켜 도착지가 각자
고르게 한다.

### 0-W. Windows CF_HTML — Excel fragment 는 `<table>` 래퍼를 뺀다 (회귀 금지)
- Windows 의 html 슬롯은 **CF_HTML**("HTML Format") 포맷이다. 헤더에 `StartFragment`/`EndFragment`
  바이트 오프셋이 있고, `extract_cf_html`(`tabledown_windows/html_clipboard.py`)이 그 구간을 잘라 쓴다.
- **함정**: Excel·Google Sheets 는 `StartFragment`/`EndFragment` 마커를 **`<table>` 엘리먼트 *안쪽*** 에
  둔다. 그래서 잘라낸 fragment 는 표 내부(`<col>`/`<tr>`/`<td>`)뿐이고 **`<table>` 여닫는 태그가 빠진다.**
  표 감지(`has_html_table` = `"<table" in html`)가 실패 → 진짜 Excel 표가 "표 아님"으로 판정 →
  변환 없이 **원본 TSV 가 그대로 남는다**(`converted_clipboard` 가 None). 이게 0.2.4~0.2.5 초기에
  "Excel 표가 마크다운으로 변환 안 됨"의 원인이었다. (웹·채팅 표는 fragment 에 `<table>` 이 들어있어
  영향 없음 — 그래서 일부만 동작해 보임.)
- **수정/불변식**: `extract_cf_html` 은 행(`<tr>`)이 있는데 `<table>` 래퍼가 없으면 `<table>…</table>` 로
  **감싼다**(`_ensure_table_wrapper`). 비-표 HTML 은 절대 건드리지 말 것(`<tr>` 없으면 그대로). **이 래핑을
  지우면 Excel→마크다운이 전부 깨진다.** 회귀 테스트: `real_excel_table_converts_to_markdown`,
  `excel_cf_html_fragment_without_table_tag_is_wrapped`(Excel 형식 CF_HTML — 마커가 `<table>` 뒤를 가리킴 — 을 재현).
- **테스트는 실제 클립보드 형식으로**: `<table>` 을 포함한 *이상적* HTML 로만 테스트하면 이 버그를 못 잡는다.
  실제 Excel CF_HTML(또는 그 형식을 흉내낸 fixture)로 검증할 것.

### 1. text 에 마크다운 표 + html 에 `<table>` → 원본 유지 (셀개수 검사 금지)
- 웹·채팅(Claude 등)에서 복사한 표는 clipboard 에 **마크다운 text 와 html `<table>` 을 함께** 싣는다.
- 이때는 **칸수(셀 개수)를 따지지 말고 원본 clipboard 를 그대로 둔다.**
  html `<table>` 의 존재 자체가 "진짜 표"라는 독립적 증거다.
- 구현: `is_markdown_table(text, strict=...)` 호출 시 html `<table>` 이 동반되면
  `strict=False` (셀개수 검사 끔). → `_converted_clipboard` 의 `has_html_table` 분기.
- **회귀 이력**: 0.1.0 은 이 동작(관대)이었다 → 0.2.0 이 셀개수 검사를 **무조건** 적용하면서
  칸수가 어긋난 진짜 표가 "표 아님"으로 판정됨 → html 이 제거되고 마크다운으로 변환됨 →
  **Excel 붙여넣기에 마크다운 원문이 한 셀에 박힘** → 0.2.1 에서 복원.

### 2. 셀개수 검사(strict)는 html `<table>` 이 없을 때의 text(= 마크다운 표 후보)에만
- **배경**: html `<table>` 이 없고 text 가 마크다운 표 모양(`| a | b |` + `| --- | --- |`)이면,
  그건 보통 "마크다운 에디터에서 마크다운 표를 복사한" 경우다. 이때 TableDown 은 그 text 를
  Excel 이 셀로 받을 수 있게 html `<table>` 을 **생성**해 붙인다 (`markdown_table_to_html`,
  `_converted_clipboard` 의 첫 분기). ← "표가 없는데 왜 검사하나"가 아니라, **이 text 가
  진짜 마크다운 표라서 Excel 용 html 표로 변환할 가치가 있는지** 판정하는 것.
- 셀개수 일치 검사의 목적은 그 변환 직전의 **false positive(거짓양성) 차단** — 셸 출력·ASCII
  박스처럼 우연히 `| ... |` 모양이고 둘째 줄이 `-` 인 일반 텍스트를 표로 오인해, 멀쩡한
  텍스트에 html 표를 멋대로 만들어 클립보드를 오염시키는 것 방지.
- html `<table>` 이 있으면 이 판정이 불필요(이미 진짜 표 증거)하므로 검사를 끈다(불변식 1).
- 즉 셀개수 검사를 **삭제하지도, 무조건 적용하지도 말 것.** 적용 범위가 핵심이다.

### 3. html 슬롯은 어떤 표 변환에서도 제거하지 않는다 (0.2.4 에서 통일)
- text 슬롯에 마크다운을 넣는 것과 html 슬롯을 지우는 것은 **완전히 별개**의 조작이다.
  html 이 사라지는 건 오직 `write_clipboard(..., drop_types=세트에 HTML_TYPES 포함)` 일 때뿐.
- **모든 표 케이스(웹표·문서·순수 Excel/Sheets 표)에서 html `<table>` 을 유지한다.**
  text 슬롯엔 마크다운을 보강하고 html 은 원본 그대로 둔다. 그래야 한 번 복사로 Excel·Word 는
  표 형식을, 마크다운 에디터는 마크다운을 받는다. `_converted_clipboard` 의 모든 분기는
  `drop_types` 에 `RENDERED_TABLE_TYPES`(PNG/PDF/RTF 등 이미지)만 넣고
  **`HTML_TYPES` 는 절대 넣지 않는다.**
- **이력**: 0.2.3 까지는 "순수 Excel 표 → 마크다운 에디터" 한 경우만 html 을 제거했다
  (Obsidian 이 html 을 우선해 마크다운이 안 되는 것 회피 — 커밋 `7c677b8`). 그러나 그 표를
  다시 Excel·Word 에 붙이면 표가 깨지는 손실이 있어(불변식 1 과 같은 문제), 0.2.4 에서 다른
  표 케이스와 동일하게 **html 유지로 통일**했다.
- **트레이드오프(주의)**: html 자동변환이 켜진 마크다운 에디터(Obsidian 등)는 html 을 우선
  받아 리치 표로 붙일 수 있다(text 의 마크다운이 무시될 수 있음). 이는 도착지 앱 정책이라
  TableDown 통제 밖이며, "표 형식 붙여넣기를 잃지 않는다"는 이득과의 교환이다.

### 4. 표가 포함된 "문서" → text 에 마크다운 표 보강 + html 유지 (0.2.3)
- 문단·헤딩·리스트 사이에 표가 섞인 문서(`html_has_content_outside_table` 가 True)는
  **표만 추출하면 안 된다** (나머지 텍스트가 통째로 소실됨 — 0.2.2 이전 버그).
- text 슬롯: `convert_document_tables` 로 **표 부분만 마크다운 표**로, 나머지는 plain text 로 보강.
- html 슬롯: **원본 유지** (PNG/PDF/RTF 등 RENDERED 이미지 형식만 drop, html 은 절대 drop 금지).
- 결과: 마크다운 에디터는 마크다운 표를, Word·Excel 은 원본 표 형식을 받는다. **양쪽 동시 만족.**

### 5. XML 표 변환 — 표→XML은 명시적 메뉴 클릭 전용, 자동 역변환은 없음 (0.3.0)
- **형식 (v2 — 다단 헤더 중첩, 방안1: 일반 태그)**: LLM 친화 XML. 다단 헤더(가로·세로 양방향)를
  XML **중첩 계층**으로 보존한다(flatten 금지). 루트 `<표>`, 세로 상위그룹 `<{헤더}그룹 이름="값">`,
  행 `<행 {헤더}="값">`, 가로 상위그룹 `<열그룹 이름="값">`, 가로 리프(셀) `<열 n="헤더값">셀값</열>`.
  로직은 `converter/table_xml.py`(`model_to_xml`/`table_xml_to_model`), 입력 모델은
  `converter/html_to_md.py` 의 `html_table_to_model`(헤더 레벨들 + 데이터 행).
  ```xml
  <표>
    <직급그룹 이름="부장">
      <행 직책="대족장">
        <열그룹 이름="1분기"><열 n="1">동</열><열 n="2">해</열><열 n="3">물</열></열그룹>
        <열그룹 이름="2분기"><열 n="4">과</열><열 n="5">백</열><열 n="6">두</열></열그룹>
      </행>
    </직급그룹>
  </표>
  ```
- **루트는 한글 `<표>`, 절대 `<table>`/`<tr>` 금지**: `<table>`(영문) 은 진짜 HTML 태그라, 이 XML
  텍스트가 HTML 로 렌더링되는 곳(브라우저·Obsidian 미리보기·리치텍스트)에 가면 HTML 표 파싱이
  걸려 비-table 자식을 **표 밖으로 쫓아내(foster-parenting)** 표가 텅 빈다. 한글 `표`/`행`/`열`/
  `열그룹`/`{X}그룹` 은 HTML 에 없는 이름이라 트리가 보존된다(v1 이 `<dataset>` 을 쓴 것과 같은
  이유). (이력: 초기 0.3.0 은 루트가 `<table>` 이라 대화창·Obsidian 등에서 셀이 사라지는 듯
  보였다 — html5lib 로 재현·확인.)
- **가로축 태그(`열그룹`/`열`)는 고정 일반 태그, 헤더는 속성값(`이름`/`n`)**: 헤더는 *데이터*라
  공백·앞자리 숫자·기호(`( ) % /`)·중복·예약어 `xml` 등 XML 이름 규칙에 안 맞는 게 흔하다. 가로는
  헤더를 속성값(escape 만 하면 무엇이든 OK)에 넣어 정규화·이름충돌·백업속성 같은 군더더기를
  없앤다. 가로 차원 이름(`분기`,`월`) 은 표에 없으므로 만들지 않는다(방안1). **가로 헤더를 태그로
  돌리지 말 것.** (이력: 초기 0.3.0 은 `<Q1_2024 header="Q1 2024">` 식으로 헤더를 태그로 써서
  지저분한 헤더에선 태그가 뭉개지고 이름이 두 번 적혔다 — 분석 후 폐기.)
- **세로 키 헤더만 태그에 노출 → 항상 유효 XML 이름 보장**: 세로 상위그룹은 `<{헤더}그룹>`(접미사
  `그룹`), 행 최하위 키는 `<행 {헤더}="값">` 의 속성 key 로 헤더를 쓴다. 그래서 세로 키 열은
  **헤더가 유효 XML 이름인 동안만** 키로 인정한다 — 왼쪽부터 첫 가로 group 앞까지의 leaf 중,
  XML 이름 규칙 위반(공백·기호·숫자시작·`xml` 접두) 열을 만나면 거기서 키 종료(그 열부터 데이터).
  → **정규화 불필요, 무손실 roundtrip 유지.** 가로 group 이 하나도 없으면(단순 표) 세로 키 0개 →
  모든 열이 데이터, 행은 `<행>`(키 속성 없음). 키 1개면 `<행 {헤더}="값">`(세로 그룹 없음).
- **가로도 세로도 중첩 (v2 의 핵심 변경: 세로 forward-fill → 세로 중첩)**: 가로 그룹 헤더(병합으로
  그린 1분기>1,2,3)는 `<열그룹>` 으로 중첩한다(종전과 동일). v1 은 세로 병합(직급 부장/차장)을
  계층이 아니라 "같은 값 반복"으로 보고 **각 행에 forward-fill 해 자기완결 레코드**로 만들었으나,
  v2 는 세로 그룹도 부모 노드 `<{헤더}그룹 이름="값">` 로 **중첩**한다(인접 행의 동일 그룹값을
  연속 런으로 묶고, 다단계면 다단 중첩). **트레이드오프(의도된 변경)**: 세로 중첩이라 **행이
  자기완결이 아니다** — 직급 값이 부모(`<직급그룹>`)에만 있어 행 하나만 떼면 직급을 모른다. v1 의
  자기완결성을 포기하고 계층 보존을 택한 **사용자 확정 형식**이다(spec §7). **세로를 다시 매 행
  forward-fill 로 평면화하지 말 것.**
- **자동 XML→표 역변환은 없다 (의도적)**: 워처(`_converted_clipboard`)는 클립보드의 XML 을
  표로 되돌리지 **않는다**. XML 은 오직 메뉴의 `copy_as_xml` 클릭으로만 생성된다. (초기 0.3.0
  엔 자동 역변환 분기가 있었으나 사용자 요청으로 제거 — 일반·설정 XML 을 자동으로 건드릴 위험
  회피 + 메뉴 단순화["XML 변환 사용" 토글도 함께 삭제]. 되살리려면 `_converted_clipboard` 에
  `is_table_xml`/`table_xml_to_markdown` 분기를 다시 넣으면 된다.) 단, `copy_as_xml` 의 소스
  추출(`_clipboard_table_model`)은 클립보드의 (새 형식) XML 도 표로 인정하므로, 거기 쓰는
  `is_table_xml` 은 여전히 **보수적**이어야 한다 — config/문서/임의 XML 오인 금지. 가드(spec §6):
  루트의 직접 자식이 모두 `행` 또는 `{X}그룹`(재귀적으로 `행` 포함) + `행` 의 자식은 `열`
  또는 `열그룹`(재귀)만이고 `열` 은 leaf(자식 없음) + (`행` 2개 이상 **또는** 루트가 알려진
  표 루트(`표`,`dataset`,`table`…)·행 태그가 알려진 것). **이 가드를 느슨하게 풀지 말 것.**
- **표 → XML (메뉴 `copy_as_xml`)**: 이건 **사용자의 명시적 동작**이라 불변식 3(html 유지)과
  **일부러 다르게** 동작한다 — text 슬롯에 XML 을 넣고 **HTML 을 drop** 한다
  (`drop_types=RENDERED_TABLE_TYPES | HTML_TYPES`). 그래야 어디에 붙여도 "표가 아니라 XML"이
  나온다. 불변식 3 은 **자동 변환(`_converted_clipboard`)** 에만 적용되는 규칙이므로, 이
  메뉴 동작의 html drop 을 "회귀"로 보고 되돌리지 말 것.
- **병합 셀은 XML 경로 전용 `html_table_to_model` 이 처리** (마크다운 경로 `html_table_to_markdown`
  은 건드리지 말 것 — Markdown 은 병합/계층 표현 불가라 포기). XML 경로는: ① rowspan 값을 아래로
  채우고(forward-fill), ② 전체 열 병합 제목 행(단일 cell colspan=전체)은 건너뛰고, ③ 다단
  그룹 헤더는 `<th>`/`<thead>` 가 있으면 그걸로, 없으면(=실제 Excel 은 전부 `<td>`) **colspan 으로
  추론**해(상단의 가로병합 있는 행들 + 그 아래 leaf 한 줄) **헤더 레벨들을 분리 보존**한다 →
  `model_to_xml` 이 `<열그룹>` 으로 중첩(평면 결합 금지). 실제 Excel 은 `th` 를 안 쓰므로 **colspan
  기반 헤더 추론을 제거하면 다단 헤더가 다시 깨진다.** (`html_table_to_rows` 는 이 모델을 한 줄
  헤더로 평면 결합한 뷰 — 마크다운·빈칸 채우기용.)
- **‘XML: 빈칸을 자동 채우기’ 옵션(`forward_fill_key_columns`, 기본 꺼짐, `settings.py`/NSUserDefaults
  영속)**: 병합을 안 하고 빈칸으로 그룹을 표현한 표(직급을 그룹 첫 행에만 쓰고 아래는 비움 —
  현실에서 흔함)를 위해, 클릭 변환(`copy_as_xml`→`_clipboard_table_model`) 시 **왼쪽 키 열의 빈칸만**
  채운다 — ① 먼저 **위(세로)** 값으로, ② 그래도 빈 칸은 **좌측(가로)** 값으로(병합의 rowspan→위·
  colspan→좌측 origin 과 같은 원리). 가드: 열을 왼쪽→오른쪽으로 보다 **빈칸 없는(꽉 찬) 열을 만나면 멈춤** — 그
  오른쪽 값 열의 빈칸은 "진짜 없음"일 수 있어 건드리지 않는다(pandas·Power Query 관례: 그룹 열만
  채우고 값 열은 보존). **이 가드를 빼고 전체를 채우면 값 열 빈값이 왜곡된다.** 기본 꺼짐도 같은
  이유(안전 — 데이터를 임의 생성하지 않음). 마크다운 경로엔 적용 안 함(병합/빈칸 표현 불가).

### 회귀 방지 테스트 (변환 로직 수정 후 반드시 통과)
`scripts/run_test_matrix.py` 의 converter 테스트군:
- `html_table_with_markdown_text_preserved` → 불변식 1
- `html_clipboard_keeps_html` → 불변식 3 (순수 Excel 표도 html 유지)
- `html_table_in_document_augments_text`, `html_table_in_document_keeps_html` → 불변식 4
- `is_table_xml_rejects_config`, `config_xml_left_untouched` → 불변식 5 (XML false positive 차단)
- `table_xml_not_auto_converted` → 불변식 5 (워처가 XML 을 표로 자동 변환하지 않음)
- `model_to_xml_roundtrip` → 불변식 5 (모델↔XML 무손실 — 세로 중첩 포함)
- `model_to_xml_header_in_attribute` → 불변식 5 (가로 헤더는 `n=`/`이름=` 속성, 가로 header→tag 금지)
- `merged_excel_to_xml_nests_group_header`, `merged_excel_to_xml_keeps_subheader_cell` → 불변식 5
  (가로 다단 헤더를 `<열그룹>` 으로 중첩 보존)
- `forward_fill_fills_left_key_column`, `forward_fill_stops_at_data_column` → 불변식 5
  (빈 칸 채우기: 키 열만 채우고 값 열은 보존)
- `forward_fill_horizontal_left` → 불변식 5 (위가 비면 좌측 값으로 가로 채움)

실행 (시스템 클립보드 안 건드리는 순수 변환 테스트만):
```bash
.venv/bin/python -c "import sys; sys.path.insert(0,'scripts'); \
from run_test_matrix import run_converter_tests; \
r=run_converter_tests(); print(sum(x.ok for x in r),'/',len(r),'passed'); \
[print('FAIL',x.name,x.detail) for x in r if not x.ok]"
```

---

## 설정 영속성 정책 (UX 결정 — 2026-06-10 확정)

| 상태 | 영속 | 근거 |
|------|------|------|
| 변환 토글(`enabled`) | **비영속** (매 실행 켜짐) | 토글은 "일시정지" 용도. 영속화하면 "몇 주 전에 꺼둔 걸 잊고 고장으로 오인"하는 사고가 더 흔함. macOS·Windows 동일. |
| `fill_blanks`, 언어, Pro 캐시 | 영속 (NSUserDefaults / `%APPDATA%` JSON) | 명시적 환경설정. |
| `welcome_shown` (첫 실행 안내 1회) | 영속 | 알럿 표시 **전에** 마킹 — 실패해도 매 실행 반복 금지. |
| 로그인 시 자동 실행 | 영속 — **OS 가 보관**, 우리 설정 파일엔 안 씀 | macOS: SMAppService(시스템 로그인 항목). Windows: MSIX StartupTask(설정 ▸ 앱 ▸ 시작 프로그램). 둘 다 OS 가 진실의 원천이라 JSON/NSUserDefaults 그림자를 두지 말 것(드리프트 발생). |

- `enabled` 를 영속화하자는 제안이 다시 나오면 위 근거와 함께 재논의할 것 (실수로 "정리"하지 말 것).
- Windows 설정 저장은 `tabledown_windows/settings.py` 의 **read-modify-write** 만 사용 — 통째로 덮어쓰면 다른 키가 소실된다(과거 i18n 저장기가 그랬음).

### 로그인 시 자동 실행 (macOS `login_item` ↔ Windows `startup_task`)
- **macOS**: `tablemark/login_item.py`(SMAppService). **Windows**: `windows/tabledown_windows/startup_task.py`
  (WinRT `StartupTask`, `winsdk`). 둘 다 `is_supported()`/`is_enabled()`/`set_enabled()` 형태로,
  미지원 환경에선 `is_supported()` False → 메뉴에서 항목 숨김(graceful fallback). 동형 유지할 것.
- **Windows TaskId 는 매니페스트와 일치해야 함**: `startup_task.TASK_ID = "TabledownStartup"` 은
  `windows/packaging/AppxManifest.xml` 의 `<uap5:StartupTask TaskId="TabledownStartup" Enabled="false">`
  와 정확히 같아야 한다. 매니페스트는 `Enabled="false"`(기본 꺼짐) — 토글/설정으로만 켜짐.
- **한 operation = 한 전용 스레드**(`_run` → `asyncio.run`): 각 공개 함수(`current_state`/`set_enabled`)는
  get→토글→read-back 을 **하나의 코루틴**으로 묶어 **하나의 워커 스레드**에서 실행한다. 이유 ①: pystray 펌프
  스레드는 Windows 메시지 루프를 돌리는데 거기서 직접 WinRT async 를 await 하면 COM apartment/메시지 펌프와
  얽혀 교착될 수 있다(호출자는 `join()` 으로 잠깐 블록만). 이유 ②: 한 스레드(한 apartment 수명) 안에서 task
  프록시를 만들고 쓰므로 **WinRT 객체를 스레드 간 공유하지 않는다**. **enable 은 워커, disable 은 호출자 스레드**
  식으로 나누지 말 것 — `disable()` 도 코루틴 안에 둬 같은 스레드에서 호출한다(과거 리뷰 지적).
- **`asyncio.run` 은 코루틴만 받는다**: winsdk `IAsyncOperation` 은 awaitable 이지만 코루틴이 아니라 바로
  넘기면 `ValueError`. 작업을 `async def`(`_read_state_coro`/`_set_enabled_coro`) 안에 둬 워커 스레드에서 생성·await 한다.
- **거짓 비활성(false-disable) 감지**: Windows 는 사용자가 작업 관리자에서 끈 항목(`disabled_by_user`)·
  정책 제어(`disabled_by_policy`)는 `request_enable_async()` 로도 못 켠다. `set_enabled()` 가 read-back
  상태명을 돌려주므로 체크마크가 거짓으로 켜지지 않고, 앱이 안내 알림을 띄운다. **enable 결과를 무조건
  성공으로 가정하지 말 것.**

## UI 관례 (메뉴·피드백)

- 메뉴 순서: 동작(토글·XML) → 설정 → 후원 → 도움말·종료. 후원을 설정 위로 올리지 말 것(첫인상).
- '구매 복원'은 후원 서브메뉴 하단(결제 관련 항목 한곳에). Apple 요구는 "존재"이지 위치가 아님.
- 시스템 알림(UNUserNotification)은 쓰지 않는다 — 권한 프롬프트가 떠서 "권한 0개" 셀링포인트가 깨짐.
  성공 피드백은 **메뉴바 아이콘 1초 체크 플래시**(`_flash_icon_success`, 에셋은 `scripts/make_menu_icons.py` 생성).
- 가격은 하드코딩 금지: SKProduct 의 localized price 를 메뉴/시트에 동적 표기, `PRO_PRICE_FALLBACK` 은
  메타데이터 로딩 전 임시 표기 전용(ASC 가격과 일치 유지).
- 구매 흐름 피드백 규칙: 구독 성공 알럿은 **사용자가 방금 시작한 구독에만**(자동 갱신·복원은 침묵),
  사용자 취소(SKError code 2)는 침묵, 그 외 실패·복원 결과는 알럿.

## 빌드 / 실행

- **로컬 테스트**: 소스 직접 실행 `.venv/bin/python run.py`
  (iCloud Drive 폴더라 `build_release.sh` 의 ad-hoc 빌드는 `com.apple.FinderInfo`
  xattr 때문에 codesign 검증이 실패함 — 로컬 동작 확인엔 소스 실행이 가장 확실.)
- **App Store(.pkg)**: `TABLEDOWN_BUILD=<build> bash scripts/build_app_store.sh`
  (build number 는 marketing version 과 무관하게 **직전 업로드보다 높아야** 함.
  자동 탐지된 MAS provisioning profile 사용.)
- **GitHub 배포(DMG)**: `NOTARY_PROFILE=tabledown-notary bash scripts/build_dmg.sh`
  (Developer ID 서명 + Apple 공증 + staple. 앱과 DMG **둘 다** 공증해야 함 — DMG 만
  staple 하면 "Record not found" 로 실패.)
- **Windows 빌드는 `--collect-all winsdk` 필수**(`windows/build_windows.ps1`): `startup_task` 가 쓰는
  `winsdk` 는 namespace 모듈을 lazy import 하고 코드가 native `_winrt.pyd` 에 있어 PyInstaller 정적 분석이
  둘 다 놓친다. 빠지면 빌드는 통과하지만 **로그인 토글이 패키지 빌드에서 조용히 사라진다**(import 실패 →
  `is_supported()` False). 검증: frozen exe 에서 `winsdk.windows.applicationmodel.StartupTask` import 가 되고
  미패키지 실행 시 `OSError(ERROR_NOT_FOUND, -2147023728)` 로 깨끗이 떨어지면 OK.

## 버전 / 릴리스 관례
- 버전: `tablemark/__init__.py` 의 `__version__`. SemVer.
- CHANGELOG(`CHANGELOG.md`) 와 README 변경 이력(한 `README.md` / 영 `README.en.md`) 둘 다 갱신.
- git 태그 `vX.Y.Z`, GitHub Release 는 최신 버전을 Latest 로.
- README 의 DMG 다운로드 링크는 `releases/latest/download/Tabledown.dmg` —
  **Latest 릴리스에 DMG 에셋이 반드시 있어야** 404 가 안 난다.

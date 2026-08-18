# COA Downloader 이식 가이드

이 폴더의 `coa_downloader.py`는 Abcam, Sigma-Aldrich, TCI, Thermo Fisher의 품질 문서를 다운로드하고 PDF 본문을 검증하는 독립 모듈이다.

## 1. 지원 범위

| 회사 | 입력 | 저장 문서 | 반환 상태 |
|---|---|---|---|
| TCI | Catalog + Lot | COA | `downloaded` |
| Sigma-Aldrich | Catalog + Lot | COA PDF 또는 HTML COA 보관 PDF | `downloaded` |
| Thermo Fisher | Catalog + Lot | COA | `downloaded` |
| Abcam | Catalog + Lot | Datasheet + CoC | `fallback` |

Abcam의 Datasheet와 CoC는 COA가 아니다. 모듈은 이를 COA로 이름을 바꾸지 않으며 `fallback` 상태와 경고를 반환한다.

## 2. 파일

```text
coa_portable_package/
├── coa_downloader.py   # 다른 프로그램에 복사할 구현 모듈
└── README.md           # 설치 및 연동 설명
```

## 3. Python 요구사항

- Python 3.10 이상 권장
- `requests`
- `PyMuPDF`
- `seleniumbase`
- Chrome 또는 Chromium 계열 브라우저

설치 예:

```bash
pip install requests pymupdf seleniumbase
```

모듈 내부에서는 PyMuPDF를 다음 이름으로 가져온다.

```python
import pymupdf
```

환경에 따라 `fitz` 이름만 제공되는 구버전 대신 최신 `pymupdf` 패키지를 사용하는 것이 좋다.

## 4. 브라우저 컨텍스트 요구사항

공개 함수의 `context`에는 SeleniumBase의 `SB` 객체를 전달하는 것을 권장한다.

모듈에서 사용하는 주요 메서드는 다음과 같다.

```text
get(url)
get_title()
execute_script(script)
execute_async_script(script, timeout=...)
wait_for_element_visible(selector, timeout=...)
type(selector, value)
click(selector)
get_attribute(selector, attribute)
driver.execute_cdp_cmd(command, params)
```

마지막 CDP 메서드는 Sigma-Aldrich가 COA를 PDF가 아닌 웹페이지로 반환할 때 `Page.printToPDF`를 실행하기 위해 필요하다.

## 5. 기본 사용법

```python
from pathlib import Path
from seleniumbase import SB

from coa_downloader import download_quality_documents

output_dir = Path("downloaded_coa")

with SB(uc=True, headless=True) as sb:
    result = download_quality_documents(
        context=sb,
        vendor="TCI",
        catalog="C0119",
        lot="6JK3O",
        output_dir=output_dir,
    )

print(result)
```

## 6. COA만 허용하는 사용법

```python
from coa_downloader import download_coa, QualityDocumentError

try:
    coa_path = download_coa(
        context=sb,
        vendor="Sigma-Aldrich",
        catalog="M6250",
        lot="BCCL9247",
        output_dir="downloaded_coa",
    )
except QualityDocumentError as exc:
    print(f"COA를 확보하지 못했습니다: {exc}")
```

`download_coa()`는 실제 `document_type == "COA"` 파일만 반환한다. Abcam처럼 Datasheet와 CoC만 있는 경우에는 대체 문서를 COA로 반환하지 않고 예외를 발생시킨다.

## 7. 반환 형식

정상 COA 예:

```python
{
    "vendor": "TCI",
    "catalog": "C0119",
    "lot": "6JK3O",
    "status": "downloaded",
    "documents": [
        {
            "document_type": "COA",
            "path": ".../TCI_C0119_6JK3O_COA.pdf",
            "source_url": "https://...",
            "catalog": "C0119",
            "lot": "6JK3O",
            "verified": True,
        }
    ],
    "warnings": [],
}
```

Abcam 대체 문서 예:

```python
{
    "vendor": "Abcam",
    "catalog": "AB48394",
    "lot": "1147669-2",
    "status": "fallback",
    "documents": [
        {"document_type": "Datasheet", "path": "...", "verified": True},
        {"document_type": "CoC", "path": "...", "verified": True},
    ],
    "warnings": ["CoC는 시험 결과가 기재된 COA와 동일한 문서가 아닙니다."],
}
```

## 8. 회사명 입력값

다음 문자열을 지원한다. 공백과 하이픈은 내부에서 제거해 비교한다.

```text
TCI
Aldrich
Sigma-Aldrich
SigmaAldrich
Sigma
ThermoFisher
Thermo Fisher
Thermo
Abcam
```

## 9. 회사별 핵심 로직

### TCI

1. `https://www.tcichemicals.com/KR/ko/p/{catalog}`를 연다.
2. `#LotNumbere`에 Lot를 입력한다.
3. `button.js-cofa`를 클릭한다.
4. 페이지의 XHR 또는 `window.open`이 반환하는 PDF Blob을 가로챈다.
5. Blob을 Base64로 변환해 Python으로 전달한다.
6. PDF 본문에서 catalog, lot, COA 문구를 확인하고 저장한다.

TCI는 일반 다운로드 링크가 아니라 브라우저에서 생성되는 Blob을 사용하므로 `requests.get()`만으로 처리하면 실패할 수 있다.

### Sigma-Aldrich

1. `sigma`, `aldrich`, `sial`, `supelco`, `mm` 브랜드 경로를 순서대로 확인한다.
2. 제품 페이지의 COA Lot 목록에서 일치하는 Lot를 찾는다.
3. 목록에 없으면 공식 `documents-search?tab=coa` 화면에서 다시 검색한다.
4. 숫자 Lot가 10자리보다 짧으면 선행 0을 붙인 10자리 후보도 조회한다.
5. 반환값이 PDF이면 브라우저 세션 또는 HTTP 요청으로 저장한다.
6. 반환값이 공식 `/documents/coa/` HTML 페이지이면 Chrome CDP로 PDF를 생성한다.

예:

```text
입력 Lot: 494416
공식 Batch Number: 0000494416
```

HTML COA 보관 PDF에는 catalog, lot, 보관 시각, 원본 URL과 페이지 번호가 들어간다.

### Thermo Fisher

제품 페이지와 동일한 출처에서 공식 인증서 API를 호출한다.

```text
/api/store/Assets/Documents/Certificates/v2/search
  ?skus={catalog}
  &country=kr
  &targetSite=TF
  &partialLotNumber=true
  &erpType=Global_E1
  &lotNumbers={lot}
```

`lotNumbers`의 대문자 `N`이 중요하다. `lotNumber` 또는 `lotnumbers`를 사용하면 Lot 필터가 적용되지 않을 수 있다.

API 응답에서 다음 조건이 모두 일치하는 asset만 사용한다.

- `lotNumber == 입력 Lot`
- `sku`에 입력 Catalog 포함
- `documentType == Certificate of Analysis`

### Abcam

1. 제품 페이지에서 `.ds-button`을 클릭해 Datasheet URL을 얻는다.
2. `.coc-button`을 클릭한다.
3. `input.lotNumber-input`에 Lot를 입력한다.
4. 생성된 CoC 다운로드 링크를 읽는다.
5. Datasheet와 CoC를 서로 다른 파일명과 문서 유형으로 저장한다.

CoC에서는 catalog, lot 및 다음 문구 중 하나를 확인한다.

- Certificate of Conformance
- Certificate of Conformity
- Certificate of Compliance

## 10. PDF 검증 기준

다운로드 성공은 HTTP 상태만으로 결정하지 않는다.

1. 응답 크기가 1,000바이트 이상이어야 한다.
2. 파일이 `%PDF`로 시작해야 한다.
3. PyMuPDF가 한 페이지 이상 열 수 있어야 한다.
4. 추출 본문에 catalog가 있어야 한다.
5. Lot 문서는 추출 본문에 lot가 있어야 한다.
6. COA는 Certificate of Analysis 계열 문구가 있어야 한다.

비교 시 공백, 하이픈, 기타 기호와 대소문자 차이는 제거한다.

## 11. 다른 프로그램에 삽입하는 방법

### 단일 모듈로 복사

`coa_downloader.py`를 대상 프로젝트의 모듈 폴더에 복사한다.

```text
your_app/
├── app.py
└── integrations/
    └── coa_downloader.py
```

```python
from integrations.coa_downloader import download_quality_documents
```

### 기존 SeleniumBase 세션 재사용

제품 정보 수집에 이미 SeleniumBase를 사용한다면 같은 `SB` 객체를 전달한다. 제조사 사이트의 쿠키와 봇 방지 세션을 재사용할 수 있어 가장 안정적이다.

```python
with SB(uc=True, headless=True) as sb:
    product = scrape_product(sb, catalog)
    quality = download_quality_documents(sb, vendor, catalog, lot, coa_dir)
```

### GUI 또는 작업 큐에서 호출

다운로드 작업은 네트워크와 브라우저를 사용하므로 GUI 메인 스레드에서 직접 호출하지 말고 작업 스레드 또는 큐에서 실행하는 것이 좋다.

```python
def coa_job(context, vendor, catalog, lot, output_dir):
    try:
        return download_quality_documents(
            context, vendor, catalog, lot, output_dir
        )
    except QualityDocumentError as exc:
        return {
            "status": "error",
            "vendor": vendor,
            "catalog": catalog,
            "lot": lot,
            "error": str(exc),
        }
```

## 12. 저장 파일명

```text
{Vendor}_{Catalog}_{Lot}_COA.pdf
Abcam_{Catalog}_{Lot}_Datasheet.pdf
Abcam_{Catalog}_{Lot}_CoC.pdf
```

파일명에 사용할 수 없는 문자와 공백은 `_`로 변경한다.

## 13. 오류 처리

공통 예외 클래스는 `QualityDocumentError`다.

주요 실패 원인:

- 지원하지 않는 제조사
- 제품 페이지 또는 문서 컨트롤 변경
- 해당 Lot 문서 미등록
- 브라우저 fetch/XHR 실패
- PDF가 아닌 HTML 오류 응답
- PDF 본문에서 catalog 또는 lot 불일치
- COA 문구 누락

실패한 응답을 PDF 파일로 남기지 않고 예외로 처리하는 것이 원칙이다.

## 14. 감사 추적 권장 항목

다른 프로그램의 데이터베이스에는 다음 정보를 함께 기록하는 것이 좋다.

```text
vendor
input_catalog
input_lot
resolved_lot
document_type
local_path
source_url
downloaded_at
verified
status
warnings
```

현재 반환 데이터에는 `resolved_lot` 필드가 별도로 없으므로, Sigma-Aldrich 선행 0 보정 이력을 엄격히 관리해야 한다면 호출부 또는 모듈 반환 구조에 이 필드를 추가하는 것이 좋다.

## 15. 운영상 주의사항

- 제조사 페이지의 DOM과 API는 예고 없이 바뀔 수 있다.
- 선택자 변경 시 실제 PDF를 내려받아 catalog와 lot 검증까지 회귀 시험한다.
- 서명된 Sigma-Aldrich URL에는 만료 시간이 있으므로 URL보다 로컬 검증 PDF를 보관한다.
- Abcam `fallback`을 COA 충족으로 자동 승인하지 않는다.
- 인증 또는 시험법 문서 절차에서 요구하는 문서 유형은 품질 담당자가 결정해야 한다.
- 동시 다운로드 시 회사별 브라우저 세션과 요청 간격을 제한하는 것이 좋다.

## 16. 실제 검증 사례

| 회사 | Catalog | Lot | 결과 |
|---|---|---|---|
| Abcam | AB48394 | 1147669-2 | Datasheet + CoC |
| Abcam | AB6789 | 1133456-9 | Datasheet + CoC |
| Sigma-Aldrich | I6256 | 494416 | COA, `0000494416`으로 보정 |
| Sigma-Aldrich | M6250 | BCCL9247 | COA |
| TCI | C0119 | 6JK3O | COA |
| TCI | A1252 | T2UGO | COA |
| Thermo Fisher | M30550 | 2291655 | COA |
| Thermo Fisher | 21578 | 3412368 | COA |

## 17. 권장 확장

다른 프로그램에서 장기적으로 사용할 경우 다음 개선을 권장한다.

1. `resolved_lot`, `downloaded_at`, SHA-256 필드를 반환값에 추가한다.
2. 회사별 어댑터를 별도 클래스로 분리한다.
3. 다운로드 재시도와 지수 백오프를 추가한다.
4. 동일 catalog/lot의 검증 완료 파일 캐시를 추가한다.
5. 제조사별 통합 테스트와 샘플 PDF 회귀 테스트를 자동화한다.
6. 로그에 서명 URL 전체를 남길 때 토큰 노출 여부를 검토한다.

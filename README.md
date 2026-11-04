# 손익계산서 IFRS 18 변환

기존 Excel 손익계산서를 읽어 IFRS 18 범주로 분류한 새 손익계산서를 만듭니다. 애매한 계정은 OpenAI API가 분류를 추천하고, 바로 옆에 판단 근거와 필요한 증빙을 적습니다. 원본 파일과 원본 시트는 보존합니다.

화면 프로그램 없이 명령 하나로 실행합니다. Python 3.11 이상을 사용하세요. Windows / Python 3.12에서 실행 확인했습니다.

## 설치

압축을 풀고 해당 폴더의 PowerShell에서 실행합니다.

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item .env.example .env
notepad .env
```

`.env`에 OpenAI 인증키와 계정에서 사용 가능한 Structured Outputs 지원 모델명을 넣습니다. OpenDART 키와는 다른 키입니다.

```dotenv
OPENAI_API_KEY=발급받은_OpenAI_API_키
OPENAI_MODEL=사용할_모델명
```

키가 없거나 호출에 실패해도 Excel 생성은 계속합니다. 해당 계정은 ‘판단 보류’와 필요한 자료를 표시합니다. 키·API 예외 원문은 출력하지 않습니다. AI를 사용하면 계정명·금액·설명·주석·사업활동이 OpenAI API로 전송됩니다. 원본 통합문서 전체를 업로드하지 않습니다.

## 실행

입력 Excel에 `계정과목`, `당기`, `전기`, `주석`처럼 명확한 열 제목이 있으면 자동으로 찾습니다. 입력의 금액은 원 단위, 수익은 양수·비용은 음수로 준비합니다.

```powershell
.\.venv\Scripts\python.exe convert.py "기존손익계산서.xlsx" --company "기업명" --year 2025 --activities "자동차 부품 제조 및 판매" --output "변환결과.xlsx"
```

인식이 안 되거나 표가 여러 개라면 시트·헤더행·열을 지정합니다.

```powershell
.\.venv\Scripts\python.exe convert.py "기존손익계산서.xlsx" --company "기업명" --year 2025 --activities "제품 제조 및 판매" --sheet "공시 IS" --header-row 11 --account-col C --current-col D --output "변환결과.xlsx"
```

전기 금액 열이 E라면 `--previous-col E`를 추가합니다. 백만원 단위 표는 `--scale 1000000`, 비용을 양수로 적은 표는 `--expenses-positive`를 추가합니다. 이 옵션은 원가·비용·손실 명칭을 음수로 바꾸므로 금액 부호를 반드시 확인하세요. 원래 부호가 있으면 기본 옵션을 사용합니다.

API 호출 없이 확인:

```powershell
.\.venv\Scripts\python.exe convert.py examples/input_example.xlsx --company "가상제조" --year 2025 --activities "제품 제조 및 판매" --no-ai --output "가상제조_변환.xlsx"
.\.venv\Scripts\python.exe -m pytest -q
```

입력에 계산 수식이 있지만 Excel이 저장한 계산값이 없다면 임의 계산하지 않고 입력 확인으로 표시합니다. 원본을 Excel에서 재계산·저장하거나 숫자로 된 입력표를 사용하세요.

## 결과 확인

기존 시트들에 다음 네 시트가 추가됩니다.

- **계정재분류**: 규칙 분류, AI 추천, 판단 근거, 문단, 원천자료, 필요한 증빙을 나란히 표시합니다. 노란색 M열에서 최종 분류를 선택하고 N열을 ‘확인’으로 바꾸면 반영됩니다.
- **공시IS_IFRS18**: 영업·투자·재무·법인세·중단영업과 중간합계, 당기순손익을 작성합니다. OCI는 별도 표시합니다. 미분류 금액도 별도로 보여줍니다.
- **검증결과**: 원본 순손익과 변환 순손익, 원본 상세 합계, 미분류·미검토 계정을 확인합니다.
- **분류근거**: 적용한 규칙과 문단을 정리합니다.

AI 추천이 손익계산서에 반영되어 있어도 사용자가 확인하기 전에는 **초안**으로 표시합니다. 근거가 없으면 분류를 만들지 않고 ‘판단 보류’로 남깁니다. 예를 들어 외환차익은 매출채권·차입금 등 발생 원천별 명세가 필요하다고 표시합니다.

모든 상세 금액과 당기순손익 통제 금액이 있고, 미검토 항목이 없고, 손익 대사가 맞을 때만 검토 완료로 표시합니다. 원본 중간합계와 당기순손익은 상세 계정에 중복 합산하지 않습니다. 부모 합계와 세부 내역이 함께 있는 표는 한쪽만 입력하거나 행유형 열에 `subtotal`을 지정하세요.

Excel 수식은 파일을 열 때 재계산됩니다. openpyxl 자체는 수식을 실행하지 않으므로 수식을 계산하지 않는 미리보기에서는 비어 보일 수 있습니다. M/N열 변경은 Excel의 계산식으로 손익계산서와 검증에 반영됩니다. 원천 설명·금액을 바꿔 AI 판단을 새로 받으려면 원본을 고친 뒤 새 파일명으로 다시 실행하세요.

## 입력의 선택 열

`기존범주`, `설명`, `주석`, `원천유형`, `행유형`을 추가하면 더 정확히 읽습니다. 기존 영업·영업외 구분은 새 분류를 확정하는 근거로 쓰지 않습니다.

원천유형은 경제적 실질을 확인한 경우만 입력합니다. 예: `operating_ppe`(영업에 쓰는 설비), `equity_method`(지분법), `loan_funding`(일반 차입), `lease_interest`(리스 이자), `trade_receivable_fx`(매출채권 외환), `loan_fx`(차입금 외환), `cash_equivalent`, `independent_asset`, `customer_funding`. 설명이나 주석 없이 원천코드만 적으면 분류를 보류합니다.

주된 사업활동은 반드시 입력합니다. 금융제공·투자가 주된 사업이면 `--finances-customers`, `--invests-assets`, `--invests-financial-assets`를 사용할 수 있습니다. 금융업은 명칭만으로 관련 부채·현금·정책 선택을 확정하지 않고 추가 확인 대상으로 남깁니다.

복수 원천이 섞인 계정은 원본에서 원천별 행으로 나눠 입력하고 합계행은 `subtotal`로 표시하세요. 단일 비조달 부채에서 발생한 외환차이를 임의로 복수 범주에 배분해서는 안 됩니다(문단 B67). 이 간단 버전에는 별도의 금액 분할 화면을 두지 않았습니다.

## 코드와 예제

`convert.py`가 실행 파일입니다. `ifrs18/parser.py`가 Excel을 읽고, `engine.py`와 `rules.json`이 규칙 분류를 하며, `ai.py`가 애매한 계정의 구조화된 추천을 검증합니다. 수식 생성과 Excel 저장은 `convert.py`에 모았습니다.

`examples/input_example.xlsx`와 `output_example.xlsx`는 가상 자료입니다. 외환차익은 API 없이 실행했으므로 판단 보류인 것이 예상 결과입니다. 사용자 확인 전의 예제 결과를 완료된 실제 공시로 제시하지 마세요.

사용자가 지정한 무료 체험판의 ‘양식변동 예시’와 ‘공시 IS’를 분석해 영업·투자·재무 범주 및 중간합계 구조를 참고했습니다. 타인의 강의자료는 배포 ZIP에 포함하지 않았습니다. 기준서 합계는 원본 예제 수식을 그대로 복제하지 않고 새로 계산합니다.

## 기준과 범위

K-IFRS 제1118호 문단 47, 49, 52~68, 69~74, B65~B67과 국내 한132.4~한132.5를 확인했습니다. 영업손익과 투자손익을 합한 뒤 재무손익을 더하는 누적 구조로 작성합니다. 고객 금융제공 기업의 문단 73 예외 등 복잡한 정책 판단은 전문 검토가 필요합니다. 전체 표시·주석 요구사항 준수를 보증하는 도구는 아닙니다.

- [한국회계기준원 제정 공표](https://kasb.or.kr/front/board/comm010View.do?seq=2101)
- [금융위원회 국내 도입 안내](https://www.fsc.go.kr/no010101/85886)
- [OpenAI Structured Outputs](https://developers.openai.com/api/docs/guides/structured-outputs)

`.env`는 GitHub에 올리지 마세요. 이 버전의 확인 사항: 핵심 테스트 8개 통과. 실제 OpenAI API 호출은 키가 없어 수행하지 않았고, 구조화 응답·잘못된 원천 인용·실패 처리는 모의 응답으로 검사했습니다.

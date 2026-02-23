# 📊 xlwings 학습 가이드

> Python으로 Excel을 **실시간** 제어하는 자동화 라이브러리

---

## 📌 xlwings란?

1. 파이썬 엑셀 자동화 라이브러리
2. 엑셀을 직접적으로 제어할 수 있다 (실행 중인 Excel과 실시간 통신)
3. DRM 우회 가능 (보안 폴더도 우회 가능)

### ✅ 자동화 가능 목록

| # | 기능 | 설명 |
|:---:|:---|:---|
| 1 | 파일/시트 관리 | 생성, 수정, 저장 |
| 2 | 셀 데이터 조작 | 추가, 수정, 삭제 |
| 3 | 행/열 관리 | 생성, 삭제 |
| 4 | 스타일 변경 | 폰트, 배경색, 테두리, 정렬 |
| 5 | 데이터 취합 | 여러 시트/파일 → 하나로 합치기 |
| 6 | 복사/붙여넣기 | 값·서식·수식 복사 |
| 7 | 셀 병합 | 병합/병합해제 |
| 8 | 수식/PDF | 수식 입력, PDF 변환 |
| 9 | 대용량 처리 | 데이터 분석, 그래프 시각화 |

---

## 🧱 xlwings 구성요소 (계층 구조)

```
App (엑셀 프로그램)
 └─ Book (워크북 = .xlsx 파일)
     └─ Sheet (워크시트 = 탭)
         └─ Range (셀 범위 = 하나 또는 여러 개)
```

---

## 📚 챕터별 학습 내용

### 📗 01. 엑셀 파일 다루기 — 기초

> 📄 [01.엑셀파일다루기_기초.ipynb](./01.엑셀파일다루기_기초.ipynb)

**워크북(Book) 다루기**

```python
import xlwings as xw

app = xw.App(add_book=False)        # 엑셀 앱 실행 (빈 상태)
wb = app.books.add()                 # 새 워크북 생성
wb = app.books.open('파일경로')       # 기존 워크북 열기
wb.save('파일경로')                   # 다른 이름으로 저장
wb.save()                            # 저장
app.quit()                           # 엑셀 앱 닫기
```

**워크시트(Sheet) 다루기**

```python
wb.sheets.add('이름')                # 새 시트 생성
ws = wb.sheets['이름']               # 이름으로 시트 선택
ws = wb.sheets[0]                    # 인덱스로 시트 선택
ws.name = '변경할 이름'              # 시트 이름 변경
wb.sheets['이름'].delete()           # 시트 삭제
wb.sheets['이름'].activate()         # 시트 활성화
wb.sheets['이름'].clear()            # 시트 내용 전체 삭제
```

---

### 📗 02. 셀 다루기 — 기초

> 📄 [02.셀다루기_기초.ipynb](./02.셀다루기_기초.ipynb)

**셀 값 읽기/쓰기**

```python
ws.range('A1').value = '값'              # 값 입력
ws.range('A1').value                     # 값 읽기
ws.range('A1:D5').value                  # 범위 읽기 (2차원 리스트)
ws.range('A1').value = [[1,2],[3,4]]     # 2차원 배열 입력
```

**동적 범위 선택 (expand)**

```python
ws.range('A1').expand('table')    # 표 전체 범위 (아래+오른쪽)
ws.range('A1').expand('down')     # 아래로 데이터 있는 만큼
ws.range('A1').expand('right')    # 오른쪽으로 데이터 있는 만큼
```

> 💡 **팁**: `expand('table')`은 빈 행/열을 만나면 멈춥니다. 데이터 중간에 빈 행이 있으면 잘릴 수 있어요!

---

### 📗 03. 셀 서식 & 스타일링

> 📄 [03.셀서식_스타일링.ipynb](./03.셀서식_스타일링.ipynb)

**폰트 설정**

```python
ws.range('A1').font.name = '맑은 고딕'
ws.range('A1').font.size = 12
ws.range('A1').font.bold = True           # 굵게
ws.range('A1').font.italic = True         # 기울임
ws.range('A1').font.color = (255,0,0)     # 글자색 (RGB)
ws.range('A1').font.underline = True      # 밑줄
ws.range('A1').font.strikethrough = True  # 취소선
```

**배경색 설정**

```python
ws.range('A1').color = (255, 0, 0)    # RGB 색상
ws.range('A1').color = '#FF0000'      # HEX 색상
```

**테두리 설정 (API 방식)**

```python
# Borders 인덱스: 7=왼쪽, 8=위, 9=아래, 10=오른쪽
ws.range('A1:D5').api.Borders(7).LineStyle = 1   # 왼쪽
ws.range('A1:D5').api.Borders(8).LineStyle = 1   # 위쪽
ws.range('A1:D5').api.Borders(9).LineStyle = 1   # 아래쪽
ws.range('A1:D5').api.Borders(10).LineStyle = 1  # 오른쪽
```

**셀 크기 설정**

```python
ws.range('A1').column_width = 15     # 열 너비
ws.range('A1').row_height = 25       # 행 높이
```

**셀 병합/해제**

```python
ws.range('A1:D1').merge()     # 병합
ws.range('A1:D1').unmerge()   # 병합 해제
```

**정렬 설정 (API 상수)**

```python
# 가로 정렬
ws.range('A1').api.HorizontalAlignment = -4131   # 왼쪽
ws.range('A1').api.HorizontalAlignment = -4108   # 가운데
ws.range('A1').api.HorizontalAlignment = -4152   # 오른쪽

# 세로 정렬
ws.range('A1').api.VerticalAlignment = -4160     # 위쪽
ws.range('A1').api.VerticalAlignment = -4108     # 가운데
ws.range('A1').api.VerticalAlignment = -4107     # 아래쪽
```

**숫자 서식**

```python
ws.range('A1').number_format = '0.00'          # 소수점 둘째자리
ws.range('A1').number_format = '#,##0'         # 천 단위 구분
ws.range('A1').number_format = '0.00%'         # 백분율
ws.range('A1').number_format = 'yyyy/mm/dd'    # 날짜
ws.range('A1').number_format = '#,##0 "원"'    # 통화
```

---

### 📗 04. 반복 자동화 & 여러 시트 처리

> 📄 [04.반복자동화_여러시트처리.ipynb](./04.반복자동화_여러시트처리.ipynb)

**핵심 패턴: 모든 시트 순회**

```python
for sheet in wb.sheets:
    print(f'시트 이름: {sheet.name}')
    data = sheet.range('A1').expand('table').value
    # 각 시트별 처리 로직
```

**조건부 시트 처리 (특정 시트 제외)**

```python
skip_sheets = ['종합', '목차']
for sheet in wb.sheets:
    if sheet.name in skip_sheets:
        continue
    # 처리 로직
```

**시트별 요약 자동 생성**

```python
for sheet in wb.sheets:
    data = sheet.range('A2').expand('table').value
    if data:
        total = sum(row[2] for row in data if row[2])  # C열 합계
        sheet.range('F1').value = f'합계: {total}'
```

> 💡 **실무 핵심**: `for sheet in wb.sheets` 반복문이 업무 자동화의 90%를 차지합니다!

---

### 📗 05. 데이터 취합 & 복사/붙여넣기

> 📄 [05.데이터취합_복사붙여넣기.ipynb](./05.데이터취합_복사붙여넣기.ipynb)

**복사 방법 3가지 비교**

| 방식 | 코드 | 특징 |
|:---|:---|:---|
| **값만 복사** | `.value` 대입 | 가장 빠르고 단순, 서식 ✗ |
| **서식 포함** | `.api.Copy()` + `.api.PasteSpecial()` | VBA 방식, 서식+값 모두 ✓ |
| **pandas 연동** | `DataFrame` ↔ `Range` | 분석+출력 최강 콤보 |

**값만 복사**

```python
src = wb.sheets['마케팅팀']
dst = wb.sheets['영업1팀']

# 단일 셀
dst.range('A1').value = src.range('A1').value

# 범위 복사 (동적)
data = src.range('A2').expand('table').value
dst.range('A2').value = data
```

**수식 복사 vs 값 복사**

```python
ws.range('C6').value     # → 계산된 값 (예: 36.0)
ws.range('C6').formula   # → 수식 문자열 (예: '=SUM(C3:C5)')

# 수식 그대로 복사
dst.range('C6').formula = src.range('C6').formula
```

**서식 포함 복사 (API)**

```python
src.range('A2:D6').api.Copy()
dst.range('A2').api.PasteSpecial(Paste=-4104)  # 전체(값+서식)
app.api.CutCopyMode = False                    # 클립보드 정리
```

| PasteSpecial 상수 | 의미 |
|:---|:---|
| `-4104` | 전체 (xlPasteAll) |
| `-4163` | 값만 (xlPasteValues) |
| `-4122` | 서식만 (xlPasteFormats) |

**여러 시트 → 하나로 취합**

```python
def merge_sheets(wb, target='종합'):
    # 종합 시트 생성
    ws_target = wb.sheets.add(target)
    current_row = 1

    for sheet in wb.sheets:
        if sheet.name == target:
            continue
        data = sheet.range('A2').expand('table').value
        if data:
            ws_target.range(f'A{current_row}').value = data
            current_row += len(data) + 1
```

**pandas ↔ xlwings 연동**

```python
import pandas as pd

# Excel → DataFrame
df = ws.range('A1').expand('table').options(pd.DataFrame, header=1).value

# DataFrame → Excel
ws_result = wb.sheets.add('분석결과')
ws_result.range('A1').options(pd.DataFrame).value = df
```

> 💡 **실무 최강 콤보**: pandas로 데이터 분석 → xlwings로 서식 입혀서 출력!

---

## ⚠️ 자주 만나는 에러와 해결

| 증상 | 원인 | 해결 |
|:---|:---|:---|
| `data`가 `None` | 빈 시트에서 `expand()` | expand 전에 None 체크 |
| 단일값이 리스트가 아님 | 한 행이면 1차원 반환 | `if not isinstance(data[0], list)` 체크 |
| `COM Error` | Excel이 응답 없음 상태 | 작업관리자에서 Excel 종료 후 재실행 |
| 파일 열기 실패 | 다른 프로세스가 점유 | Excel에서 파일 닫기 |
| `PermissionError` | 파일이 읽기 전용 | DRM/보안 설정 확인 |
| 한글 깨짐 | 인코딩 문제 | `encoding='utf-8'` 지정 |

---

## 📚 학습 로드맵

```
01. 엑셀 파일 다루기 (기초)
 │   └─ App, Book, Sheet 개념
 ↓
02. 셀 다루기 (기초)
 │   └─ 값 읽기/쓰기, expand
 ↓
03. 셀 서식 & 스타일링
 │   └─ 폰트, 배경색, 테두리, 정렬
 ↓
04. 반복 자동화 & 여러 시트 처리
 │   └─ for 루프, 조건부 처리
 ↓
05. 데이터 취합 & 복사/붙여넣기
 │   └─ 복사 3가지 방법, pandas 연동
 ↓
06. 실전 자동화 프로젝트 (예정)
     └─ PDF 변환, 차트, 이메일, 스케줄링
```

---

## 🔗 참고 자료

- [xlwings 공식 문서](https://docs.xlwings.org/)
- [xlwings GitHub](https://github.com/xlwings/xlwings)
- [xlwings API Reference](https://docs.xlwings.org/en/stable/api/index.html)

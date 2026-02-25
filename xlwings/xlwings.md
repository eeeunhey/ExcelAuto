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

### 📗 06. 실전 자동화 프로젝트 — 차트 생성 & 스케줄링

> 📄 [06.실전자동화_차트_스케줄링.ipynb](./06.실전자동화_차트_스케줄링.ipynb)

**차트(Chart) 생성**

```python
# 차트 코드 3단계
# 1) 차트 객체 생성
chart = ws.charts.add(left=10, top=130, width=500, height=280)
# 2) 데이터 범위 연결
chart.set_source_data(ws.range('A1:C7'))
# 3) 종류 설정 ('line' / 'bar_clustered' / 'pie')
chart.chart_type = 'line'
```

**API 세부 설정**

```python
c = chart.api[1]      # COM 객체 접근
c.HasTitle = True
c.ChartTitle.Text = '월별 매출 현황'
c.Axes(2).HasTitle = True
c.Axes(2).AxisTitle.Text = '금액(만원)'
c.SeriesCollection(1).HasDataLabels = True
```

**스케줄링 (schedule 라이브러리)**

```python
import schedule, time

schedule.every().day.at('09:00').do(my_job)   # 매일 9시
schedule.every().monday.do(my_job)            # 매주 월요일
schedule.every(10).minutes.do(my_job)         # 10분마다

while True:
    schedule.run_pending()
    time.sleep(60)
```

> 💡 **실무 추천**: Windows 작업 스케줄러(`schtasks`)를 사용하면 Python이 실행 중이지 않아도 하이!

---

### 📗 07. pandas 심화 — 데이터 분석 & Excel 입출력

> 📄 [07.pandas_심화_데이터분석.ipynb](./07.pandas_심화_데이터분석.ipynb)

**조건 필터링**

```python
# 단순
df[df['선과점수'] >= 80]
# 복합 (AND &, OR |)
df[(df['부서'] == '영업팀') & (df['성과점수'] >= 80)]
# query 방식 (실주 추천)
df.query('성과점수 >= 80 and 교육시간 >= 15')
```

**groupby 집계**

```python
df.groupby('부서').agg(
    인원수=('이름', 'count'),
    평균교육시간=('교육시간', 'mean'),
    평균성과=('성과점수', 'mean')
).round(1).reset_index()
```

**피벗 테이블**

```python
pd.pivot_table(df, values='성과점수',
               index='부서', columns='등급',
               aggfunc='count', fill_value=0,
               margins=True, margins_name='합계')
```

**pandas ↔ xlwings 연동**

```python
# xlwings → DataFrame
df = ws.range('A1').expand('table').options(pd.DataFrame, header=1).value

# DataFrame → xlwings
ws.range('A1').options(pd.DataFrame, index=False).value = df
```

---

### 📗 08. openpyxl — 서식 완전 제어

> 📄 [08.openpyxl_서식완전제어.ipynb](./08.openpyxl_서식완전제어.ipynb)

**xlwings vs openpyxl**

| | xlwings | openpyxl |
|:---|:---:|:---:|
| Excel 실행 필요 | ✅ 기본 | ❌ 불필요 |
| 실시간 제어 | ✅ | ❌ |
| pandas 연동 | ✅ | ✅ (ExcelWriter) |
| 조건부 서식 | ⚠️ 제한적 | ✅ 주체 |
| 드롭다운/유효성 | ⚠️ | ✅ |
| 이미지 삽입 | ⚠️ | ✅ |

**서식 설정 패턴**

```python
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

cell = ws['A1']
cell.font   = Font(bold=True, color='FF0000', size=14)
cell.fill   = PatternFill(fill_type='solid', fgColor='FFFF00')
cell.border = Border(top=Side(style='thin'), bottom=Side(style='thin'))
cell.alignment = Alignment(horizontal='center', vertical='center')
cell.number_format = '#,##0"원"'
```

**조건부 서식**

```python
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

# 값 비교
ws.conditional_formatting.add('B2:B100',
    CellIsRule(operator='greaterThanOrEqual', formula=['90'],
               fill=green_fill, font=green_font))

# 컴러 스케일 (자동 그라데이션)
ws.conditional_formatting.add('C2:C100',
    ColorScaleRule(start_type='min', start_color='FF0000',
                   end_type='max',   end_color='00FF00'))
```

**드롭다운 (DataValidation)**

```python
from openpyxl.worksheet.datavalidation import DataValidation

dv = DataValidation(type='list', formula1='"A,B,C"')
ws.add_data_validation(dv)
dv.add('C2:C100')
```

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
06. 실전 자동화 프로젝트
     └─ 차트 생성, 스케줄링, PDF 변환
 ↓
07. pandas 심화
     └─ groupby, 피벗, 필터링, xlwings 연동
 ↓
08. openpyxl — 서식 완전 제어
      └─ 조건부 서식, 드롭다운, 이미지, 완성형 보고서
 ↓
09. 대용량 처리 & 실전 종합 심화
      └─ 청크 처리, 멀티파일 통합, 대시보드, 이메일 발송, 파이프라인
```

---

### 📗 09. 대용량 처리 & 실전 종합 심화

> 📄 [09.대용량처리_실전종합.ipynb](./09.대용량처리_실전종합.ipynb)

**대용량 CSV/Excel 처리 3가지 방법**

| 방법 | 코드 | 특징 |
|:---|:---|:---|
| 전체 읽기 | `pd.read_csv()` | 간단하지만 메모리 위험 |
| dtype 최적화 | `dtype={'열': 'category'}` | 메모리 50~70% 절약 |
| 청크 처리 | `chunksize=10_000` | 메모리 한계 초과할 때 |

**청크 처리 패턴**

```python
results = []
for chunk in pd.read_csv('big.csv', chunksize=10_000):
    agg = chunk.groupby('부서')['매출액'].sum()
    results.append(agg)

final = pd.concat(results).groupby(level=0).sum()
```

**멀티파일 자동 통합**

```python
from pathlib import Path

def merge_excel_files(folder: str, pattern: str = '*.xlsx') -> pd.DataFrame:
    files = sorted(Path(folder).glob(pattern))
    frames = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df['출처'] = f.stem   # 파일명 메타정보
            frames.append(df)
        except Exception as e:
            print(f'❌ {f.name}: {e}')  # 오류 파일 건너뛰기
    return pd.concat(frames, ignore_index=True)
```

**로깅 설정**

```python
import logging

logger = logging.getLogger('ExcelAuto')
logger.setLevel(logging.DEBUG)

# 파일 + 콘솔 동시 출력
fh = logging.FileHandler('automation.log', encoding='utf-8')
ch = logging.StreamHandler()
logger.addHandler(fh)
logger.addHandler(ch)

logger.info('자동화 시작')
logger.error('파일 없음: missing.xlsx')
```

**이메일 자동 발송 (Excel 첨부)**

```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

msg = MIMEMultipart()
msg['Subject'] = '월간 보고서'
msg['From']    = 'sender@gmail.com'
msg['To']      = 'recipient@example.com'

# 파일 첨부
with open('report.xlsx', 'rb') as f:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(f.read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="report.xlsx"')
msg.attach(part)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
    server.login('sender@gmail.com', 'app_password_16chars')
    server.sendmail('sender@gmail.com', 'recipient@example.com', msg.as_string())
```

> 💡 **Gmail 앱 비밀번호**: Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호

---

## 🔗 참고 자료

- [xlwings 공식 문서](https://docs.xlwings.org/)
- [xlwings GitHub](https://github.com/xlwings/xlwings)
- [xlwings API Reference](https://docs.xlwings.org/en/stable/api/index.html)

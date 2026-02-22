1. 파이썬 엑셀 자동화 라이브러리이다
2. 엑셀을 직접적으로 제어할 수 있다 
3. DRM 우회가능 (보안 폴더도 우회 가능)

- 자동화 가능 목록
1. 파일 및 시트 생성, 수정, 저장
2. 셀 데아터 추가, 수정, 삭제
3. 행 생성, 삭제
4. 스타일 변경
5. 취합하기
6. 복붙! 
7. 병합, 병합해제
8. 수식입력, pdf 변환
9. 대용량 데이터 처리 및 분석, 그래프 시각화 

xlwings의 구성요소

app - 프로그램  
book - 워크북 
sheet - 워크시트
range - 셀범위 (하나 또는 여러개)

워크북 다루기 명령어
app = xw.App(add_book = False) 엑셀 앱 만들기
wb = app.books.add() 엑셀 워크북 생성하기
wb = app.books.open(파일경로) 기존 워크북 불러오기
wb.save('파일 경로') 다른 이름으로 저장하기
wb.save() 저장하기
app.quit() 엑셀 앱 닫기

워크시트 다루기 명령어
wb.sheets.add('이름') 새로운 시트 생성하기
ws = wb.sheets['이름'] 이름으로 시트 선택하기
ws = wb.sheet[0] 인덱스로 시트 선택하기
ws.name = '변경할 이름'
wb.sheets['이름'].delete() 시트삭제
wb.sheets['이름'].activate() 시트 활성화
wb.sheets['이름'].clear() 시트 내용 전체 삭제



# 📊 Excel Automation Study

> Python을 활용한 엑셀 자동화 학습 기록

---

## 🎯 학습 목표

반복적인 엑셀 업무를 Python으로 자동화하기 위한 라이브러리 학습

---

## 🛠️ 학습 대상 라이브러리

| 라이브러리 | 역할 | 학습 상태 |
|:---:|:---|:---:|
| **xlwings** | Excel ↔ Python 실시간 연동 | 📝 학습 중 |
| **pandas** | 데이터 분석 및 가공 | ✅ 학습 완료 |
| **OpenPyXL** | `.xlsx` 파일 읽기/쓰기/서식 | ✅ 학습 완료 |
| **Pywin32** | Windows COM 인터페이스 | ⬜ 예정 |

---

## 📚 학습 노트

### 1. xlwings

> 📄 [xlwings 학습 가이드](./xlwings/xlwings.md)

| 챕터 | 주제 | 노트북 | 상태 |
|:---:|:---|:---|:---:|
| 01 | 엑셀 파일 다루기 (기초) | [01.엑셀파일다루기_기초.ipynb](./xlwings/01.엑셀파일다루기_기초.ipynb) | ✅ |
| 02 | 셀 다루기 (기초) | [02.셀다루기_기초.ipynb](./xlwings/02.셀다루기_기초.ipynb) | ✅ |
| 03 | 셀 서식 & 스타일링 | [03.셀서식_스타일링.ipynb](./xlwings/03.셀서식_스타일링.ipynb) | ✅ |
| 04 | 반복 자동화 & 여러 시트 처리 | [04.반복자동화_여러시트처리.ipynb](./xlwings/04.반복자동화_여러시트처리.ipynb) | ✅ |
| 05 | 데이터 취합 & 복사/붙여넣기 | [05.데이터취합_복사붙여넣기.ipynb](./xlwings/05.데이터취합_복사붙여넣기.ipynb) | ✅ |
| 06 | 실전 자동화 프로젝트 (차트 & 스케줄링) | [06.실전자동화_차트_스케줄링.ipynb](./xlwings/06.실전자동화_차트_스케줄링.ipynb) | ✅ |
| 07 | pandas 심화 — 데이터 분석 & Excel 입출력 | [07.pandas_심화_데이터분석.ipynb](./xlwings/07.pandas_심화_데이터분석.ipynb) | ✅ |
| 08 | openpyxl — 서식 완전 제어 | [08.openpyxl_서식완전제어.ipynb](./xlwings/08.openpyxl_서식완전제어.ipynb) | ✅ |

---

## ⚙️ 환경

- **Python** 3.9+
- **OS**: Windows
- **Microsoft Excel** 설치 필요

```bash
pip install xlwings pandas openpyxl pywin32
```

---

## 📖 참고 자료

- [xlwings 공식 문서](https://docs.xlwings.org/)
- [pandas 공식 문서](https://pandas.pydata.org/docs/)
- [OpenPyXL 공식 문서](https://openpyxl.readthedocs.io/)
- [Pywin32 GitHub](https://github.com/mhammond/pywin32)

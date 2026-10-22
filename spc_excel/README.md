# 기록별 SPC 엑셀 자동 생성

공정능력 HTML 에서 **[결과 저장]** 을 누르면, 공용 폴더에 아래 파일이 함께 만들어집니다.

```
공용 폴더\2026\
 ├─ 공정능력_결과기록.csv      결과 1건 = 1줄
 ├─ 공정능력_측정데이터.csv    측정값 1개 = 1줄
 ├─ 그래프\  기록ID_히스토그램.png, 기록ID_XbarR.png
 └─ SPC\     기록ID.xlsx        ← 완성된 SPC 엑셀 (기록마다 1개)
```

## SPC 엑셀
- 기존 SPC 파일(`127K_S-SC04C-_B__SPC_2026.xlsx`)의 시트 서식을 그대로 사용합니다
  (행 높이, 열 너비, 글꼴·크기, 테두리, 병합, 인쇄 영역, 제목 도형).
- 측정 Data(Lot 25개 × 20개), 평균(X)/범위(R), Specification, 통계치 Summary 가 값으로 채워집니다.
- 그래프는 사진으로 붙여넣습니다: 히스토그램 **세로 13cm × 가로 10cm** (E31), Xbar-R Chart **세로 13cm × 가로 24cm** (M31).
- 시트 이름: `조성 첫Lot~끝Lot` (예: `P-127-K 1033Y~1060Y`).
- '관리자 확인' 행은 비워 둡니다 (확인 후 직접 입력).
- Lot 이 25개를 넘으면 SPC 양식에 들어가지 않아 SPC 엑셀은 만들지 않습니다 (CSV·그래프는 저장).

## 템플릿 다시 만들기
양식을 바꾸고 싶으면 기존 SPC 파일에서 시트를 고친 뒤:

```
python3 make_spc_record_template.py 원본SPC.xlsx SPC_기록템플릿.xlsx ../process_capability_report.html
```

마지막 인자로 HTML 을 주면 템플릿이 HTML 안(`SPC_TEMPLATE_B64`)에 들어갑니다. 원본 시트 이름은 스크립트의 `SRC_SHEET` 입니다.

#Include "Mail_filePath.ahk"


; Excel 어플리케이션 객체 생성
excel := ComObject("Excel.Application")

; 엑셀 파일 열기
workbook := excel.Workbooks.Open(filePath)

; 사용할 시트 선택 (첫 번째 시트)
sheet := workbook.Sheets("지출 (2)")

firstRow := 1
lastRow := Sheet.UsedRange.Rows(Sheet.UsedRange.Rows.Count).Row

; firstRow 행부터 처리 시작
row := firstRow

; Excel 파일에서 데이터 마지막 행까지 반복
Loop (lastRow-firstRow)
{

    ; 시작 다음 행에서 데이터 읽기
	row++
	title := sheet.Cells(row, 7).Value
	name := sheet.Cells(row, 1).Value
	Depositor := sheet.Cells(row, 6).Value

	if (Depositor = "한국노동연구원") {

        Continue  ; "한국노동연구원"이면 루프의 다음 반복으로 넘어갑니다.

    } else {

		#Include "Mail_write.ahk"
		
	}
	}

; 엑셀 파일 닫기 및 Excel 어플리케이션 객체 정리
workbook.Close(False)  ; 변경사항을 저장하지 않고 닫기
excel.Quit
excel := ""

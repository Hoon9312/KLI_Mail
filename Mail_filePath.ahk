; Test_Mail_filePath.ahk

; file-selection dialog.ahk

filePath := FileSelect(3, A_Desktop, "Open a file", "Excel Documents (*.xlsx; *.xls)")
if filePath = "" {
    MsgBox "파일 선택 대화상자를 닫습니다."
    Exit
}
else
    MsgBox "파일 경로 :`n" filePath
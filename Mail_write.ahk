; Mail_write.ahk

	; 창 제목으로 특정 크롬 창 활성화
	SetTitleMatchMode 1
	WinActivate("ahk_exe chrome.exe")
	WinActivate("한국노동연구원")
	Sleep(100)

	; "메일쓰기" 버튼 클릭
	CoordMode("Mouse", "client")
	MouseClick("Left", 70, 210)
	Sleep(100)

	; 데이터를 입력할지 물어보기
	click1 := MsgBox("데이터를 입력하시겠습니까?", "주의", "Y/N")
	if (click1 = "No"){

		MsgBox ("작업을 중단합니다")
        break

	}else {

		; 데이터를 순서대로 입력 ( 제목, 이메일 )
		Send(title)
		Sleep(100)


		CoordMode("Mouse", "client")
		MouseClick("Left", 320, 125)  ; '수신' 버튼의 스크린 좌표
		Sleep(100)

		;이메일은 영한 변환문제가 있어 클립보드 사용
		A_Clipboard := name
		Send("^c")
		Send("^v")
		Sleep(100)

		; 사용자에게 다음 행으로 진행할지 물어보기
		click2 := MsgBox("메일을 전송하시겠습니까?", "주의", "Y/N")
		if (click2 = "Yes") { ; Yes를 클릭했을 때

			CoordMode("Mouse", "client")
			MouseClick("Left", 35, 85)  ; '보내기' 버튼의 스크린 좌표
			Sleep(500)

		}else {

			MsgBox ("작업을 중단합니다")
			break
			
		}
	}
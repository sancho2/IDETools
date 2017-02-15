/'
	Dialog Example, by fsw

	compile with:	fbc -s gui dialog.rc dialog.bas

'/

#Include Once "windows.bi"

#Include "AsciiChart.bi"

Declare Sub set_clipboard (Byref x As String)
Declare Sub CreatecmdButtons(ByVal hWin As HWND)
Declare Function DlgProc(ByVal hWin As HWND, ByVal uMsg As UINT, ByVal wParam As WPARAM, ByVal lParam As LPARAM) As Integer

'''
''' Program start
'''

	LoadNoPrintAscii()
	
	''
	'' Create the Dialog
	''
	hInstance=GetModuleHandle(NULL)
	DialogBoxParam(hInstance, Cast(ZString Ptr,dlgTest), NULL, @DlgProc, NULL)
	''
	'' Program has ended
	''

	ExitProcess(0)
	End

'''
''' Program end
'''
Function DlgProc(ByVal hWin As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As Integer
	Dim As Long id, Event, x, y
	Dim hBtn As HWND
	Dim rect As RECT
	Select Case uMsg
		Case WM_INITDIALOG
			'
			CreatecmdButtons(hWin)

		Case WM_CLOSE
			EndDialog(hWin, 0)
			'
		Case WM_LBUTTONUP
			Dim As String s
			id=LoWord(wParam)
			Event=HiWord(wParam)
			's = Str(id)
			If id >= cmdButtons AndAlso id <= cmdButtons + 255 Then 
				MessageBox(NULL, "yes", "yes", MB_OK)
			Else
				MessageBox(NULL, StrPtr(s), "yes", MB_OK)
				
			EndIf 
		Case WM_NOTIFY
			MessageBox(NULL, "notiry", "wwww",MB_OK)
					
		Case WM_COMMAND
			id=LoWord(wParam)
			Event=HiWord(wParam)
			Select Case id
				Case cmdExit
					EndDialog(hWin, 0)
					'
				Case cmdCopy
					If event = BN_CLICKED Then
						Dim s As String
						Dim As ZString * 256 z
						Dim As HWND copy
						
						copy = GetDlgItem(hwin, lblCopy)
						GetWindowText(copy, z, 256)
						s = z
						set_clipboard(s)
					EndIf

				Case cmdButtons To cmdButtons + 255
					If event = BN_CLICKED Then
						Dim s As String
						Dim As ZString * 256 z
						Dim As HWND index, value, copy
						Dim As hwnd txt = GetDlgItem(hwin, id)
						
						index = GetDlgItem(hwin, lblIndex)
						value = GetDlgItem(hwin, lblValue)
						copy = GetDlgItem(hwin, lblCopy)
					
						GetWindowText(txt, z, 256)
						
						s = Trim(Left(z, InStr(z, "-") - 2))
						SetWindowText(index, StrPtr(s))
						
						s = "Chr(" + s + ")"
						SetWindowText(copy, StrPtr(s))

						s = Trim(Mid(z, InStr(z, "-") + 2)) 
						SetWindowText(value, StrPtr(s))

					
					'	MessageBox(NULL, "kkkkkk","QQQQQQ",MB_OK)
					EndIf
			End Select
		Case WM_SIZE
			GetClientRect(hWin,@rect)
			'hBtn=GetDlgItem(hWin,IDC_BTN1)
			x=rect.right-100
			y=rect.bottom-35
			'MoveWindow(hBtn,x,y,97,31,TRUE)
		Case Else
			Return FALSE
			'
	End Select
	Return TRUE

End Function
Sub CreatecmdButtons(ByVal hWin As HWND)
	Dim As HWND tBox
	Dim As hwnd ccc
	Dim As Integer nX, nY, n
	Dim As Integer u = GetConsoleoutputCP()
	Dim As String s = Str(u)

	dim As HFONT hfont1 = CreateFont( _
		12, 0, 0, 0, _
		FW_DONTCARE, _
		FALSE, _
		FALSE, _
		FALSE, _
		DEFAULT_CHARSET, _
		OUT_DEFAULT_PRECIS, _
		CLIP_DEFAULT_PRECIS, _
		DEFAULT_QUALITY, _
		DEFAULT_PITCH, _
		"Terminal")	

	For y As UByte = 1 To 16
		For x As UByte = 1 To 8
			nY = (y - 1) * 18 + 5
			nX = (x - 1) * 62 + 3

			Dim As String s
			s = GetAsciiChar(n)
			s = RightAlignNumber(n) + " - " + s

			tBox = CreateWindowEx(NULL, StrPtr("BUTTON"),NULL, WS_CHILD Or WS_VISIBLE Or SS_SUNKEN Or SS_NOTIFY,_	'ES_READONLY
				 nX, nY, 62, 18, hWin, Cast(HMENU, cmdButtons + n),GetModuleHandle(NULL), NULL)

			SendMessage(tBox, WM_SETFONT, Cast(lparam, hfont1), 0)
			SetWindowText(tbox, StrPtr(s))
			s = GetAsciiChar(n + 128)
			s = Str(n + 128) + " - " + s

			tBox = CreateWindowEx(NULL, StrPtr("Button"),NULL,WS_CHILD Or WS_VISIBLE Or SS_SUNKEN Or SS_NOTIFY,_
				 nX + 510, nY, 62, 18, hWin, Cast(HMENU, cmdButtons + n + 128),GetModuleHandle(NULL), NULL)

			SendMessage(tBox, WM_SETFONT, Cast(lparam, hfont1), 0)
			SetWindowText(tbox, StrPtr(s))
			
			s = GetAsciiChar(n)
			n += 1
		Next
	Next
	hfont1 = CreateFont( _
		18, 0, 0, 0, _
		FW_DONTCARE, _
		FALSE, _
		FALSE, _
		FALSE, _
		DEFAULT_CHARSET, _
		OUT_DEFAULT_PRECIS, _
		CLIP_DEFAULT_PRECIS, _
		DEFAULT_QUALITY, _
		DEFAULT_PITCH, _
		"Terminal")	
	
		Dim As HWND lbl 
		lbl = GetDlgItem(hWin, lblIndex) 
		SendMessage(lbl, WM_SETFONT, Cast(lparam, hfont1), 0)
		lbl = GetDlgItem(hWin, lblValue)
		SendMessage(lbl, WM_SETFONT, Cast(lparam, hfont1), 0)
		lbl = GetDlgItem(hWin, lblCopy)
		SendMessage(lbl, WM_SETFONT, Cast(lparam, hfont1), 0)
	
	
End Sub

Sub set_clipboard (Byref x As String)
  Dim As HANDLE hText = NULL
  Dim As Ubyte Ptr clipmem = NULL
  Dim As Integer n = Len(x)

  If n > 0 Then
    hText = GlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, n + 1)
    Sleep 15
    If (hText) Then
      clipmem = GlobalLock(hText)
      If clipmem Then
        CopyMemory(clipmem, Strptr(x), n)
      Else
        hText = NULL
      End If
      If GlobalUnlock(hText) Then
        hText = NULL
      End If
    End If
    If (hText) Then
      If OpenClipboard(NULL) Then
        Sleep 15
        If EmptyClipboard() Then
          Sleep 15
          If SetClipboardData(CF_TEXT, hText) Then
            Sleep 15
          End If
        End If
        CloseClipboard()
      End If
    End If
  End If
End Sub

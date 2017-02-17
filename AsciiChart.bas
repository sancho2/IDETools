/'
	Dialog Example, by fsw

	compile with:	fbc -s gui dialog.rc dialog.bas

'/

#Include Once "windows.bi"

#Include "AsciiChart.bi"

Declare Sub ReLoadAscii(ByVal hWin As HWND, start As ubyte)
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
				Case cmdPage
					If event = BN_CLICKED Then
						aStart = IIf(aStart = 0, 128, 0)
						ReLoadAscii(hWin, aStart) 
					EndIf
						
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
	Dim As HWND hBtn, hTmp
	Dim As HDC dc
	Dim As SIZE sz
	Dim As Integer nX, nY, n, txtLen, txtHght, dlgWidth, dlgHeight, btn, btnHght, clientHeight, btnRow
	Dim As String s 
	Dim As RECT r, dlgRect 
	Const As UByte MARGIN = 5, COPY_OFFSET = 123
	Dim As ZString * 256 z
	
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

	sendmessage(hWin, WM_SETFONT, Cast(LPARAM, hfont1), 0)

	dc = GetDC(hWin)
	SetMapMode(dc,MM_TEXT)
	
	' get an estimate of the size of button face
	GetTextExtentPoint32(dc, "888 - WWW", 9, @sz)
	
	ReleaseDC(hWin, dc)

	' the dialog
	GetClientRect(hWin, @dlgRect)
	
	txtLen = sz.cx
	txtHght = sz.cy + 2
	dlgWidth = txtLen * 8 + 24
	dlgHeight = dlgRect.bottom - dlgRect.top 			' we need to add room for bottom button row
	clientHeight = dlgHeight  

	' the copy button  
	hBtn = GetDlgItem(hWin, cmdCopy)
	GetWindowRect(hBtn, @r)

	' we now have the height of the button
	btnHght = r.bottom - r.top
	dlgHeight += btnHght + 16
	
	' resize the dialog
	SetWindowPos(hWin, HWND_TOP, 0, 0, dlgWidth, dlgHeight, SWP_NOMOVE Or SWP_NOZORDER)  
	
	' align the page button to the left edge of the button grid
	hTmp = hBtn 		' hold on to the copy button handle 
	hBtn = GetDlgItem(hWin, cmdPage)
	GetWindowRect(hBtn, @r)
	
	' move the page button
	nX = MARGIN
	btnRow = clientHeight - (btnHght + 8)
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)
	
	' move the index
	nX += r.right - r.left + 5
	hBtn = GetDlgItem(hWin, lblIndex)
	GetWindowRect(hBtn, @r)
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)
	
	' move the value
	nX += r.right - r.left + 5
	hBtn = GetDlgItem(hWin, lblValue)
	GetWindowRect(hBtn, @r)
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)
	  
	' move the copy label
	nX += r.right - r.left + 5
	hBtn = GetDlgItem(hWin, lblCopy)
	GetWindowRect(hBtn, @r)
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)

	' move the copy button
	nX += r.right - r.left + 5
	GetWindowRect(hTmp, @r)
	MoveWindow(hTmp, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)
	
	' move the exit button
	hBtn = GetDlgItem(hWin, cmdExit)
	GetWindowRect(hBtn, @r)
	nx = (8 * txtLen + MARGIN) - (r.right - r.left)   
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)
	
	For x As UByte = 1 To 8
		For y As UByte = 1 To 16
			nY = (y - 1) * txtHght + MARGIN
			nX = (x - 1) * txtLen + MARGIN

			Dim As String s
			s = GetAsciiChar(n)
			s = RightAlignNumber(n) + " - " + s

			hBtn = CreateWindowEx(NULL, StrPtr("BUTTON"),NULL, WS_CHILD Or WS_VISIBLE Or SS_SUNKEN Or SS_NOTIFY,_	'ES_READONLY
				 nX, nY, txtLen, txtHght, hWin, Cast(HMENU, cmdButtons + n),GetModuleHandle(NULL), NULL)

			SendMessage(hBtn, WM_SETFONT, Cast(lparam, hfont1), 0)
			SetWindowText(hBtn, StrPtr(s))

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
Sub ReLoadAscii(ByVal hWin As HWND, start As ubyte)
	'
	Dim As HWND btnHandle
	Dim As UByte n
	Dim As String s	
	
	For x As UByte = 1 To 8
		For y As UByte = 1 To 16
			btnHandle = GetDlgItem(hWin, cmdButtons + n) 

			s = GetAsciiChar(n + start)
			s = Str(n + start) + " - " + s

			SetWindowText(btnHandle, StrPtr(s))
			n += 1
		Next
	Next
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

#include once "windows.bi"
#Include Once "win/uxtheme.bi"
#Include Once "win/richedit.bi"

#Include "..\..\Inc\Addins.bi"

#include "ASCIIChartA.bi"
'------------------------------------------------------------------------------------------------------------------
Declare Function DlgProc(ByVal hWin As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As Integer
Declare function DllFunction CDECL alias "DllFunction" (byval hWin as HWND,byval uMsg as UINT,byval wParam as WPARAM,byval lParam as LPARAM) as bool 
Declare sub AddToMenu(byval id as integer)
Declare Sub set_clipboard (Byref x As String)
Declare Sub ReLoadAscii(ByVal hWin As HWND, start As ubyte)
Declare Sub CreatecmdButtons(ByVal hWin As HWND)
'------------------------------------------------------------------------------------------------------------------
Sub AddToMenu(byval id as integer)
	dim hMnu as HMENU
	Dim buff As ZString*256

	' Get handle to 'Tools' popup
	hMnu=GetSubMenu(lpHANDLES->hmenu,7)

	buff="AsciiXChart"
	AppendMenu(hMnu,MF_STRING,id,@buff)

end sub

' Returns info on what messages the addin hooks into (in an ADDINHOOKS type).
Function InstallDll CDECL alias "InstallDll" (byval hWin as HWND,byval hInst as HINSTANCE) as ADDINHOOKS ptr EXPORT

	' The dll's instance
	hInstance=hInst
	' Get pointer to ADDINHANDLES
	lpHandles=Cast(ADDINHANDLES ptr,SendMessage(hWin,AIM_GETHANDLES,0,0))
	' Get pointer to ADDINDATA
	lpData=Cast(ADDINDATA ptr,SendMessage(hWin,AIM_GETDATA,0,0))
	' Get pointer to ADDINFUNCTIONS
	lpFunctions=Cast(ADDINFUNCTIONS ptr,SendMessage(hWin,AIM_GETFUNCTIONS,0,0))

	' Get a menu ID
	IDM_ASCIICHART=SendMessage(hWin,AIM_GETMENUID,0,0)
	AddToMenu(IDM_ASCIICHART)

	' Messages this addin will hook into
	hooks.hook1=HOOK_COMMAND
	hooks.hook2=0
	hooks.hook3=0
	hooks.hook4=0
	return @hooks

end function

' FbEdit calls this function for every addin message that this addin is hooked into.
' Returning TRUE will prevent FbEdit and other addins from processing the message.
function DllFunction CDECL alias "DllFunction" (byval hWin as HWND,byval uMsg as UINT,byval wParam as WPARAM,byval lParam as LPARAM) as bool EXPORT

	select case uMsg
		case AIM_COMMAND
			'
			if loword(wParam)=IDM_ASCIICHART then
				' Our menu item has been selected. Show the dialog
				DialogBoxParam(hInstance, Cast(ZString Ptr,dlgTest), NULL, @DlgProc, NULL)
				'DialogBoxParam(hInstance,Cast(zstring ptr,IDD_DLGSNIPLET),hWin,@SnipletProc,NULL)
			endif
			
		case AIM_CLOSE
			'
	end select
	return FALSE

end Function
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
					
		Case WM_COMMAND
			id=LoWord(wParam)
			Event=HiWord(wParam)
			Select Case id
				Case cmdInsert
					Dim As LPDWORD a, b 
					Dim As ZString * 9 z
					Dim As CHARRANGE cr
					Dim As HWND label
					
					label = GetDlgItem(hWin, lblCopy)
					SendMessage(label, WM_GETTEXT, 9, Cast(LPARAM, @z))
					
					SendMessage(lpHandles->hred,EM_EXGETSEL,0, Cast(LPARAM,@cr))
					'z = "Chr(000)"
					sendmessage(lpHandles->hred, EM_REPLACESEL, TRUE, Cast(LPARAM, @z))
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

	' move the insert button
	nX += r.right - r.left + 5
	hBtn = GetDlgItem(hWin, cmdInsert)
	GetWindowRect(hBtn, @r)
	MoveWindow(hBtn, nX, btnRow, r.right - r.left, r.bottom - r.top,  1)

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

			hBtn = CreateWindowEx(NULL, StrPtr("BUTTON"),NULL, WS_CHILD Or WS_VISIBLE Or bS_FLAT,_	'ES_READONLY  Or SS_SUNKEN Or SS_NOTIFY
				 nX, nY, txtLen, txtHght, hWin, Cast(HMENU, cmdButtons + n),GetModuleHandle(NULL), NULL)
			SetWindowTheme(hBtn, " ", " ")

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

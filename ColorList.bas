'---------------------------------------------------------------------------------------------------------------
' ColorList - IDE Color selecting tool
'
' Sancho2 February 21, 2017 version 1.5.1
' 1.5.1: Fixed exe freeze when copy is pressed with no items in the out list 
'---------------------------------------------------------------------------------------------------------------
Dim Shared As BOOLEAN OkToEvent = TRUE
#Define GET_X_LPARAM(lp) clng(cshort(LOWORD(lp)))
#define GET_Y_LPARAM(lp) clng(cshort(HIWORD(lp)))
'#define GET_WHEEL_DELTA_WPARAM(wParam) cshort(HIWORD(wParam))

#include once "windows.bi"
#Include Once "win/commctrl.bi"
#Include Once "win/commdlg.bi"
#Include Once "win/shellapi.bi"
#Include "colordata.bi"
#Include "ColorsA.bi"
'---------------------------------------------------------------------------------------------------------------
Declare Sub InitScroll(ByVal hWin As HWND)
Declare Sub CreateSwatches()
Declare Sub set_clipboard (Byref x As String)
Declare Function GetVisibleItemCount(ByVal hWin As HWND) As Integer
Declare Sub AddColorsToLists()
Declare Sub ListBox_ScrollDown(ByVal hWin As HWND, ByVal amt As UByte)
Declare Sub ListBox_ScrollUp(ByVal hWin As HWND, ByVal amt As UByte)
Declare Sub ListBox_ScrollTo(ByVal hWin As HWND, ByVal amt As UByte)
Declare Sub DrawSwatches(ByVal hWin As HWND)
Declare Sub InvalidateSwatchRect()
Declare Sub UnSelectItems(ByVal hLst As HWND)
Declare Sub ChangeSelection(ByVal lstBox As HWND)
Declare Sub AddItems()
Declare Sub DeleteItems()
Declare Sub CreateClipboardEntry()
Declare Function IsPointOnScrollBar(ByVal x As Integer, ByVal y As Integer) As BOOLEAN
Declare Function GetListBoxVisibleItems(ByVal hLst As HWND) As Integer
Declare Function GetMaxTopIndex() As Integer
'---------------------------------------------------------------------------------------------------------------
Function IsDuplicateItem(byref item as string) as boolean
	'
	Dim As HWND hOut
	Dim As Integer index		
	hOut = GetDlgItem(hWnd, lstOut)
	
	index = SendMessage(hOut, LB_FINDSTRINGEXACT, -1, Cast(WPARAM, StrPtr(item)))
	If index = LB_ERR Then
		Return FALSE 
	EndIf

	Return TRUE 
 	
End Function

Function GetListBoxVisibleItems(ByVal hLst As HWND) As Integer
	'
	Dim As Integer rowH
	Dim As RECT r 
	
	rowH = SendMessage(hLst, LB_GETITEMHEIGHT, 0, 0)
	GetClientRect(hLst, @r)
	Return  (r.bottom - r.top) \ rowH  
	
End Function
Function GetMaxTopIndex() As Integer
	'
	Dim As HWND hNames
	Dim As Integer count, n
	Dim As String s
	
	hNames = GetDlgItem(hWnd, lstNames)
	count = SendMessage(hNames, LB_GETCOUNT, 0, 0)
	n = GetListBoxVisibleItems(hNames)
	
	Return count - n + 1
	
End Function

Sub CreateClipboardEntry()
	'
	Dim As HWND hOut, hPrefix, hConstant
	Dim As String txt, s
	Dim As ZString * 256 constant, prefix
	Dim As Integer l, count, buffer(Any), index
	
	hOut = GetDlgItem(hWnd, lstOut)
	hConstant = GetDlgItem(hWnd, txtConstant)
	hPrefix = GetDlgItem(hWnd, txtPrefix)
	
	l = GetWindowTextLength(hConstant)
	GetWindowText(hConstant, @constant, l)	

	l = GetWindowTextLength(hPrefix)
	GetWindowText(hPrefix, @prefix, l + 1)	

	hOut = GetDlgItem(hWnd, lstOut)
	count = SendMessage(hOut, LB_GETCOUNT, 0, 0)
	If count < 1 Then
		Exit Sub
	EndIf
	txt = ""
	s = constant + " " + prefix
	For x As UByte = 0 To count - 1
		txt += s
		index = SendMessage(hOut, LB_GETITEMDATA, x, 0)
		'SendMessage(hOut, LB_GETTEXT,  count,  Cast(LPARAM, @buffer(1)))
		s = colors(index).Name
		txt += s
		txt += " = "
		s = colors(index).ToRGBHex
		txt = txt + "&H" + s  
		s = ", _" + CRLF + Space(14) + prefix
	Next

	set_clipboard(txt)

End Sub

Sub DeleteItems()
	'
	Dim As HWND hOut
	Dim As String s
	Dim As Integer buffer(Any), count 

	hOut = GetDlgItem(hWnd, lstOut)
	count = SendMessage(hOut, LB_GETSELCOUNT, 0, 0)

	ReDim buffer(1 To count)
	SendMessage(hOut, LB_GETSELITEMS, count, Cast(LPARAM, @buffer(1)))

	For x As UByte = count To 1 Step - 1
		SendMessage(hOut, LB_DELETESTRING, buffer(x), 0)
	Next
	
End Sub
Sub AddItems()
	'
	Dim As HWND hNames, hOut
	Dim As zString * 256 z  
	Dim As Integer buffer(Any), count, n, index

	hNames = GetDlgItem(hWnd, lstNames)
	hOut = GetDlgItem(hWnd, lstOut)
	
	UnSelectItems(hOut)
	
	count = SendMessage(hNames, LB_GETSELCOUNT, 0, 0)
	ReDim buffer(1 To count)
	SendMessage(hNames, LB_GETSELITEMS, count, Cast(LPARAM, @buffer(1)))
	
	For x As UByte = 1 To count
		SendMessage(hNames, LB_GETTEXT, buffer(x), Cast(LPARAM, @z))
		If IsDuplicateItem(z) = FALSE  then
			index = SendMessage(hNames, LB_GETITEMDATA, buffer(x), 0)
			SendMessage(hOut, LB_ADDSTRING, 0, Cast(LPARAM, @z))
			n = SendMessage(hOut, LB_GETCOUNT, 0, 0)
			SendMessage(hOut, LB_SETSEL, TRUE, n - 1)
			SendMessage(hOut, LB_SETITEMDATA, n - 1,  index)
		EndIf
	Next
	'UnSelectItems(hNames)
End Sub

Sub InvalidateSwatchRect()
	'
	Dim As HWND hColor
	Dim As RECT r  
	
	hColor = GetDlgItem(hWnd, lstNames)
	GetWindowRect(hColor, @r)

	r.right = r.left
	r.left = 0
	r.top = 0

	InvalidateRect(hwnd, @r, TRUE)
	
End Sub
Function UpdateSwatch(hStatic As HWND, dc As HDC, ByVal hB As HBRUSH) As HBRUSH
	'
	Dim As HWND hColor
  	Dim As Integer id, index, topIndex
  	Dim As String s
  
	hColor = GetDlgItem(hWnd, lstNames)
	topIndex = SendMessage(hColor, LB_GETTOPINDEX, 0, 0)

  	id = GetDlgCtrlID(hStatic)
	index = id -swatches + topindex 

	hB = CreateSolidBrush(colors(index).ToBGRValue)  
  	SetBkColor(dc, colors(index).ToBGRValue)
	
	Return hB
End Function
Sub CreateSwatches()
	'
	Dim As HWND hStatic, hColor
	Dim As Point p
	Dim As Integer rowH, topIndex, nX, nY, w 
	Dim As RECT r
	
	hColor = GetDlgItem(hWnd, lstNames)

	ClientToScreen(hWnd, @p)

	rowH = SendMessage(hColor, LB_GETITEMHEIGHT, 0, 0)
	topIndex = SendMessage(hColor, LB_GETTOPINDEX, 0, 0)
	GetWindowRect(hColor, @r)
	For x As UByte = 1 To 26
		ny = (r.top - p.y) + ((x - 1) * rowH) 
		w = r.left '+ p.x
		hStatic = CreateWindowEx(NULL, StrPtr("Static"),NULL, WS_CHILD Or WS_VISIBLE Or SS_SIMPLE Or ss_notify,_	'ES_READONLY  Or SS_SUNKEN Or SS_NOTIFY
		 nX, nY, w, rowH, hWnd, Cast(HMENU, swatches + x ),GetModuleHandle(NULL), NULL)
		 SetWindowText(hStatic, StrPtr("            "))
	Next

End Sub

Sub InitScroll(ByVal hWin As HWND)
	'
	Dim As HWND hScroll
	Dim As SCROLLINFO sInfo
	Dim As Integer count = UBound(colors), visibleItemCount
	
	visibleItemCount = GetVisibleItemCount(hWin)	
	With sInfo
		.cbSize = SizeOf(SCROLLINFO)
		.fMask = SIF_RANGE
		.nMin = 1
		.nMax = count - visibleItemCount
	End With
	
	hScroll = GetDlgItem(hWin,scbVert)
	
	SetScrollInfo(hScroll, SB_CTL, @sInfo, TRUE) 
	 
		
End Sub
Sub UnSelectItems(ByVal hLst As HWND)
	'
	Dim As Integer buffer(Any), count
	count = SendMessage(hLst, LB_GETSELCOUNT, 0, 0)
	ReDim buffer(1 To count)
	
	SendMessage(hLst, LB_GETSELITEMS, count, Cast(LPARAM, @buffer(1)))
	
	For x As UByte = 1 To count
		SendMessage(hLst, LB_SETSEL, FALSE, buffer(x))
	Next
	
End Sub
Sub ChangeSelection(ByVal lstBox As HWND)
	'
	Dim As Integer index 
	Dim As String s
	Dim As HWND hTemp
	Dim As Integer buffer(Any), count
	
	If GetDlgCtrlID(lstBox) = lstNames Then
		hTemp = GetDlgItem(hWnd, lstRGB) 
	Else
		hTemp = GetDlgItem(hWnd, lstNames)
	EndIf
	
	UnSelectItems(hTemp)
	
	count = SendMessage(lstBox, LB_GETSELCOUNT, 0, 0)
	ReDim buffer(1 To count)
	
	SendMessage(lstBox, LB_GETSELITEMS, count, Cast(LPARAM, @buffer(1)))
	
	For x As UByte = 1 To count
		SendMessage(hTemp, LB_SETSEL, TRUE, buffer(x))
	Next
	
End Sub
Function IsPointOnScrollBar(ByVal x As Integer, ByVal y As Integer) As BOOLEAN
	'
	Dim As HWND hScroll
	Dim As RECT r
	Dim As Point p
	
	p = Type<Point>(x, y)
	
	hScroll = GetDlgItem(hWnd, scbVert)
	GetWindowRect(hScroll, @r)
	
	r.left = 0		' include everyting on the left
	
	Return cbool(PtInRect(@r, p))	

End Function
Sub WHereSwatch(ByVal index As Integer)
	'
	Dim As String s
	Dim As HWND hSwatch
	Dim As RECT r
	s =Str(index - 1 + swatches)
	MessageBox(NULL, StrPtr(s), "hello",MB_OK)
	'End
	hSwatch = GetDlgItem(hWnd, index + swatches)
	s =Str(hSwatch)
	MessageBox(NULL, s, "hello",MB_OK)
	getwindowrect(hSwatch, @r)
	
	s = Str(r.left) + " " + Str(r.top) + " " + Str(r.right) + " " + Str(r.bottom)
	
	MessageBox(NULL, s, "hello",MB_OK) 
	
End Sub

Function WndProc(ByVal hWin As HWND,ByVal uMsg As UINT,ByVal wParam As WPARAM,ByVal lParam As LPARAM) As Integer
	'
	Dim As HBRUSH hB 
	Dim As Integer id
	'Static test As Integer = 0
	Select Case uMsg
		

		Case WM_CTLCOLORSTATIC 
			id = GetDlgCtrlID(Cast(HWND, lParam))
			If id >1200 AndAlso id <= 1226 Then 
				If hB <> NULL Then
					DeleteObject(hB)
				EndIf
				hB = UpdateSwatch(Cast(HWND, lParam), Cast(HDC, wParam), hB)
				Return Cast(INT_PTR, hB)
			EndIf
			
		
		Case WM_INITDIALOG

			hWnd=hWin
			Dim As HWND lstBox
			Dim As LVCOLUMN column
			lstBox = GetDlgItem(hWin, lstNames)
			AddColorsToLists()
			InitScroll(hWin)
			CreateSwatches()

		Case WM_MOUSEWHEEL
			Dim As Integer xPos, yPos, amt
			Dim As String s
			
			xPos = GET_X_LPARAM(lParam)
			yPos = GET_Y_LPARAM(lParam)
			's = Str(xPos)
			If IsPointOnScrollBar(xPos, yPos) = TRUE Then
				amt = GET_WHEEL_DELTA_WPARAM(wparam)
				amt = amt/120
				's = Str(amt)
				'MessageBox(NULL, s, "X", MB_OK)
				If amt < 0 Then
					ListBox_ScrollDown(hWnd, -1 * amt)
				Else
					ListBox_ScrollUp(hWnd, amt)
				EndIf
				InvalidateSwatchRect()
			EndIf 
			
		
		Case WM_VSCROLL
			Dim As Long hWord, lWord
			Dim As String s
			Dim As Integer i
			Dim As RECT r = Type<RECT>(0,0, 200,600)
			hWord = HiWord(wParam)
			lWord = LoWord(wParam)
			
			Select Case lWord
				Case SB_THUMBPOSITION, SB_THUMBTRACK
					ListBox_ScrollTo(hWin, hWord)
				Case SB_LINEDOWN
					ListBox_ScrollDown(hWin, 1)
				Case SB_LINEUP
					ListBox_ScrollUp(hWin, 1)
				Case SB_PAGEDOWN
					i = GetVisibleItemCount(hWin)
					ListBox_ScrollDown(hWin, i)
				Case SB_PAGEUP
					i = GetVisibleItemCount(hWin)
					ListBox_ScrollUp(hWin, i) 
			End Select 
			InvalidateSwatchRect() 
			
		Case WM_COMMAND
					 
			Select Case HiWord(wParam)
				Case LBN_DBLCLK
					AddItems()
				Case LBN_SELCHANGE
					If OkToEvent Then
						OkToEvent = FALSE
						ChangeSelection(Cast(HWND, lParam))
						OkToEvent = TRUE 
					EndIf

				Case BN_CLICKED,1
					Select Case LoWord(wParam)
						Case cmdCopy
							WHereSwatch(1)
							CreateClipboardEntry()
						Case cmdRemove
							DeleteItems()
						Case cmdAdd
							AddItems()
						Case cmdExit		' IDM_FILE_EXIT,
							SendMessage(hWin,WM_CLOSE,0,0)
					End Select
					'
			End Select
			'
		Case WM_SIZE
			'
		Case WM_CLOSE
			DeleteObject(hB)
			DestroyWindow(hWin)
			'
		Case WM_DESTROY
			PostQuitMessage(NULL)
			'
		Case Else
			Return DefWindowProc(hWin,uMsg,wParam,lParam)
			'
	End Select
	Return 0

End Function

Function WinMain(ByVal hInst As HINSTANCE,ByVal hPrevInst As HINSTANCE,ByVal CmdLine As ZString ptr,ByVal CmdShow As Integer) As Integer
	Dim wc As WNDCLASSEX
	Dim msg As MSG

	' Setup and register class for dialog
	wc.cbSize=SizeOf(WNDCLASSEX)
	wc.style=CS_HREDRAW or CS_VREDRAW
	wc.lpfnWndProc=@WndProc
	wc.cbClsExtra=0
	wc.cbWndExtra=DLGWINDOWEXTRA
	wc.hInstance=hInst
	wc.hbrBackground=Cast(HBRUSH,COLOR_BTNFACE+1)
	wc.lpszMenuName=Cast(ZString Ptr,IDM_MENU)
	wc.lpszClassName=@ClassName
	wc.hIcon=LoadIcon(NULL,IDI_APPLICATION)
	wc.hIconSm=wc.hIcon
	wc.hCursor=LoadCursor(NULL,IDC_ARROW)
	RegisterClassEx(@wc)
	' Create and show the dialog
	CreateDialogParam(hInstance,Cast(ZString Ptr,IDD_DIALOG),NULL,@WndProc,NULL)
	ShowWindow(hWnd,SW_SHOWNORMAL)
	UpdateWindow(hWnd)
	' Message loop
	Do While GetMessage(@msg,NULL,0,0)
		TranslateMessage(@msg)
		DispatchMessage(@msg)
	Loop
	Return msg.wParam

End Function
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
Sub AddColorsToLists()
	'
	Dim As HWND hNames, hRGB
	Dim As String s
	'Dim As Integer n

	hNames = GetDlgItem(hWnd, lstNames)
	hRGB = GetDlgItem(hWnd, lstRGB)
	
	For x As UByte = 1 To UBound(colors)
		s = colors(x).Name 
		SendMessage(hNames, LB_ADDSTRING, 0, Cast(LPARAM, StrPtr(s)))
		SendMessage(hNames, LB_SETITEMDATA, x - 1, x)
		s = colors(x).ToRGBString
		SendMessage(hRGB, LB_ADDSTRING, 0, Cast(LPARAM, StrPtr(s)))
	Next
	
End Sub

Function GetVisibleItemCount(ByVal hWin As HWND) As Integer
	'
	Dim As HWND hNames
	Dim As Integer h
	Dim As RECT r

	hNames = GetDlgItem(hWin, lstNames)
	
	GetClientRect(hNames, @r)
	h = sendmessage(hNames, LB_GETITEMHEIGHT, 0, NULL)
	Return (r.bottom - r.top) \ h
End Function
Sub ListBox_ScrollTo(ByVal hWin As HWND, ByVal amt As UByte)
	'
	Dim As HWND hNames, hRGB
	Dim As Integer topIndex 
	
	hNames = GetDlgItem(hWin, lstNames)
	hRGB = GetDlgItem(hWin, lstRGB) 

	SendMessage(hNames, LB_SETTOPINDEX, amt, 0)
	SendMessage(hRGB, LB_SETTOPINDEX, amt, 0)

	Dim As SCROLLINFO sInfo
	Dim As HWND hScroll	
	With sInfo
		.cbSize = SizeOf(SCROLLINFO)
		.fMask = SIF_POS
		.nPos = amt
	End With
	
	hScroll = GetDlgItem(hWin,scbVert)
	
	SetScrollInfo(hScroll, SB_CTL, @sInfo, TRUE)		
	

End Sub
Sub ListBox_ScrollUp(ByVal hWin As HWND, ByVal amt As UByte)
	'
	Dim As HWND hNames, hRGB
	Dim As Integer topIndex 
	
	hRGB = GetDlgItem(hWin, lstRGB) 
	hNames = GetDlgItem(hWin, lstNames)
	topIndex = sendmessage(hNames, LB_GETTOPINDEX, 0, 0)
	topIndex -= amt
	If topIndex < 0 Then
		topIndex = 0
	EndIf

	SendMessage(hNames, LB_SETTOPINDEX, topIndex, 0)
	SendMessage(hRGB, LB_SETTOPINDEX, topIndex, 0)

	Dim As SCROLLINFO sInfo
	Dim As HWND hScroll	
	With sInfo
		.cbSize = SizeOf(SCROLLINFO)
		.fMask = SIF_POS
		.nPos = topIndex
	End With
	
	hScroll = GetDlgItem(hWin,scbVert)
	
	SetScrollInfo(hScroll, SB_CTL, @sInfo, TRUE)		
	
End Sub
Sub ListBox_ScrollDown(ByVal hWin As HWND, ByVal amt As UByte)
	'
	Dim As HWND hNames, hRGB
	Dim As UByte topIndex, maxTopIndex
	
	hRGB = GetDlgItem(hWin, lstRGB) 
	hNames = GetDlgItem(hWin, lstNames)
	topIndex = sendmessage(hNames, LB_GETTOPINDEX, 0, 0)
	topIndex += amt
	maxTopIndex = GetMaxTopIndex()
	If topIndex > maxTopIndex Then
		topIndex = maxTopIndex
	EndIf

	SendMessage(hNames, LB_SETTOPINDEX, topIndex, 0)
	SendMessage(hRGB, LB_SETTOPINDEX, topIndex, 0)

	Dim As SCROLLINFO sInfo
	Dim As HWND hScroll	

	With sInfo
		.cbSize = SizeOf(SCROLLINFO)
		.fMask = SIF_POS
		.nPos = topIndex
	End With
	
	hScroll = GetDlgItem(hWin,scbVert)
	
	SetScrollInfo(hScroll, SB_CTL, @sInfo, TRUE)		
End Sub
'=================================================================================================================
' Program start
'=================================================================================================================
LoadColors()

hInstance=GetModuleHandle(NULL)

InitCommonControls

WinMain(hInstance,NULL,CommandLine,SW_SHOWDEFAULT)

ExitProcess(0)


'===================================================================================================
' MethodStubs
'===================================================================================================
#Include "windows.bi"

Declare Sub CreateDeclare(ByRef clipText As String, ByVal o as UInteger)
Declare Function get_clipboard () As String
Declare Sub set_clipboard (Byref x As String)
Declare Sub CreateProcedure(ByRef txt As String)
Declare Sub HandleClipboardText()
Declare Sub FixProcedureName(ByRef txt as String, byref o as UInteger )

Function get_clipboard () As String
  Dim As Zstring Ptr s_ptr
  Dim As HANDLE hglb
  Dim As String s = ""

  If (IsClipboardFormatAvailable(CF_TEXT) = 0) Then Return ""

  If OpenClipboard( NULL ) <> 0 Then
    hglb = GetClipboardData(cf_text)
    s_ptr = GlobalLock(hglb)
    If (s_ptr <> NULL) Then
      s = *s_ptr
      GlobalUnlock(hglb)
    End If
    CloseClipboard()
  End If

  Return s
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
'Screenres 800, 600, 32
Sub CreateDeclare(ByRef clipText As String, ByVal o as UInteger)
	'
	Dim As String s
	Dim As Integer x, y, n
	
	FixProcedureName(clipText, o)
   s = "Declare " + clipText
	set_clipboard(s)
	
End Sub
Sub CreateProcedure(txt As String)
	'
	Dim As String s 
	
	s = Trim(Mid(txt, 9))
	set_clipboard(s)
End Sub
Sub FixProcedureName(ByRef txt as String, byref o as UInteger )
   '
   Dim as uinteger n = instr(txt, ".")
   
   if n = 0 then 
      return
   endif 
      
   txt = mid(txt, n + 1)

End Sub

Sub HandleClipboardText()
	'
	Dim txt As String
	
	txt = get_clipboard()
	txt = Trim(txt, Any " " + Chr(9))
	If LCase(Left(txt, 4)) = "sub " Then 
		' create a declare statement for the clipboard
		CreateDeclare(txt, 3)
	EndIf
	If LCase(Left(txt, 9)) = "function " Then
		' create a declare statement for the clipboard
		CreateDeclare(txt, 9)
	EndIf
   If lcase(left(txt, 8)) = "property" Then
		' create a declare statement for the clipboard
		CreateDeclare(txt, 8)
   Endif
	If LCase(Left(txt, 8)) = "declare " Then
		' create proc stub from declare statement
		CreateProcedure(txt)
	EndIf
End Sub
HandleClipboardText()
'Screenres 800, 600, 32
'Declare Sub HandleClipboardText()
'sleep

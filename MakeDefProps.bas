'===================================================================================================
' MethodStubs
'===================================================================================================
#Include "windows.bi"
'===================================================================================================
Const As String CRLF = Chr(13) + Chr(10)
Const as String moo1 = "Declare Property test(Byval n As String)"
Const As String moo2 = "Declare Property test(Byval index as Ubyte) as string"
'							  123456789 123456789 123456789 1234"		
Const As String moo3 = "Declare Property test() as string "
'===================================================================================================
Declare Function CreatePropertyGet(ByRef clipText As String, ByVal o as UInteger) As String 
Declare Function CreatePropertySet(ByRef clipText As String, ByVal o as UInteger) As String
Declare Function GetProcedureName(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
Declare Function GetProcedureParams(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
Declare Function GetParameterVariable(ByRef txt As String, _
									ByRef x1 As UInteger, ByRef x2 As UInteger) As String

'===================================================================================================
Declare Function get_clipboard () As String
Declare Sub set_clipboard (Byref x As String)
Declare Sub HandleClipboardText()
'===================================================================================================
Function  CreatePropertyGet(Byref txt As String , Byval o As Uinteger ) As String 
	'
	Dim As UInteger x1, x2
	Dim As String pn, pp, outS, pv
	 

	pn = GetProcedureName(txt, o, x1, x2)
	pp = GetProcedureParams(txt, o, x1, x2)
	pv = GetParameterVariable(pp, x1, x2)


	
	outS = "Property ~." 
	outS += pn
	outS += "(" + pp + ")"
	outS += CRLF
	outS += "'"
	outS += CRLF
	outS += Chr(9) + "Return This._" + pn
	If pv <> "" Then 
		outS += "(" + pv + ")" 
	EndIf
	outS += CRLF
	outS += "End Property"
	outS += CRLF

	'? txt
	'? ">";pn;"<"
	'? ">";pp;"<"
	'? ">";pv;"<"
	'? outS
	'Sleep
	'end
	Return outS	

End Function 

Function GetParameterVariable(ByRef txt As String, _
									ByRef x1 As UInteger, ByRef x2 As UInteger) As String
	'
	Dim As Byte n = InStrRev(LCase(txt), "as") - 1  
	
	If n < 1 Then
		Return ""
	EndIf

	While Mid(txt, n, 1) = " " AndAlso n > 1
		n -= 1
	Wend
	
	If n = 0 Then 
		Return ""
	EndIf
	'Cls
	'? txt
	'? n
	'Sleep
	'end
	x2 = n
	n -= 1
	While Mid(txt, n, 1) <> " " AndAlso n > 1
		n -= 1
	Wend
	
	If n = 0 Then
		Return ""
	EndIf
	x1 = n + 1
	Return Mid(txt, x1, x2 - x1 + 1)

End Function

Function CreatePropertySet(Byref txt As String , Byval o As Uinteger ) As String 
	'
	Dim As UInteger x1, x2
	Dim As String pn, pp, outS, pv
	 

	pn = GetProcedureName(txt, o, x1, x2)
	pp = GetProcedureParams(txt, o, x1, x2)
	pv = GetParameterVariable(pp, x1, x2)
	
	outS = "Property ~." 
	outS += pn
	outS += "(" + pp + ")"
	outS += CRLF
	outS += "'"
	outS += CRLF
	outS += Chr(9) + "This._" + pn
	outS += " " + pv
	outS += CRLF
	outS += "End Property"
	outS += CRLF

	Return outS	

End Function 

Function GetProcedureParams(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
	'
	x1 = InStr(txt, "(") + 1 
	x2 = Instr(txt, ")") - 1
	
	Return Mid(txt, x1, x2 - x1 + 1)
		
End Function

Function GetProcedureName(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
	'			
	Dim As Ubyte n = Instr(txt, "(")
	
	n -= 1
	While Mid(txt, n, 1) = " "
		n -= 1	
	Wend
	'Print txt
	'Print o, n
	x1 = o
	x2 = n
	Return Mid(txt, o, n - o + 1) 

end function


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

Sub HandleClipboardText()
	'
	Dim As String txt
	Dim As UByte n1, n2, n3
	txt = get_clipboard()
	'txt = moo3
	'txt = moo2
	'txt = moo1
	
	txt = Trim(txt, Any " " + Chr(9))

	If  Left(LCase(txt), 16) <> "declare property" Then
		Exit Sub
	End If
	
	n3 = InStrRev(txt, ")")
	
	n2 = InStrRev(txt, " ")
	
	n1 = InStrRev(LCase(txt), "as", n2)

	If n1 = 0 OrElse n2 = 0 OrElse n3 > n2 Then
		set_clipboard(CreatePropertySet(txt, 18))
	Else
		set_clipboard(CreatePropertyGet(txt, 18))		
	EndIf
	
	'Sleep
	
	

'	If LCase(Left(txt, 4)) = "sub " Then 
'		' create a declare statement for the clipboard
'		CreateDeclare(txt, 3)
'	EndIf
'	If LCase(Left(txt, 9)) = "function " Then
'		' create a declare statement for the clipboard
'		CreateDeclare(txt, 9)
'	EndIf
'   If lcase(left(txt, 8)) = "property" Then
'		' create a declare statement for the clipboard
'		CreateDeclare(txt, 8)
'   Endif
'	If LCase(Left(txt, 8)) = "declare " Then
'		' create proc stub from declare statement
'		CreateProcedure(txt)
'	EndIf
End Sub
HandleClipboardText()
'Screenres 800, 600, 32
'Declare Sub HandleClipboardText()
'sleep

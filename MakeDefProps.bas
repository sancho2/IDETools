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
Declare Function CreateGetSet(ByRef className As String, _
						 ByRef propertyName As String, _
						 ByRef typeName As String, _
						 ByRef method As String, _
						 ByRef indexer As String, _
						 ByRef indexTYpe As String) As String 

Declare Function CreatePropertyGet(ByRef clipText As String, ByVal o as UInteger) As String 
Declare Function CreatePropertySet(ByRef clipText As String, ByVal o as UInteger) As String
Declare Function GetProcedureName(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
Declare Function GetProcedureParams(Byref txt As String, Byval o As UInteger, _
								  ByRef x1 As UInteger, ByRef x2 As UInteger) As String
Declare Function GetParameterVariable(ByRef txt As String, _
									ByRef x1 As UInteger, ByRef x2 As UInteger) As String
Declare Function GetProcedureType(ByRef txt As String, _
									ByRef x1 As UInteger, ByRef x2 As UInteger) As String

'===================================================================================================
Declare Function get_clipboard () As String
Declare Sub set_clipboard (Byref x As String)
Declare Sub HandleClipboardText()
'===================================================================================================
Function CreateGetSet(ByRef className As String, _
						 ByRef propertyName As String, _
						 ByRef typeName As String, _
						 ByRef method As String, _
						 ByRef indexer As String, _ 
						 ByRef indexType As String) As String
						 
	'
	Dim As String outS, pv = ""
	
	indexer = Trim(indexer)
	method = IIf(LCase(method) = "byref", "ByRef", "ByVal")
	indexType = IIf(LCase(indexType) = "", "UInteger", indexType)
	
	' property get
	outS = "Property " 
	outS += className + "."
	outS += propertyName
	outS += "(" 
	If indexer <> "" Then
		outS += method + " "
		outS += indexer + " As "
		outs += indexType 
	EndIf
	outS += ")"
	
	outS += " As " + typeName
	outS += CRLF
	outS += chr(9) + "'"
	outS += CRLF
	outS += Chr(9) + "Return This._" + propertyName
	If indexer <> "" Then
		outS += "("
		outS += indexer + ")"
	EndIf
	outS += CRLF
	outS += "End Property"
	outS += CRLF
	
	' property set
	outS += "Property " 
	outS += className + "." 
	outS += propertyName
	outS += "("
	If indexer <> "" Then
		outS += method + " "
		outS += indexer + " As "
		outS += indexType
		outS += ", "
	EndIf
	
	outS += method + " "
	outS += "value As " + typeName + ")"
	outS += CRLF
	outS += "'"
	outS += CRLF
	outS += Chr(9) + "This._" + propertyName
	If indexer <> "" Then
		outS += "(" 
		outS += indexer + ")"
	EndIf
	outS += " = value"
	outS += CRLF
	outS += "End Property"
	outS += CRLF
	
	? outS
	Return outS 
End Function
						  

Function GetProcedureType(ByRef txt As String, _
									ByRef x1 As UInteger, ByRef x2 As UInteger) As String
	'
	x1 = InStrRev(txt, " ") + 1
	x2 = Len(txt)
	
	Return Mid(txt, x1, x2 - x1 + 1)									
End Function

Function  CreatePropertyGet(Byref txt As String , Byval o As Uinteger ) As String 
	'
	Dim As UInteger x1, x2
	Dim As String pn, pp, outS, pv, pt
	 

	pn = GetProcedureName(txt, o, x1, x2)
	pp = GetProcedureParams(txt, o, x1, x2)
	pv = GetParameterVariable(pp, x1, x2)
	pt = GetProcedureType(txt, x1, x2)
	
	outS = "Property ~." 
	outS += pn
	outS += "(" + pp + ")"
	outS += " As " + pt
	outS += CRLF
	outS += chr(9) + "'"
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
	? outS
	Sleep 
	Return outS	

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
	outS += " = " + pv
	outS += CRLF
	outS += "End Property"
	outS += CRLF

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
	Dim As String txt, cName, pName, tName, pMethod, indexer, iType
	Dim As UByte n1, n2, n3
	txt = get_clipboard()
	'txt = moo3
	'txt = moo2
	'txt = moo1
	'txt = "~TestType moo string byref index ushort"
	'txt = "~?"
	txt = Trim(txt, Any " " + Chr(9))
'~?
	If Left(txt, 2) = "~?" OrElse Left(LCase(txt), 5) = "~help" Then
		? "MakeDefProps"
		? "------------"
		?
		? "~? or ~help: Show help window"
		? 
		? "~class_name property_name type_name [specifier] [indexer [index_type]]" 
		? "   class_name: The name of the class the property belongs to."
		? "property_name: The name of the property."
		? "    type_name: The intrinsic type of the property."
		? "    specifier: Optional byval/byref designation of the setter parameter."
		? "      indexer: Optional get/set indexer of property."
		? "   index_type: Variable type of indexer."
		Sleep
		Exit Sub
	EndIf
	 
	If Left(txt, 1) = "~" Then
		' classname property name type
		n1 = 2
		n2 = InStr(txt, " ") - 1
		cName = Mid(txt, n1, n2 - n1 + 1)
		
		n1 = n2 + 2
		n2 = InStr(n1, txt, " ") - 1
		pName = Mid(txt, n1, n2 - n1 + 1)
		
		n1 = n2 + 2		
		n2 = InStr(n1, txt, " ") - 1
		tName = Mid(txt, n1, n2 - n1 + 1)
		
		If n1 < Len(txt) Then
			n1 = n2 + 2
			n2 = InStr(n1, txt, " ") - 1
			pMethod = Mid(txt, n1, n2 - n1 + 1)
			If LCase(pMethod) = "byval" OrElse LCase(pMethod) = "byref" Then
				n1 = n2 + 2
				If n1 < Len(txt) Then
					n2 = InStr(n1, txt, " ")
					indexer = Mid(txt, n1, n2 - n1 + 1)
				EndIf
			EndIf
			
		EndIf


		set_clipboard(CreateGetSet(cName, pName, tName, pMethod, indexer, iType))
		Exit Sub
	EndIf
	
	If Left(LCase(txt), 4) = "type" Then
		' type name
		' 
	EndIf
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
	
End Sub
HandleClipboardText()
'Screenres 800, 600, 32
'Declare Sub HandleClipboardText()
'sleep

'===================================================================================================
' AutoProperty version 1.0
'===================================================================================================
#Include "windows.bi"
'===================================================================================================
Const As String CRLF = Chr(13) + Chr(10)
'===================================================================================================
Declare Function GetWord(ByRef txt As String, ByRef outS As String, ByRef n1 As Integer, ByRef n2 As Integer) As boolean
Declare Function GetWordRev(ByRef txt As String, ByRef outS As String, ByRef n1 As Integer, ByRef n2 As Integer) As boolean
'===================================================================================================
Enum AutoCommands explicit
	None 
	Getter
	Setter
	Get_Set
End Enum
Type Section
	As Integer n1
	As Integer n2
	Declare Property Length() As Integer
End Type
Property Section.Length() As Integer
	'
	Return n2 - n1 + 1
End Property
'---------------------------------------------------------------------------------------------------
Type Parameter
	As Section _specifier
	As Section _name
	As Section _array
	As Section _type
	Declare Sub SplitText(ByRef txt As String, ByVal lOffset As Integer)
End Type
Sub Parameter.SplitText(ByRef txt As String, ByVal lOffset As Integer)
	'
	Dim As String word
	Dim As Integer n1, n2
	Dim As Section Ptr sp

	With This

		' txt is single parameter
		' if the first word is byref or byval then it is a specifier
		n1 = 1
		n2 = 1
		If GetWord(txt, word, n1, n2) = FALSE Then
			' some kind of error
			Exit Sub
		EndIf  
		
		If LCase(word) = "byval" OrElse LCase(word) = "byref"Then
			._specifier = Type<Section>(n1 + lOffset, n2 + lOffset)
			
			n1 = n2 + 1
			If GetWord(txt, word, n1, n2) = FALSE Then
				Exit Sub
			EndIf
		
		EndIf
		._name = Type<Section>(lOffset + n1, lOffset + n2)
		
		n1 = InStr(txt, "(")
		If n1 <> 0 Then
			' this parameter is an array
			n2 = InStr(txt, ")")
			._array = Type<Section>(n1 + lOffset, n2 + lOffset)
		EndIf
		
		n1 = n2 + 1
		If GetWord(txt, word, n1, n2) = FALSE Then
			Exit Sub
		EndIf
		
		n1 = n2 + 1
		If GetWord(txt, word, n1, n2) = FALSE Then
			Exit Sub
		EndIf

		' we have the parameter type
		._type = Type<Section>(n1 + lOffset, n2 + lOffset)		
		
	End With
	
End Sub
'---------------------------------------------------------------------------------------------------
Type Definition
	Declare Property IsDeclarationGet() As BOOLEAN
	Declare Property IsDeclarationSet() As BOOLEAN
	Declare Property AssignmentType() As String 
	Declare Property AssignmentSpecifier() As String 
	Declare Property AssignmentVar() As String
	Declare Property Sect(ByRef index As Section) As String 
	Declare Property IsIndexed() As BOOLEAN
	Declare Property IndexerSpecifier() As String
	Declare Property IndexerType() As String
	Declare Property Indexer() As String 
	Declare Property Specifier(ByVal index As Integer) As String
	Declare Property ParamVar(ByVal index As Integer) As string 
	Declare Property ClassName() As String
	Declare Property Declaration() As String
	Declare Property Procedure() As String
	Declare Property PropertyName() As String 
	Declare Property ReturnType() As String
	Declare Property ReturnSpecifier() As String
	Declare Property AutoCommand() As AutoCommands
	
	Declare Constructor(ByRef value As String, ByRef valid As BOOLEAN)
	Declare Function GetPropertyGet() As String
	Declare Function GetPropertySet() As String 
	Declare Sub ParseCommand()

	As String txt
	
	As String _className
	As String _rem

	As AutoCommands _command
	As Section _declaration
	As Section _procedure
	As Section _name
	As Section _paramBrackets
	As Parameter _parameters(Any)
	'As section _propertyType
	As Integer _parameterCount
	As Section _return  
	As section _returnSpecifier
End Type
Property Definition.AutoCommand() As AutoCommands
	'
	Return this._command
End Property
Sub Definition.ParseCommand()
	'
	Dim As UByte n
	Dim As Integer n1, n2
	Dim As BOOLEAN okToExit = FALSE  
	
	With This
		Do  
			n += 1
			Select Case LCase(mid(._rem, n, 2))
				Case "~g"
					._command = AutoCommands.Getter
				Case "~ "
					._command = AutoCommands.Get_Set
				Case "~s"
					._command = AutoCommands.Setter
				Case "~t"
					n1 = 3
					GetWord(._rem, ._className, n1, n2)
				Case Else
					okToExit = TRUE
			End Select
			n += 1
		Loop While okToExit = FALSE 
	End With
	
End Sub
Property Definition.IsDeclarationGet() As BOOLEAN
	'
	Return cbool(this.ReturnType <> "")
End Property
Property Definition.IsDeclarationSet() As BOOLEAN
	'
	Return cbool(this.ReturnType = "")
End Property
Property Definition.AssignmentType() As String
	'
	If this.IsDeclarationGet Then
		Return this.ReturnType
	EndIf
	
	If this._parameterCount > 1 Then
		' we have everything in the parameters
		Return this.Sect(this._parameters(2)._type)
	EndIf 
	
	Return this.Sect(this._parameters(1)._type)

End Property 
Property Definition.AssignmentSpecifier() As String
	'
	If this._parameterCount > 1 Then
		' we have everything in the parameters
		Return this.Sect(this._parameters(2)._specifier)
	EndIf 
	
	If this.IsIndexed = TRUE Then
		' we can't use the parameter because its just an indexer
		' we have to get the type from the return variable
		Return This.ReturnSpecifier
	EndIf

	Return this.Sect(this._parameters(1)._specifier)
End Property 
Property Definition.AssignmentVar() As String
	'
	If this._parameterCount > 1 Then
		' we have everything in the parameters
		Return this.Sect(this._parameters(2)._name)
	EndIf 
	
	If this.IsIndexed = TRUE Then
		' we can't use the parameter because its just an indexer
		' we have to get the type from the return variable
		Return this.PropertyName
	EndIf

	Return this.Sect(this._parameters(1)._name)
End Property
Property Definition.Sect(ByRef index As Section) As String
	'
	Return Mid(this.txt, index.n1, index.Length)
End Property
Property Definition.IsIndexed() As BOOLEAN
	' TODO: i don't like the 2nd half of the return statement here
	Dim As BOOLEAN b

	If this._parameterCount > 1 Then
		Return TRUE
	EndIf
	
	If this._parameterCount = 0 Then
		Return FALSE 
	EndIf
	
	If this.Sect(this._return) <> "" then
		Return TRUE
	EndIf
	
	Return FALSE
End Property
Property Definition.IndexerSpecifier() As String
	'
	Return this.Sect(this._parameters(1)._specifier)
End Property
Property Definition.IndexerType() As String
	'
	Return this.Sect(this._parameters(1)._type)
End Property
Property Definition.Indexer() As String
	'
	Return this.ParamVar(1)
End Property 
Property Definition.ParamVar(ByVal index As Integer) As String
	'
	Return this.Sect(this._parameters(index)._name)
End Property 
Property Definition.Specifier(ByVal index As Integer) As String
	'
	Dim As Integer n1, l
	
	Return this.Sect(this._parameters(1)._specifier)
End Property 
Property Definition.ClassName() As String
	'
	return this._className 
End Property
Property Definition.PropertyName() As String
	'
	Return this.Sect(this._name)
End Property
Property Definition.ReturnType() As String
	'
	Return this.Sect(this._return)
End Property 
Property Definition.ReturnSpecifier() As String
	'
	Return this.Sect(This._returnSpecifier)
End Property 
Property Definition.Declaration() As String
	'
	Return this.Sect(this._declaration)
End Property
Property Definition.Procedure() As String
	'
	With This
		Return Mid(.txt, ._procedure.n1, ._procedure.Length)
	End With
End Property
'---------------------------------------------------------------------------------------------------
Constructor Definition(ByRef value As String, ByRef valid As BOOLEAN)
	'
#Macro __Error()
	valid = FALSE						
	Exit Constructor 
#EndMacro
	Dim As String word
	Dim As Integer n1, n2, temp
	Dim As BOOLEAN flag
	
	With This
		
		.txt = value
		
		' this block strips all but the first line in the clipboard		
		n1 = InStr(.txt, Any Chr(13) + Chr(10))
		If n1 > 0 then
			.txt = Left(.txt, n1 - 1) 
		EndIf 
		
		n2 = InStrRev(.txt, "'")
		._rem = Mid(.txt, n2 + 1) 
		
		If n2 = 0 Then
			n2 = Len(.txt) + 1 
		EndIf
		.txt = Left(.txt, n2 - 1) 

		.txt = Trim(.txt, Any " " + Chr(9))
		If LCase(Left(.txt, 7)) <> "declare" Then
			valid = FALSE
			Exit Constructor
		EndIf
		
		._declaration.n1 = 1
		._declaration.n2 = 7
		
		n1 = ._declaration.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error()
		EndIf

		If LCase(word) <> "property" Then
			__Error()
		EndIf
		
		ReDim ._parameters(1 To 2)
		._parameterCount = 0
		
		._procedure = Type<Section>(n1, n2)
		
		n1 = ._procedure.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error()
		EndIf
		
		._name = Type<Section>(n1, n2)
		
		If InStr(.txt, "(") = 0 Then
			' this might be a property get without brackets
			' in which case the next word must be as
			flag = TRUE
		EndIf

		
		n1 = ._name.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error()
		EndIf

		If cbool(LCase(word) = "as") AndAlso cbool(flag = TRUE) Then
			' this is a bracketless get
			._paramBrackets.n1 = 0
			._paramBrackets.n2 = 0
			n1 = n2 + 1
			If GetWord(.txt, word, n1, n2) = FALSE Then
				__Error()
			EndIf
			
			._return = Type<Section>(n1, n2)
			
			valid = TRUE
			Exit constructor
		EndIf
		
		' the property name may have overrun the opening bracket so we need to fix
		word = .PropertyName

		' we know there is a bracket in the definition so if the location of the bracket
		' is before n2 then it did overrun and we can back it off
		n2 = InStr(.txt, "(")
		If ._name.n2 > n2 Then
			n2 -= 1
			._name = Type<Section>(._name.n1, n2)
		EndIf

		n1 = InStr(._name.n2 + 1, .txt, "(")		' left parameter bracket
		n2 = InStrRev(.txt, ")")						' right parameter bracket 				

		._paramBrackets = Type<Section>(n1, n2)

		' at the most there will be two parameters
		' if there are 2 then there will be a comma separating them
		' its the only place where a comma can occur
		n2 = InStr(txt, ",") 
		If n2 > 0 Then
			' there is a comma so we can assume two parameters
			._parameterCount += 1
			' the first one is between the ( and the , 
			n1 = ._paramBrackets.n1 + 1
			n2 -= 1
			word = Mid(txt, n1, n2 - n1 + 1)
			._parameters(1).SplitText(word, n1 - 1)
			n2 += 2
		Else
			n2 = n1 + 1
		EndIf 

		' n2 is the first char after the ( or after the comma  
		n1 = n2 
		n2 = 1
		If GetWord(txt, word, n1, n2) = FALSE Then 
			__Error()
		EndIf

		' if word <> ) then there is a parameter
		If word <> ")" Then
			' since there can only be one more parameter we can
			' use the ) location to close out our parameter string
			word = Mid(.txt, n1, ._paramBrackets.n2 - n1)
			
			._parameterCount += 1
			._parameters(._parameterCount).SplitText(word, n1 - 1)

		EndIf

		' Property type can be found after the parameter )
		n1 = ._paramBrackets.n2 + 1
		n2 = 1

		GetWord(txt, word, n1, n2) 
		If LCase(word) = "byref" Then
			._returnSpecifier = Type<Section>(n1, n2)
			n1 = n2 + 1
		EndIf

		If LCase(word) <> "as" Then
			__Error()
		EndIf

		n1 = n2 + 1
		GetWord(txt, word, n1, n2)
		._return = Type<Section>(n1, n2)

		this.ParseCommand()

	End With

	valid = TRUE 

End Constructor
'---------------------------------------------------------------------------------------------------
Function Definition.GetPropertySet() As String 
	'
	Dim As String outS
	
	With This

		._className = IIf(.ClassName = "", "~", .ClassName) 
		outS = "Property " + .ClassName					' class name 
		outS = outS + "." + .PropertyName + "("		' property name

		If .IsIndexed = TRUE Then			
			outS = outS + .IndexerSpecifier + " "			' indexer specifier
			outS = outS + .Indexer + " As "					' indexer
			outs = outS + .IndexerType + ", "				' indexer type
		EndIf
		
		If .AssignmentSpecifier = "" Then				' parameter specifier
			outS += "ByVal"
		Else
			outS = outS + .AssignmentSpecifier
		EndIf
		outS += " value As "					
		outS = outS + .AssignmentType + ")" + CRLF + "'"				' parameter type

		outS += CRLF + Chr(9) + "This._"  
		outS = outS + .PropertyName							' member name
		
		If .IsIndexed = TRUE  Then
			outS = outS + "(" + .Indexer + ")"				' memeber index 
		EndIf
		
		outS += " = value" + CRLF + "End Property" + CRLF

	End With

	Return outS
	
End Function
Function Definition.GetPropertyGet() As String
	'
	Dim As String outS

	With This
		._className = IIf(.ClassName = "", "~", .ClassName) 
		outS = "Property " + .ClassName					' class name 
		outS = outS + "." + .PropertyName + "("		' property name

		If .IsIndexed = TRUE Then			
			outS = outS + .IndexerSpecifier + " "			' indexer specifier
			outS = outS + .Indexer + " As "					' indexer
			outs = outS + .IndexerType 						' indexer type
		EndIf

		outS = outS + ")" + " As " + .ReturnType						' property type
		outS += CRLF + chr(9) + "'" + CRLF
		outS = outS + Chr(9) + "Return This._" + .PropertyName	' return line

		If .IsIndexed = TRUE then			
			outS = outS + "(" + indexer + ")"		' return indexer
		EndIf
		
		outS += CRLF + "End Property" + CRLF

	End With

	Return outS

End Function
'===================================================================================================
Declare Sub TestMe(ByRef d As Definition)
Declare Sub TestTest()
Declare Sub ReadClipboard()
Declare Function get_clipboard () As String
Declare Sub set_clipboard (Byref x As String)
'===================================================================================================
Function GetWordRev(ByRef txt As String, ByRef outS As String, ByRef n1 As Integer, ByRef n2 As Integer) As boolean
	'
	If Len(txt) < 1 OrElse n1 > Len(txt) Then
		outS = ""
		Return FALSE
	EndIf
	
	If n2 = 0 Then 
		n2 = Len(txt)
	EndIf
	While Mid(txt, n2, 1) = " "	' track back until a non-space char
		n2 -= 1
	Wend

	n1 = InStrRev(txt, " ", n2) 
	If n1 = 0 Then
		n1 = Len(txt)
	Else
		n1 += 1
	EndIf
	
	outS = Mid(txt, n1, n2 - n1 + 1)
	Return TRUE
	
End Function
Function GetWord(ByRef txt As String, ByRef outS As String, ByRef n1 As Integer, ByRef n2 As Integer) As boolean
	'
	If Len(txt) < 1 OrElse n1 > Len(txt) Then
		outS = ""
		Return FALSE
	EndIf
	
	If n1 = 0 Then 
		n1 = 1
	EndIf
	While Mid(txt, n1, 1) = " "
		n1 += 1
	Wend

	n2 = InStr(n1, txt, " ") - 1
	If n2 < 1 Then
		n2 = Len(txt) 
	EndIf

	outS = Mid(txt, n1, n2 - n1 + 1)
	Return TRUE
	
End Function
'===================================================================================================
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
'===================================================================================================
Sub Main()
	'
	Dim As String txt, outS

	txt = get_clipboard()

	Dim b As BOOLEAN	
	Dim d As Definition = Definition(txt, b)	
	
	If b = FALSE Then
		Exit Sub 
	EndIf
	
	Select Case d.AutoCommand
		Case AutoCommands.None
			If d.IsDeclarationGet Then
				outS = d.GetPropertyGet()
			Else
				outS = d.GetPropertySet()
			EndIf
		Case AutoCommands.Getter
			outS = d.GetPropertyGet()
		Case AutoCommands.Setter
			outS = d.GetPropertySet()
		Case AutoCommands.Get_Set
			outS = d.GetPropertyGet()
			outS += d.GetPropertySet()
	End Select

	set_clipboard(outS)
End Sub

Main()

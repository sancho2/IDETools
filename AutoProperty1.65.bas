'===================================================================================================
' AutoProperty version 1.65
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
Enum ErrorHandling explicit
	None
	SuppressAll
	ShowConsoleWindow
	WriteToClipboard
End Enum
Type Section
	As Integer n1
	As Integer n2
	Declare Property Length() As Integer
	Declare Operator Let(ByVal rhs As Integer)
End Type
Operator Section.Let(ByVal rhs As Integer)	
'
	this.n1 = rhs
	this.n2 = rhs
End Operator 
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
	Declare Function SplitText(ByRef txt As String, ByVal lOffset As Integer) As BOOLEAN
	Declare Operator Let(ByVal rhs As Integer)
End Type
Operator Parameter.Let(ByVal rhs As integer)
	'
	With This
		._specifier = rhs
		._name = rhs
		._array = rhs
		._type = rhs
	End With 
End Operator 
Function Parameter.SplitText(ByRef txt As String, ByVal lOffset As Integer) As BOOLEAN 
	'
	Dim As String word
	Dim As Integer n1, n2
	Dim As Section Ptr sp

	With This

		' txt is single parameter text
		' if the first word is byref or byval then it is a specifier
		n1 = 1
		n2 = 1
		' This word is either a specifier or a param var name
		If GetWord(txt, word, n1, n2) = FALSE Then
			' some kind of error
			Return FALSE
		EndIf  
		
		If LCase(word) = "byval" OrElse LCase(word) = "byref"Then
			._specifier = Type<Section>(n1 + lOffset, n2 + lOffset)
			
			n1 = n2 + 1
			If GetWord(txt, word, n1, n2) = FALSE Then
				Return FALSE 
			EndIf
		
		EndIf
		
		' Word should be a parameter name
		._name = Type<Section>(lOffset + n1, lOffset + n2)
		
		n1 = InStr(txt, "(")
		If n1 <> 0 Then
			' this parameter is an array
			n2 = InStr(txt, ")")
			._array = Type<Section>(n1 + lOffset, n2 + lOffset)
		EndIf
		
		' retrieve the AS in the paramter
		n1 = n2 + 1
		If GetWord(txt, word, n1, n2) = FALSE Then
			Return FALSE 
		EndIf
		If LCase(word) <> "as" Then
			Return FALSE 
		EndIf
		
		n1 = n2 + 1
		If GetWord(txt, word, n1, n2) = FALSE Then
			Return FALSE 
		EndIf

		' we have the parameter type
		._type = Type<Section>(n1 + lOffset, n2 + lOffset)		
		
	End With

	Return TRUE 

End Function 
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
	Declare Property AutoCommand(ByVal c As AutoCommands)
	
	Declare Constructor()
	Declare Sub Init()
	Declare Function GetDeclareSet() As String 
	Declare Function GetDeclareGet() As String 
	Declare Function GetPropertyGet() As String
	Declare Function GetPropertySet() As String 
	Declare Sub ParseCommand()
	Declare Function ParseForPropertyDeclaration(ByRef value As String, ByRef errMsg As String) As BOOLEAN
	Declare Function ParseMemberVariable(ByRef value As String, ByRef errMsg As String) As BOOLEAN

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
Function Definition.GetDeclareSet() As String 
	'
	Dim As String outS, temp
	Dim As Integer n1
	Dim As BOOLEAN flag
	
	With This
		temp = .PropertyName

		n1 = InStr(temp, "(")
		If n1 > 0 Then
			temp = Left(temp, Len(temp) - n1 + 1)
			flag = TRUE  
		EndIf
		
		' strip the _ from the declaration
		If Left(temp, 1) = "_" Then
			temp = Mid(temp, 2)
			Mid(temp, 1, 1) = UCase(Mid(temp, 1, 1))
		EndIf
		outS = "Declare Property " + temp + "("
		If flag = TRUE Then
			outS += "ByVal index as integer, "
			
		EndIf
		outS = outS + "ByVal value As " + .ReturnType
		outS = outS + ")"  

	End With
	
	Return outS

End Function
Function Definition.GetDeclareGet() As String
	'
	Dim As String outS, temp
	Dim As Integer n1
	Dim As BOOLEAN flag
	
	With This
		temp = .PropertyName
		n1 = InStr(temp, "(")
		If n1 > 0 Then
			temp = Left(temp, Len(temp) - n1 + 1)
			flag = TRUE  
		EndIf
		
		' strip the _ from the declaration
		If Left(temp, 1) = "_" Then
			temp = Mid(temp, 2)
			Mid(temp, 1, 1) = UCase(Mid(temp, 1, 1))
		EndIf

		outS = "Declare Property " + temp + "("
		If flag = TRUE Then
			outS += "ByVal index as integer"
		EndIf
		outS = outS + ")" + " As " + .ReturnType 

	End With
	
	Return outS
	
End Function 

Sub Definition.Init()
	'
	With This
		.txt = ""
		._className = ""
		._rem = ""
		._command = AutoCommands.None
		._declaration =0
		._procedure = 0
		._name = 0
		._paramBrackets = 0
		For x As UByte = 1 To ._parameterCount
			_parameters(x) = 0
		Next 
		._parameterCount = 0
		_return = 0  
		._returnSpecifier = 0
	End With 
End Sub
Function Definition.ParseMemberVariable(ByRef value As String, ByRef errMsg As String) As BOOLEAN
	'
#Macro __Error(e)
	errMsg = e
	Return FALSE						
#EndMacro
#Macro __ReturnTrue()
	this.ParseCommand()
	Return TRUE
#EndMacro

	Dim As String word
	Dim As Integer n1, n2, temp
	Dim As BOOLEAN flag
	
	With This
		.Init()
		.txt = value
		
		' this block strips all but the first line in the clipboard		
		n1 = InStr(.txt, Any Chr(13) + Chr(10))
		If n1 > 0 then
			.txt = Left(.txt, n1 - 1) 
		EndIf 
		
		' this finds the rem char at the end of the line and stores that text 
		n2 = InStrRev(.txt, "'")
		._rem = Mid(.txt, n2 + 1) 
		
		If n2 = 0 Then
			n2 = Len(.txt) + 1 
		EndIf
		.txt = Left(.txt, n2 - 1)		' strip everything after the rem 

		' remove trailing spaces and tabs
		.txt = Trim(.txt, Any " " + Chr(9))
		
		' remove dim if it exists
		If LCase(Left(.txt, 4)) = "dim " Then
			.txt = Mid(.txt, 5)
		EndIf 
		
		' this block searches for the required 'as'
		n1 = 1
		Do
			If GetWord(.txt, word, n1, n2) = FALSE Then
				__Error("Clipboard text is not a valid member variable declaration")
			EndIf
			If LCase(word) = "as" Then
				Exit Do
			EndIf
			word = ""
			n1 = n2 + 1
		Loop While n1 < Len(.txt)
		
		If word = "" Then
			__Error("'AS' not found. The clipboard text is not a valid member variable declaration")
		EndIf 

		' n1 now contains the location of the 'a' in 'as'
		' if as is not the first character then we are going to move the 
		' text prior to it to the end of the line. This way we always
		' have as [typename] [variablename]
		word = Trim(Left(.txt, n1 - 1))
		If word <> "" Then 
			.txt = Mid(.txt, n1, Len(txt) - n1 + 1) + " " + word
		EndIf
		
		' next is the member type which should start at 4+			
		n1 = 4
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error("Error in member type. The text in the clipboard is not a valid member declaration.")
		EndIf 
		 
		._return = Type<Section>(n1, n2)
		
		' next is the member variable name
		n1 = n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error("Error in member variable. The text in the clipboard is not a valid member declaration.")
		EndIf 
		
		._name = Type<Section>(n1, n2)		
		
		' we know that n1 is the first char of var name
		' if this is an array then it may have the '(' inside the name
		n1 = InStr(word, "(")
		If n1 = 0 Then
			' there can be a space before an array '('
			n1 = InStr(n2, .txt, "(")
		EndIf
		
		If n1 <> 0 Then
			' we have found an (
			' we have an array 
			' all we are going to do is add the entire (...) into the var name
			n2 = InStr(.txt, ")")
			._name.n2 = n2
		EndIf

	End With

	__ReturnTrue()

End Function
Function Definition.ParseForPropertyDeclaration(ByRef value As String, ByRef errMsg As String) As BOOLEAN
	'
#Macro __Error(e)
	errMsg = e
	Return FALSE						
#EndMacro
#Macro __ReturnTrue()
	this.ParseCommand()
	Return TRUE
#EndMacro
	Dim As String word
	Dim As Integer n1, n2, temp
	Dim As BOOLEAN flag
	
	With This
		.Init()
		.txt = value
		
		' this block strips all but the first line in the clipboard		
		n1 = InStr(.txt, Any Chr(13) + Chr(10))
		If n1 > 0 then
			.txt = Left(.txt, n1 - 1) 
		EndIf 
		
		' this finds the rem char at the end of the line and stores that text 
		n2 = InStrRev(.txt, "'")
		If n2 <> 0 Then
			._rem = Mid(.txt, n2 + 1)
		EndIf
		
		If n2 = 0 Then
			n2 = Len(.txt) + 1 
		EndIf
		.txt = Left(.txt, n2 - 1)		' strip everything after the rem 

		' remove trailing spaces and tabs
		.txt = Trim(.txt, Any " " + Chr(9))
		
		If LCase(Left(.txt, 7)) <> "declare" Then
			' this is not a property declaration line
			__Error("Clipboard text does not begin with 'declare'")
		EndIf
		
		._declaration.n1 = 1
		._declaration.n2 = 7
		
		' retrieve the first word after 'declare'
		n1 = ._declaration.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error("No text found after the 'Declare'")
		EndIf

		If LCase(word) <> "property" Then
			__Error("The text after 'declare' does not contain the word 'property'")
		EndIf
		
		ReDim ._parameters(1 To 2)
		._parameterCount = 0
		
		._procedure = Type<Section>(n1, n2)
		
		' retrieve the first word after 'property' - the property name
		n1 = ._procedure.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error("No text found after 'property'")
		EndIf
		
		._name = Type<Section>(n1, n2)
		
		If InStr(.txt, "(") = 0 Then
			' this might be a property get without brackets
			' in which case the next word must be as
			flag = TRUE
		EndIf

		' retrieve the first word after the property name		
		n1 = ._name.n2 + 1
		If GetWord(.txt, word, n1, n2) = FALSE Then
			__Error("No text found after the property name")
		EndIf

		If cbool(LCase(word) = "as") AndAlso cbool(flag = TRUE) Then
			' this is a bracketless get
			._paramBrackets.n1 = 0
			._paramBrackets.n2 = 0
			n1 = n2 + 1
			If GetWord(.txt, word, n1, n2) = FALSE Then
				__Error("No return type for property get procedure")
			EndIf
			
			._return = Type<Section>(n1, n2)
			
			__ReturnTrue()		' macro to return

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
		n2 = InStr(.txt, ",")
		If n2 > 0 Then
			' there is a comma so we can assume two parameters
			._parameterCount += 1
			' the first one is between the ( and the , 
			n1 = ._paramBrackets.n1 + 1
			n2 -= 1
			word = Mid(.txt, n1, n2 - n1 + 1)
			._parameters(1).SplitText(word, n1 - 1)
			n2 += 2
		Else
			n2 = n1 + 1
		EndIf 

		' n2 is the first char after the ( or after the comma  
		n1 = n2 
		n2 = 1
		If GetWord(.txt, word, n1, n2) = FALSE Then 
			__Error("Invalid paramters in property declartion")
		EndIf

		' if word <> ) then there is a parameter
		If word <> ")" Then
			' since there can only be one more parameter we can
			' use the ) location to close out our parameter string
			word = Mid(.txt, n1, ._paramBrackets.n2 - n1)

			._parameterCount += 1
			If ._parameters(._parameterCount).SplitText(word, n1 - 1) = FALSE Then
				._parameterCount -= 1		' probably not necessary as we are erroring out
				__Error("Invalid paramter text")
			EndIf
		EndIf

		' Property type can be found after the parameter )
		n1 = ._paramBrackets.n2 + 1
		n2 = 1

		GetWord(.txt, word, n1, n2) 
		If LCase(word) = "byref" Then
			._returnSpecifier = Type<Section>(n1, n2)
			n1 = n2 + 1
		EndIf
		If LCase(word) <> "as" Then
			If word = "" Then
				' this is a setter
				__ReturnTrue()
			EndIf
			__Error("Invalid specifier in property type")
		EndIf

		n1 = n2 + 1
		GetWord(.txt, word, n1, n2)
		._return = Type<Section>(n1, n2)

	End With

	__ReturnTrue() 
	
End Function
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
Property Definition.AutoCommand(ByVal c As AutoCommands)
	'
	this._command = c
End Property
Property Definition.AutoCommand() As AutoCommands
	'
	Return this._command
End Property
Property Definition.IsDeclarationGet() As BOOLEAN
	'
	If this._parameterCount = 0 Then
		Return TRUE 
	EndIf

	If this._parameterCount > 1 Then
		Return FALSE 
	EndIf
	
	If this._parameterCount = 1 AndAlso this.ReturnType <> "" Then
		Return TRUE
	EndIf 

	Return FALSE 

End Property
Property Definition.IsDeclarationSet() As BOOLEAN
	'
	Return cbool(this.IsDeclarationGet = FALSE)
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
		' we have to get the name from the return variable
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
Constructor Definition()

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
Declare Sub Main()
Declare Sub ParseCommandLine(ByRef eh As ErrorHandling)
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
	If Len(outS) < 1 Then
		Return FALSE 
	EndIf

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
Sub ParseCommandLine(ByRef eh As ErrorHandling)
	'
	Dim As UByte i = 1
	Do
		Dim As String arg = Command(i)
		If Len(arg) < 1 Then
			Exit Sub
		EndIf
		
		Select Case arg
			Case "-es"	
				eh = ErrorHandling.SuppressAll
			Case "-ew"
				eh = ErrorHandling.ShowConsoleWindow
			Case "-ec"
				eh = ErrorHandling.WriteToClipboard
		End Select
		i += 1
	Loop
End Sub 
Sub HandleError(ByVal eh As ErrorHandling, ByRef e As String)
	'
	Dim As String outS = get_clipboard()
	Select Case eh
		Case ErrorHandling.SuppressAll
			Exit Sub
		Case ErrorHandling.ShowConsoleWindow
			? "Error processing clipboard text: "
			?
			? e
			Sleep
		Case ErrorHandling.WriteToClipboard
			outS = "' Error processing clipboard text: " + CRLF + CRLF + "'" + e + CRLF + CRLF + outS 
			set_clipboard(outS)			 
	End Select
End Sub
Sub Main()
	'
	Dim As Definition d 
	Dim As String txt, outS, e, com, getS, setS 
	Dim As ErrorHandling eh = ErrorHandling.WriteToClipboard
	Dim As BOOLEAN b
	 
	ParseCommandLine(eh)
	txt = get_clipboard()
	
	b = d.ParseForPropertyDeclaration(txt, e)  
	If b = FALSE Then
		' check for member variable declaration
		b = d.ParseMemberVariable(txt, e)
		' todo: change this to produce a proper error message (combined)
		If b = FALSE Then
			HandleError(eh, e)
			Exit Sub
		EndIf 
		getS = d.GetDeclareGet()
		setS = d.GetDeclareSet()
		
		b = d.ParseForPropertyDeclaration(getS, e)
		If b = FALSE Then
			HandleError(eh, e)
		EndIf
		outS = d.GetPropertyGet()
		outS += crlf 
		b = d.ParseForPropertyDeclaration(setS, e)
		If b = FALSE Then
			HandleError(eh, e)
		EndIf
		outS = outS + d.GetPropertySet()
		outS = getS + CRLF + setS + CRLF + outS + CRLF

		set_clipboard(outS)
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




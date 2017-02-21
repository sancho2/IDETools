#Define IDD_DIALOG			1000
#Define lstNames  	1001
#Define lstRGB 		1002
#Define scbVert 		1003
#Define cmdAdd 		1007
#Define lstOut			1006
#Define cmdExit 		1008
#Define cmdRemove 	1009
#Define cmdInsert		1010
#Define cmdCopy		1011
#Define txtPrefix 	1012

#Define swatches 		1200

#Define IDM_MENU				10000
#Define IDM_FILE_EXIT		10001
#Define IDM_HELP_ABOUT		10101

Const As String CRLF = Chr(13) + Chr(10)

Dim Shared hInstance As HMODULE
Dim Shared CommandLine As ZString Ptr
Dim Shared hWnd As HWND

Const ClassName="DLGCLASS"
Const AppName="Dialog as main"
Const AboutMsg=!"FbEdit Dialog as main\13\10Copyright © FbEdit 2007"
'--------------------------------------------------------------------------------------------------------------------
Declare Sub PadNumberString(ByRef txt As String, ByVal amt As UByte = 3)
'--------------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
#define RGBA_R( c ) ( CUInt( c ) Shr 16 And 255 )
#define RGBA_G( c ) ( CUInt( c ) Shr  8 And 255 )
#define RGBA_B( c ) ( CUInt( c )        And 255 )
'----------------------------------------------------------------------------------------------------------------
Type CColor
	As String Name
	Declare Property ToRGBString() As String
	Declare Property ToRGBHex() As String
	Declare Property ToBGRHex() As String
	Declare Property ToRGBValue() As ULong
	Declare Property ToBGRValue() As ULong
	Declare Sub SplitRGBString(ByRef r As String, ByRef g As String, ByRef b As string)
	As ULong value
End Type
Property CColor.ToRGBString() As String
	'
	Dim As String s, r, g, b
	
	this.SplitRGBString(r, g, b)
	PadNumberString(r)
	PadNumberString(g)
	PadNumberString(b)
	s = r + ", " + g + ", " + b 
	Return s
End Property

Sub CColor.SplitRGBString(ByRef r As String, ByRef g As String, ByRef b As string)
	'
	r = Str(RGBA_R(this.value))
	g = Str(RGBA_G(this.value))
	b = Str(RGBA_B(this.value))
End Sub 

Property CColor.ToRGBHex() As String
	'
	Return Hex(this.value, 8)
End Property
	
Property CColor.ToBGRHex() As String
	'
	Dim As ULong v = this.ToBGRValue
	Return Hex(v, 8)
End Property

Property CColor.ToRGBValue() As ULong
	'
	Return this.value
End Property

Property CColor.ToBGRValue() As ULong
	'
	Return RGBA(RGBA_B(this.value), RGBA_G(this.value),RGBA_R(this.value), 0) 
End Property
'--------------------------------------------------------------------------------------------------------------------
Dim Shared As CColor colors(Any)
'--------------------------------------------------------------------------------------------------------------------
Declare Sub LoadColors()
'--------------------------------------------------------------------------------------------------------------------
Sub LoadColors()
	'
	Dim As Integer n
	Dim As String cName
	Dim As ULong cValue
	
	Restore ColorData
	Read cName, cValue
	While cName <> ""
		Dim As CColor c
		n += 1
		ReDim Preserve colors(1 To n)
		With colors(n)
			.Name = cName
			.value = cValue
		End With
		Read cName, cValue
	Wend 
End Sub
'--------------------------------------------------------------------------------------------------------------------
Sub PadNumberString(ByRef txt As String, ByVal amt As UByte = 3)
	'
	While Len(txt) < amt
		txt = "0" + txt 
	Wend
End Sub

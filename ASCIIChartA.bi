
dim SHARED hInstance as HINSTANCE
dim SHARED hooks as ADDINHOOKS
dim SHARED lpHandles as ADDINHANDLES ptr
dim SHARED lpFunctions as ADDINFUNCTIONS ptr
dim SHARED lpData as ADDINDATA ptr

Dim Shared As Integer IDM_ASCIICHART

Declare Function RightAlignNumber(ByVal n As UByte, ByVal pad As UByte = 3) As String
Declare Sub LoadNoPrintAscii()
Declare Function GetAsciiChar(ByVal value As UByte) As String
'-------------------------------------------------
#Define dlgTest 1100
#Define cmdExit 1104
#Define cmdCopy 1105
#Define cmdPage 1107
#Define cmdInsert 1108
#Define lblIndex 1106
#Define lblCopy 1103
#Define lblValue 1102

#Define cmdButtons 1210


Dim Shared As UByte aStart = 0
'Dim Shared hInstance As HMODULE
Dim Shared As String NoPrintAscii(0 To 31)
Sub LoadNoPrintAscii()
	'
	Dim As String s
	Restore LowAscii
	For x As UByte = 0 To 31
		Read NoPrintAscii(x)
		Read s
	Next
End Sub
Function GetAsciiChar(ByVal value As UByte) As String
	'
	Dim As String s
	If value < 32 Then
		s = NoPrintAscii(value) 
	Else
		s = Chr(value)
	EndIf
	
	While Len(s) < 3
		s = " " + s
	Wend
	Return s
End Function
Function RightAlignNumber(ByVal n As UByte, ByVal pad As UByte = 3) As String
	'
	Dim As String s = Str(n)
	While Len(s) < 3
		s = " " + s
	Wend  
	Return s
End Function

LowAscii:
Data "NUL", "Null", "SOH", "Start of Heading", "STX", "Start of Text", "ETX", "End of Text"
Data "EOT", "End of Transmission", "ENQ", "Enquiry", "ACK", "Acknowledgment", "BEL", "Bell"
Data " BS", "Back Space", " HT", "Horizontal Tab", " LF", "Line Feed", " VT", "Vertical Tab"
Data " FF", "Form Feed", " CR", "Carriage Return", " SO", "Shift Out / X-On", " SI", "Shift In / X-Off"
Data "DLE", "Data Line Escape", "DC1", "Device Control 1 (oft. XON)", "DC2", "Device Control 2", "DC3", "Device Control 3 (oft. XOFF)"
Data "DC4", "Device Control 4", "NAK", "Negative Acknowledgement", "SYN", "Synchronous Idle", "ETB", "End of Transmit Block"
Data "CAN", "Cancel", " EM", "End of Medium", "SUB", "Substitute", "ESC", "Escape"
Data " FS", "File Separator", " GS", "Group Separator", " RS", "Record Separator", " US", "Unit Separator"

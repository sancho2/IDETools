'"TTF to GfxLib Font Converter.bas"
' slightly modded @jofer's code
' use fb versions 0.17, 0.18, or >=0.21
'
' saveas:     TTFtoGfxLib.bas
' save .rc as TTFtoGfxLib.rc
' compile -s gui TTFtoGfxLib.bas TTFtoGfxLib.rc
'
#include "windows.bi"
#include "win/commdlg.bi"
#include "crt.bi"
'
#define DLG_MAIN 1000
#define BTN_SAVE 1001
#define BTN_FONT 1002
#define BTN_BKGR 1003
#define CHK_TRAN 1004
#define LBL_MAIN 1005
'Zbegin change - see .rc
#define IDC_EDT1 1006
#define IDC_EDT2 1007
'Zend  change
#define ICO_MAIN 1010
'
Type Global
    BackgroundColor  As COLORREF
    BitmapInfo       As BITMAPINFO
    ChooseColor      As CHOOSECOLOR
    CustColors(15)   As COLORREF
    ChooseFont       As CHOOSEFONT
    hBrush           As HBRUSH
    hBrushPrev       As HBRUSH
    hFont            As HFONT
    hFontPrev        As HFONT
    IsTransparent    As Integer
    LogFont          As LOGFONT
    OpenFileName     As OPENFILENAME
    SaveName         As String*255
    SaveFilter       As String*255
End Type

Declare Function DlgProc (ByVal hWnd As HWND, ByVal uMsg As UINT, BYVal wParam As WPARAM, ByVal lParam As LPARAM) As BOOL
Declare Sub EditColor(ByVal hWnd As HWND)
Declare Sub EditFont(ByVal hWnd As HWND)
Declare Sub SaveFile(ByVal hWnd As HWND)
Declare Sub SaveBitmap(ByVal hWnd As HWND, ByVal hDC As HDC, ByVal hBitmap As HBITMAP, ByVal File As String)

Dim Shared Global As Global
dim shared as integer lowChar,highChar

DialogBox(GetModuleHandle(NULL), MAKEINTRESOURCE(DLG_MAIN), NULL, @DlgProc)
End

Function DlgProc (ByVal hWnd As HWND, ByVal uMsg As UINT, BYVal wParam As WPARAM, ByVal lParam As LPARAM) As BOOL
    Dim ID As Integer
    Dim Event As Integer
    Dim hDC As HDC
    Dim hBrush As HBRUSH
    Dim x As Integer
    Dim SaveFilterArray As UByte Ptr
    Dim hWndControl As HWND

'Zb
    dim as integer res
    dim as string lowstr,highstr,errstr
'Ze

    Select Case uMsg
        Case WM_INITDIALOG
            With Global
                ' Set up all structure information
                .SaveFilter = "Bitmap Files (*.bmp)%*.BMP%All Files (*.*)%*.*%%"
                SaveFilterArray = StrPtr(.SaveFilter)
                For x = 0 To Len(.SaveFilter)
                    If SaveFilterArray[x] = Asc("%") Then SaveFilterArray[x] = 0
                Next x

                .BackgroundColor = GetSysColor(COLOR_WINDOW)
                .hBrush = GetSysColorBrush(COLOR_WINDOW)
                .IsTransparent = TRUE
                memset(@Global.CustColors(0), 255, SizeOf(COLORREF)*16)
                With .ChooseColor
                    .lStructSize = SizeOf(CHOOSECOLOR)
                    .hWndOwner = hWnd
                    .rgbResult = Global.BackgroundColor
                    .lpCustColors = @Global.CustColors(0)
                    .flags = CC_ANYCOLOR Or CC_RGBINIT Or CC_SOLIDCOLOR
                End With

                With .ChooseFont
                    .lStructSize = SizeOf(CHOOSEFONT)
                    .hInstance = GetModuleHandle(NULL)
                    .hWndOwner = hWnd
                    .lpLogFont = @Global.LogFont
                    .Flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
                End With

                With .OpenFileName
                    .lStructSize = SizeOf(OPENFILENAME)
                    .hInstance = GetModuleHandle(NULL)
                    .hWndOwner = hWnd
                    .lpstrFilter = StrPtr(Global.SaveFilter)
                    .lpstrFile = StrPtr(Global.SaveName)
                    .nMaxFile = 255
                End WIth

                With .LogFont
                    .lfFaceName = "MS Sans Serif"
                    .lfHeight = -MulDiv(8, GetDeviceCaps(GetDC(NULL), LOGPIXELSY), 72)
                End With

                Global.hFont = CreateFontIndirect(Global.ChooseFont.lpLogFont)
            End With

            CheckDlgButton(hWnd, CHK_TRAN, BST_CHECKED)
            Global.IsTransparent = TRUE

        Case WM_CLOSE
            EndDialog(hWnd, 0)

        Case WM_CTLCOLORSTATIC
            If GetDlgCtrlID(Cast(hWND, lParam)) = LBL_MAIN Then
                hDC = Cast(HDC, wParam)
                SetTextColor(hDC, Global.ChooseFont.rgbColors)

                If Global.IsTransParent = TRUE Then
'Zb
'  fix transparent background
'
                    'SetBkColor(hDC, RGB(255, 255, 255))
                    SetBkColor(hDC, RGBA(255, 0, 255, 0))
'Ze
                    Return Cast(LRESULT, GetSysColorBrush(COLOR_WINDOW))
                Else
                    SetBkColor(hDC, Global.BackgroundColor)
                    Return Cast(LRESULT, Global.hBrush)
                End If
            End If
        Case WM_COMMAND
            ID = LoWord(wParam)
            Event = HiWord(wParam)

            Select Case ID
                Case BTN_SAVE
                    '
                    lowstr=space(16):highstr=space(16)
                    res=GetDlgItemText(hwnd,IDC_EDT1,(lowstr),16)
                    res=GetDlgItemText(hwnd,IDC_EDT2,(highstr),16)
                    '
                    lowChar =val(trim(lowstr))
                    highChar=val(trim(highstr))
                    '
                    errstr=""
                    if lowChar <0 or lowChar>255 then
                        errstr="MinChar <0 or >255"
                    end if
                    if errstr<>"" then
                        messagebox(hwnd,errstr,"Error",MB_ICONSTOP)
                        return false
                    end if
                    if lowChar>highChar then
                        errstr="MinChar > MaxChar"
                    end if
                    if errstr<>"" then
                        messagebox(hwnd,errstr,"Error",MB_ICONSTOP)
                        return false
                    end if
                    '
                    if highChar <0 or highChar>255 then
                        errstr="MaxChar <0 or >255"
                    end if
                    if errstr<>"" then
                        messagebox(hwnd,errstr,"Error",MB_ICONSTOP)
                        return false
                    end if
                    '
                    SaveFile(hWnd)
                Case BTN_FONT
                    EditFont(hWnd)
                Case BTN_BKGR
                    EditColor(hWnd)
                Case CHK_TRAN
                    hWndControl = GetDlgItem(hWnd, BTN_BKGR)

                    If IsDlgButtonChecked(hWnd, CHK_TRAN) = BST_CHECKED Then
                        Global.IsTransparent = TRUE
                        Global.LogFont.lfQuality = NONANTIALIASED_QUALITY

                        ' Enable the 'background color' button
                        EnableWindow(hWndControl, FALSE)
                    Else
                        Global.IsTransParent = FALSE
                        Global.LogFont.lfQuality = ANTIALIASED_QUALITY

                        ' Disable the 'background color' button
                        EnableWindow(hWndControl, TRUE)
                    End If

                    Global.hFontPrev = Global.hFont
                    Global.hFont = CreateFontIndirect(@Global.LogFont)
                    SendDlgItemMessage(hWnd, LBL_MAIN, WM_SETFONT, Cast(WPARAM, Global.hFont), TRUE)
                    DeleteObject(Global.hFontPrev)
            End Select

        Case Else
            Return FALSE
    End Select

    Return TRUE
End Function

Sub EditFont(ByVal hWnd As HWND)
    ' If a font is chosen, set all the variables
    If ChooseFont(@Global.ChooseFont) = TRUE Then
        Global.hFontPrev = Global.hFont
        If Global.IsTransparent = TRUE Then
            Global.LogFont.lfQuality = NONANTIALIASED_QUALITY
        Else
            Global.LogFont.lfQuality = ANTIALIASED_QUALITY
        End If

        Global.hFont = CreateFontIndirect(Global.ChooseFont.lpLogFont)
        SendDlgItemMessage(hWnd, LBL_MAIN, WM_SETFONT, Cast(WPARAM, Global.hFont), TRUE)
        DeleteObject(Global.hFontPrev)
    End If
End Sub

Sub EditColor(ByVal hWnd As HWND)
    Dim hWndControl As HWND

    If ChooseColor(@Global.ChooseColor) = TRUE Then
        Global.BackgroundColor = Global.ChooseColor.rgbResult
        Global.hBrushPrev = Global.hBrush
        Global.hBrush = CreateSolidBrush(Global.BackgroundColor)
        DeleteObject(Global.hBrushPrev)

        hWndControl = GetDlgItem(hWnd, LBL_MAIN)
        InvalidateRect(hWndControl,NULL,TRUE)
        UpdateWindow(hWndControl)
    End If
End Sub

Sub SaveFile(ByVal hWnd As HWND)
    Dim MemDC As HDC
    Dim MemBMP As HBITMAP
    Dim BitmapWidth As Integer
    Dim BitmapHeight As Integer
    Dim BitmapSize As Integer
    Dim BitmapInfo As BITMAPINFO
    Dim FileName As String
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    Dim WidthArray(lowChar To highChar) As Integer
    Dim ABCArray(lowChar To highChar) As ABC
    Dim ThisWidth As SIZE
    Dim TextMetric As TEXTMETRIC
    Dim IsTrueType As Integer

    Dim Buffer As UByte Ptr

    ' If a save name is chosen...
    If GetSaveFileName(@Global.OpenFileName) Then

        ' Create a memory DC and select our font into it
        MemDC = CreateCompatibleDC(NULL)
        If (MemDC = 0) Or (Global.hFont = 0) Then
            MessageBox(hWnd, "Could not create Bitmap", "Error", MB_ICONERROR)
            Exit Sub
        End If
        SelectObject(MemDC, Global.hFont)

        ' Get character widths
        IsTrueType = GetCharABCWidths(MemDC, lowChar, highChar, @ABCArray(lowChar))

        For i = lowChar To highChar
            GetTextExtentPoint32(MemDC, Chr(i), 1, @ThisWidth)
            If IsTrueType Then
                WidthArray(i) = ABCArray(i).abcB
                If ABCArray(i).abcC > 0 Then WidthArray(i) += ABCArray(i).abcC
                If ABCArray(i).abcA > 0 Then WidthArray(i) += ABCArray(i).abcA
            Else
                WidthArray(i) = ThisWidth.cx
            End If
            BitmapWidth += WidthArray(i)
            If ThisWidth.cy > BitmapHeight Then BitmapHeight = ThisWidth.cy
        next i

'Zb
'  handle descenders, "gjp" etc.
'
        if highChar>=97 then
            BitmapHeight += 1
        end if
'Ze
        BitmapSize = BitmapHeight * BitmapWidth * 4

        ' Create DIB section & select it into memory DC
        With Global.BitmapInfo.bmiHeader
            .biSize = SizeOf(BITMAPINFOHEADER)
            .biWidth = BitmapWidth
            .biHeight = BitmapHeight
            .biPlanes = 1
            .biBitCount = 32
            .biCompression = BI_RGB
        End With

        MemBMP = CreateDIBSection(MemDC, @Global.BitmapInfo, DIB_RGB_COLORS, @Buffer, NULL, 0)
        SelectObject MemDC, MemBMP
        If Global.IsTransparent = True Then
'Zb
'  fix transparent background
'
            'SetBkColor(MemDC, RGB(255, 255, 255))
            SetBkColor(MemDC, RGBA(255, 0, 255, 0))
'Ze
        Else
            SetBkColor(MemDC, Global.BackgroundColor)
        End If
        SetTextColor(MemDC, Global.ChooseFont.rgbColors)

        If MemBMP = 0 Or Buffer = 0 Then
            MessageBox(hWnd, "Could not create Bitmap", "Error", MB_ICONERROR)
            DeleteObject(MemDC)
            Exit Sub
        End If

        'Fill in font info and draw letters
        Buffer[BitmapSize-BitmapWidth*4] = 0
        Buffer[BitmapSize-BitmapWidth*4+1] = lowChar
        Buffer[BitmapSize-BitmapWidth*4+2] = highChar

        x = 0
        For i = lowChar To highChar
            If ABCArray(i).abcA < 0 Then x -= ABCArray(i).abcA
            TextOut(MemDC, x, 1, Chr(i), 1)
'Zb
'  accomodate different character ranges
'
'            Buffer[BitmapSize-BitmapWidth*4 + i - 29] = WidthArray(i)
            Buffer[BitmapSize-BitmapWidth*4 + i - lowChar + 3] = WidthArray(i)
'Ze
            If ABCArray(i).abcA < 0 Then x += ABCArray(i).abcA
            x += WidthArray(i)
        Next i

        FileName = *Global.OpenFileName.lpstrFile
        If UCase(Right(FileName, 4)) <> ".BMP" Then FileName += ".bmp"

        SaveBitmap(hWnd, MemDC, MemBMP, FileName)

        MessageBox hWnd, "Complete!", "Hurray!", MB_OK
    End If
End Sub

Sub SaveBitmap(ByVal hWnd As HWND, ByVal hDC As HDC, ByVal hBitmap As HBITMAP, ByVal File As String)
    Dim fp As FILE Ptr
    Dim Bitmap As BITMAP
    Dim BitmapInfo As BITMAPINFO
    Dim BitmapFileHeader As BITMAPFILEHEADER
    Dim Buffer As UByte Ptr
    fp = fopen(File, "wb")
    If fp = 0 Then
        MessageBox hWnd, "Error Saving Bitmap", "Error", MB_ICONERROR
        Exit Sub
    End If

    BitmapInfo.bmiHeader.biSize = SizeOf(BITMAPINFOHEADER)
    BitmapInfo.bmiHeader.biBitCount = 0

    If GetDIBits(hDC, hBitmap, 0, 0, NULL, @BitmapInfo, DIB_RGB_COLORS) = 0 Then
        MessageBox hWnd, "GetDIBits Error Saving Bitmap", "Error", MB_ICONERROR
        fclose(fp)
        Exit Sub
    End If

    Bitmap.bmHeight = BitmapInfo.bmiHeader.biHeight
    Bitmap.bmWidth = BitmapInfo.bmiHeader.biWidth

    With BitmapFileHeader
        .bfType = &h4d42
        .bfSize = (((3 * Bitmap.bmWidth + 3) And Not 3) * Bitmap.bmHeight)
        .bfSize += SizeOf((BITMAPFILEHEADER))
        .bfSize += SizeOf(BITMAPINFOHEADER)
        .bfOffBits = SizeOf((BITMAPFILEHEADER)) + SizeOf(BITMAPINFOHEADER)
    End With

    BitmapInfo.bmiHeader.biCompression = 0
    fwrite(@BitmapFileHeader, SizeOf((BITMAPFILEHEADER)),1,fp)
    fwrite(@BitmapInfo.bmiHeader, SizeOf(BITMAPINFOHEADER),1,fp)

    Buffer = Allocate(BitmapInfo.bmiHeader.biSizeImage + 5)
    If GetDIBits(hDC, hBitmap, 0, Bitmap.bmHeight, Buffer, @BitmapInfo, DIB_RGB_COLORS) = 0 Then
        Deallocate Buffer
        MessageBox hWnd, "Error Saving Bitmap", "Error", MB_ICONERROR
        fclose(fp)
        Exit Sub
    End If

    fwrite(Buffer,1,BitmapInfo.bmiHeader.biSizeImage,fp)
    fclose(fp)
End Sub

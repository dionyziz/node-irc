Attribute VB_Name = "mGraphics"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
'API Functions
    'Mouse Stuff
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function GetCapture Lib "user32" () As Long
    '//
    
    'System Color Stuff
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
    Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
    Private Const CLR_INVALID = -1
    Private Const COLOR_HIGHLIGHT = 13
    Private Const COLOR_BTNFACE = 15
    Private Const COLOR_BTNSHADOW = 16
    Private Const COLOR_BTNTEXT = 18
    Private Const COLOR_BTNHIGHLIGHT = 20
    Private Const COLOR_BTNDKSHADOW = 21
    Private Const COLOR_BTNLIGHT = 22
    '//
    
    'Text Stuff
    Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Private Const DT_CALCRECT = &H400
    Private Const DT_WORDBREAK = &H10
    Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4
    Private Const DT_WORD_ELLIPSIS = &H40000
    '//
    
    'Graphics Stuff
    Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
    Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
    Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Type RGBTRIPLE
            rgbBlue As Byte
            rgbGreen As Byte
            rgbRed As Byte
    End Type
    Private Type BITMAPINFOHEADER
            biSize As Long
            biWidth As Long
            biHeight As Long
            biPlanes As Integer
            biBitCount As Integer
            biCompression As Long
            biSizeImage As Long
            biXPelsPerMeter As Long
            biYPelsPerMeter As Long
            biClrUsed As Long
            biClrImportant As Long
    End Type
    Private Type BITMAPINFO
            bmiHeader As BITMAPINFOHEADER
            bmiColors As RGBTRIPLE
    End Type
    Private Const FXDEPTH As Long = &H28
    Private Const PS_SOLID = 0
    '//
    
    'Reigons and Rects
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
    Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
    Private Const RGN_DIFF = 4
    
    Public Type RECT
            Left As Long
            Top As Long
            Right As Long
            Bottom As Long
    End Type
    
    Public Type POINTAPI
            x As Long
            y As Long
    End Type
    '//
'END API Stuff //

Public Function AdjustToOLE_COLOR(ByVal oClr As Variant, Optional OffSetColor As Long = 0, Optional hPal As Long = 0) As OLE_COLOR
Dim iRGB As Long
Dim R As Long
Dim G As Long
Dim B As Long
    If OleTranslateColor(oClr, hPal, iRGB) Then
        AdjustToOLE_COLOR = CLR_INVALID
    End If
    
    R = ((iRGB And &HFF&) + OffSetColor)
    G = (((iRGB And &HFF00&) \ &H100) + OffSetColor)
    B = (((iRGB And &HFF0000) \ &H10000) + OffSetColor)
    
    If R < 0 Then
        R = 0
    Else
    If R > 255 Then
        R = 255
    End If
    End If

    If G < 0 Then
        G = 0
    Else
    If G > 255 Then
        G = 255
    End If
    End If

    If B < 0 Then
        B = 0
    Else
    If B > 255 Then
        B = 255
    End If
    End If
    AdjustToOLE_COLOR = RGB(R, G, B)
End Function

Public Sub DrawGradient(DestDC As Long, oStartColor As OLE_COLOR, oEndColor As OLE_COLOR, rc As RECT, iHeightWidth As Long, Optional iDirection As Long = 0)
Dim dR(1 To 3) As Double
Dim iStep As Long
Dim i As Long
Dim bRGB(1 To 3) As Integer
Dim hBr As Long
Dim RGBStartCol1(1 To 3) As OLE_COLOR
Dim RGBEndCol1(1 To 3) As OLE_COLOR
Dim iColor As Long

    OleTranslateColor oEndColor, 0, iColor
    RGBStartCol1(1) = iColor And &HFF&
    RGBStartCol1(2) = ((iColor And &HFF00&) \ &H100)
    RGBStartCol1(3) = ((iColor And &HFF0000) \ &H10000)

    OleTranslateColor oStartColor, 0, iColor
    RGBEndCol1(1) = iColor And &HFF&
    RGBEndCol1(2) = ((iColor And &HFF00&) \ &H100)
    RGBEndCol1(3) = ((iColor And &HFF0000) \ &H10000)

    iStep = iHeightWidth \ 255
    If (iStep = 0) Then
        iStep = 1
    End If
    bRGB(1) = RGBStartCol1(1)
    bRGB(2) = RGBStartCol1(2)
    bRGB(3) = RGBStartCol1(3)
    dR(1) = RGBEndCol1(1) - RGBStartCol1(1)
    dR(2) = RGBEndCol1(2) - RGBStartCol1(2)
    dR(3) = RGBEndCol1(3) - RGBStartCol1(3)
    For i = iHeightWidth To 0 Step -iStep
        If iDirection = 0 Then
            rc.Top = rc.Bottom - iStep
        Else
            rc.Left = rc.Right - iStep
        End If
        hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
        FillRect DestDC, rc, hBr
        DeleteObject hBr
        If iDirection = 0 Then
            rc.Bottom = rc.Top
        Else
            rc.Right = rc.Left
        End If
        bRGB(1) = RGBStartCol1(1) + dR(1) * (iHeightWidth - i) / iHeightWidth
        bRGB(2) = RGBStartCol1(2) + dR(2) * (iHeightWidth - i) / iHeightWidth
        bRGB(3) = RGBStartCol1(3) + dR(3) * (iHeightWidth - i) / iHeightWidth
    Next i
End Sub

Public Sub DrawASquare(DestDC As Long, rc As RECT, oColor As OLE_COLOR, Optional bFillRect As Boolean)
Dim iBrush As Long
Dim i(0 To 3) As Long
oColor = AdjustToOLE_COLOR(oColor)
    i(0) = rc.Top
    i(1) = rc.Left
    i(2) = rc.Right
    i(3) = rc.Bottom
    iBrush = CreateSolidBrush(oColor)
    If bFillRect = True Then
        FillRect DestDC, rc, iBrush
    Else
        FrameRect DestDC, rc, iBrush
    End If
    DeleteObject iBrush
End Sub

Public Sub DrawALine(DestDC As Long, x As Long, y As Long, X1 As Long, Y1 As Long, oColor As OLE_COLOR, Optional iWidth As Long = 1)
Dim pt As POINTAPI
Dim iPen As Long
Dim iPen1 As Long

    iPen = CreatePen(PS_SOLID, iWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    
    MoveToEx DestDC, x, y, pt
    LineTo DestDC, X1, Y1

    SelectObject DestDC, iPen1
    DeleteObject iPen
End Sub

Public Sub DrawADot(DestDC As Long, x As Long, y As Long, oColor As OLE_COLOR)
    SetPixel DestDC, x, y, oColor
End Sub

Public Function RoundCorners(iHwnd As Long, iRadius As Long) As Long
Dim rc As RECT
Dim iRgn As Long
    DeleteObject iRgn
    GetClientRect iHwnd, rc
    iRgn = CreateRoundRectRgn(rc.Left, rc.Top - 1, rc.Right + 1, rc.Bottom + 1, iRadius, iRadius)
    SetWindowRgn iHwnd, iRgn, True
    RoundCorners = iRgn
End Function

Public Function GetRect(iHwnd As Long) As RECT
    GetClientRect iHwnd, GetRect
End Function

Public Sub CleanDC(iHDC As Long)
Dim i As Long
    i = DeleteDC(iHDC)
End Sub

Public Sub SetTheTextColor(DestDC As Long, oColor As OLE_COLOR)
    SetTextColor DestDC, oColor
End Sub

Public Sub DrawTheText(DestDC As Long, sText As String, iTextLength As Long, rc As RECT, DTF As Long)
    DrawText DestDC, sText, iTextLength, rc, DTF
End Sub

Public Sub GetTextRect(DestDC As Long, sText As String, iTextLength As Long, rc As RECT)
    DrawText DestDC, sText, iTextLength, rc, DT_CALCRECT
End Sub

Public Sub CopyTheRect(DestinationRECT As RECT, SourceRECT As RECT)
    CopyRect DestinationRECT, SourceRECT
End Sub

Public Sub PositionRect(rc As RECT, ByVal x As Long, ByVal y As Long)
    OffsetRect rc, x, y
End Sub

Public Sub DeleteObjectReference(ByVal iReference As Long)
    DeleteObject iReference
End Sub

Public Function SetAccessKey(sCaption As String) As String
Dim iPos As Long
    If Len(sCaption) > 0 Then
        iPos = InStr(1, sCaption, "&", vbTextCompare)
        If iPos < Len(sCaption) And iPos > 0 Then
            If Mid(sCaption, iPos + 1, 1) <> "&" Then
                SetAccessKey = LCase$(Mid$(sCaption, iPos + 1, 1))
            Else
                iPos = InStr(iPos + 2, sCaption, "&", vbTextCompare)
                If Mid(sCaption, iPos + 1, 1) <> "&" Then
                    SetAccessKey = LCase(Mid(sCaption, iPos + 1, 1))
                End If
            End If
        End If
    End If
End Function

Public Sub TabNext(Optional TabBack As Boolean = False)
    If TabBack = False Then
        SendKeys "{TAB}"
    Else
        SendKeys "+{TAB}"
    End If
End Sub

Public Function ResizeRect(rc As RECT, X1, Y1) As RECT
    InflateRect rc, X1, Y1
    ResizeRect = rc
End Function

Public Sub ClearRect(rc As RECT)
    SetRectEmpty rc
End Sub

Public Function MouseCoordinatesIsOver(hwnd As Long, Optional mX As Long, Optional mY As Long) As Boolean
    'Determine if a mouse is currently over the control or not.
    Dim pt As POINTAPI
    GetCursorPos pt
    mX = pt.x
    mY = pt.y
    MouseCoordinatesIsOver = (WindowFromPoint(pt.x, pt.y) = hwnd)
End Function

Public Function MouseOver(hwnd As Long) As Boolean
    If MouseCoordinatesIsOver(hwnd) = True Then
        If GetCapture <> hwnd Then
            SetCapture hwnd
        End If
        MouseOver = True
    Else
        If GetCapture = hwnd Then
            ReleaseCapture
        End If
        MouseOver = False
    End If
End Function

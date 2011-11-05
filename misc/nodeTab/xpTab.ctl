VERSION 5.00
Begin VB.UserControl xpTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   MaskColor       =   &H00974D37&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picBlank 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   1575
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1095
      Width           =   705
      Begin VB.PictureBox pLeftUp 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   90
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   255
      End
      Begin VB.PictureBox pRightDown 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   60
         Width           =   255
      End
   End
End
Attribute VB_Name = "xpTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

    Private Const DT_WORDBREAK = &H10
    Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4
    Private Const DT_WORD_ELLIPSIS = &H40000

    Dim cPic As cImageManipulation
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Const WM_LBUTTONDOWN = &H201
'Tab Alignment
    Public Enum eTabAlignment
        tTop = 0
        tLeft = 1
        tRight = 2
        tBottom = 3
    End Enum
    Private eTab As eTabAlignment
'//
Private mMouseOver                  As Boolean
Private mMouseOverLScroll           As Boolean
Private mMouseOverRScroll           As Boolean
Private iSelectedTab                As Long
Private iHotTab                     As Long
Private MouseInBody                 As Boolean
Private MouseInTab                  As Boolean
Private hasFocus                    As Boolean
Private sAccessKeys                 As String
Private iPrevTab                    As Integer
Private iFirstVisibleTab            As Long
Private iLastVisibleTab             As Long
Private bLeftIsDown                 As Boolean
Private bRightIsDown                As Boolean
Private iMaxTabWidth                As Long
Private iTabHeight                  As Long
Private StopScrolling               As Boolean
Private Tabs()                      As New cTabs
Dim rcTabs()                        As RECT
Dim rcBody                          As RECT
Dim MouseLeft                       As Long
Dim MouseTop                        As Long
'Property Variables
    Private oBackColor              As OLE_COLOR
    Private oForeColor              As OLE_COLOR
    Private oActiveForeColor        As OLE_COLOR
    Private oForeColorHot           As OLE_COLOR
    Private oFrameColor             As OLE_COLOR
    Private oMaskColor              As OLE_COLOR
    Private oScrollColor            As OLE_COLOR
    Private oScrollBackColor        As OLE_COLOR
    Private oTabHotStripColor       As OLE_COLOR
    Private oForeColorDisabled      As OLE_COLOR
    Private iNumberOfTabs           As Long

'Events
Event TabPressed(PreviousTab As Integer)
Event MouseIn(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseOut(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TabDblClick(Index As Long)
Event DblClick()
    
Private Sub pLeftUp_Click()
    'SendMessage UserControl.hwnd, WM_LBUTTONDOWN, 0, 0
End Sub

Private Sub pLeftUp_DblClick()
    'SendMessage UserControl.hwnd, WM_LBUTTONDOWN, 0, 0
End Sub

Private Sub pLeftUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'If iFirstVisibleTab > 1 Then
        bLeftIsDown = True
        iFirstVisibleTab = iFirstVisibleTab - 1
        DrawTab
        Sleep 500
        DoEvents
        Do Until bLeftIsDown = False
            Sleep (1)
            iFirstVisibleTab = iFirstVisibleTab - 1
            DrawTab
            DoEvents
        Loop
    'End If
End Sub

Private Sub pLeftUp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MouseOver(pLeftUp.hwnd) = True Then
        If mMouseOverLScroll = True Then Exit Sub
        mMouseOverLScroll = True
        DisplayScrollButtons
    Else
        mMouseOverLScroll = False
        DisplayScrollButtons
    End If
End Sub

Private Sub pLeftUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bLeftIsDown = False
End Sub

Private Sub pRightDown_Click()
    'SendMessage UserControl.hwnd, WM_LBUTTONDOWN, 0, 0
End Sub

Private Sub pRightDown_DblClick()
    'SendMessage UserControl.hwnd, WM_LBUTTONDOWN, 0, 0
End Sub

Public Sub pRightDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If StopScrolling = True Then Exit Sub
    bRightIsDown = True
    If iLastVisibleTab = iNumberOfTabs Then Exit Sub
    iFirstVisibleTab = iFirstVisibleTab + 1
    DrawTab
    Sleep 500
    DoEvents
    Do Until bRightIsDown = False
        Sleep (1)
        iFirstVisibleTab = iFirstVisibleTab + 1
        DrawTab
        DoEvents
    Loop
End Sub

Private Sub pRightDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MouseOver(pRightDown.hwnd) = True Then
        If mMouseOverRScroll = True Then Exit Sub
        mMouseOverRScroll = True
        DisplayScrollButtons
    Else
        mMouseOverRScroll = False
        DisplayScrollButtons
    End If
    
End Sub

Private Sub pRightDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bRightIsDown = False
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Dim i As Long
    'Return the tab number that repersents the
    'Access Key.
    i = GetTabAccessKey(KeyAscii)
    If Tabs(i).TabEn = True Then
        SelectedTab = i
    End If
End Sub

Private Sub UserControl_EnterFocus()
    hasFocus = True
    DrawTab
End Sub

Private Sub UserControl_InitProperties()
    AddTab
    iFirstVisibleTab = 1
    iTabHeight = 22
    iSelectedTab = 1
    'oBackColor = AdjustToOLE_COLOR(UserControl.Parent.BackColor)
    UserControl.BackColor = oBackColor
    oScrollBackColor = oBackColor
    pLeftUp.BackColor = oScrollBackColor
    pRightDown.BackColor = oScrollBackColor
    oForeColor = AdjustToOLE_COLOR(vbButtonText)
    UserControl.ForeColor = oForeColor
    oActiveForeColor = RGB(56, 80, 152)
    oForeColorHot = RGB(0, 0, 255)
    oFrameColor = AdjustToOLE_COLOR(oBackColor, -40)
    oForeColorDisabled = AdjustToOLE_COLOR(oFrameColor)
    oScrollColor = AdjustToOLE_COLOR(oFrameColor, -20)
    oMaskColor = RGB(255, 0, 255)
    oTabHotStripColor = RGB(232, 144, 40)
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_Terminate()
    Erase Tabs
    Erase rcTabs
End Sub

Private Sub UserControl_GotFocus()
    hasFocus = True
    'DrawTab
End Sub

Private Sub UserControl_LostFocus()
    hasFocus = False
    DrawTab
    
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++
'Begin, UserControl Mouse Stuff.
'+++++++++++++++++++++++++++++++++++++++++++++++
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim hRgn            As Long
Dim i               As Long
Dim iTabRight       As Long
Dim iTabBottom      As Long
Dim iBlankLeft      As Long
Dim iBlankTop       As Long
    MouseLeft = CLng(x)
    MouseTop = CLng(y)
    If Button = 1 Then
        For i = iFirstVisibleTab To iLastVisibleTab
        'Create a region to hit test our tab.
        hRgn = CreateRectRgnIndirect(rcTabs(i))
        '//
            'If PtInRegion = True then do the if block.
            If PtInRegion(hRgn, CLng(x), CLng(y)) Then
                'First check that the tab is Enabled
                If Tabs(i).TabEn = True Then
                    If picBlank.Visible = True Then
                        'Check to see if the tab is under the scroll
                        'buttons, Meaning partially visible.
                        Select Case eTab
                            Case 0, 3
                                iTabRight = rcTabs(i).Right
                                iBlankLeft = picBlank.Left
                                If iTabRight > iBlankLeft Then
                                    'Its under the scroll so shunt it over.
                                    iFirstVisibleTab = iFirstVisibleTab + 1
                                End If
                            Case 1, 2
                                iTabBottom = rcTabs(i).Bottom
                                iBlankTop = picBlank.Top
                                If iTabBottom > iBlankTop Then
                                    'Its under the scroll so shunt it upwards.
                                    iFirstVisibleTab = iFirstVisibleTab + 1
                                End If
                        End Select
                    End If
                    DeleteObjectReference hRgn
                    iSelectedTab = i
                    SelectedTab = i
                    Exit For
                End If
                '//
            Else
            'Not in the tab.
                DeleteObjectReference hRgn
                RaiseEvent MouseDown(Button, Shift, x, y)
            End If
            '//
        Next i
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim hRgn            As Long
Dim i               As Long
Dim iLocalHotTab    As Long
Dim DoRedraw        As Boolean
    If MouseOver(UserControl.hwnd) = True Then
        RaiseEvent MouseIn(Button, Shift, x, y)
        hRgn = CreateRectRgnIndirect(rcBody)
        If PtInRegion(hRgn, CLng(x), CLng(y)) Then
            iHotTab = 0
            If MouseInBody = False Then
                MouseInTab = False
                MouseInBody = True
                DrawTab
                DeleteObjectReference hRgn
                Exit Sub
            End If
        Else
            DeleteObjectReference hRgn
        End If
    
        For i = iFirstVisibleTab To iLastVisibleTab
        hRgn = CreateRectRgnIndirect(rcTabs(i))
            If PtInRegion(hRgn, CLng(x), CLng(y)) Then
                iLocalHotTab = i
                If iLocalHotTab <> iHotTab Then
                    If Tabs(i).TabEn = True Then
                        DoRedraw = True
                        MouseInTab = True
                        MouseInBody = False
                        iHotTab = i
                        DeleteObjectReference hRgn
                    End If
                End If
            Else
                DeleteObjectReference hRgn
            End If
        Next i
        If DoRedraw = True Then
            DrawTab
        End If
    Else
        RaiseEvent MouseOut(Button, Shift, x, y)
        iHotTab = 0
        DrawTab
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_DblClick()
Dim x As Long
Dim y As Long
Dim hRgn As Long
Dim i As Long
Dim bInTab As Boolean
    'Did this happen in the tab part
    'of the control, or the body
    For i = iFirstVisibleTab To iLastVisibleTab
        hRgn = CreateRectRgnIndirect(rcTabs(i))
        If PtInRegion(hRgn, MouseLeft, MouseTop) Then
            If Tabs(i).TabEn = True Then
                bInTab = True
                RaiseEvent TabDblClick(i)
            End If
            DeleteObjectReference hRgn
            Exit For
        Else
            DeleteObjectReference hRgn
        End If
    Next i
    If bInTab = False Then
        RaiseEvent DblClick
    End If
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++
'End, UserControl Mouse Stuff.
'+++++++++++++++++++++++++++++++++++++++++++++++

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft, vbKeyUp
            If iSelectedTab = 1 Then
                SelectedTab = 1
            Else
                SelectedTab = SelectedTab - 1
            End If
        Case vbKeyRight, vbKeyDown
            If iSelectedTab = iNumberOfTabs Then
                SelectedTab = iNumberOfTabs
            Else
                SelectedTab = SelectedTab + 1
            End If
        Case vbKeyPageUp
            If picBlank.Visible = True Then
                pLeftUp_MouseDown 0, 0, 0, 0
            End If
        Case vbKeyPageDown
            If picBlank.Visible = True Then
                pRightDown_MouseDown 0, 0, 0, 0
            End If
    End Select
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyPageDown
            pRightDown_MouseUp 0, 0, 0, 0
        Case vbKeyPageUp
            pLeftUp_MouseUp 0, 0, 0, 0
    End Select
End Sub

Private Sub UserControl_Resize()
    DrawTab
End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++'
'Begin Properties.
'+++++++++++++++++++++++++++++++++++++++++++++++++++'
Public Property Get TabHeight() As Long
    TabHeight = iTabHeight
End Property

Public Property Let TabHeight(ByVal NewTabHeight As Long)
    iTabHeight = NewTabHeight
    If iTabHeight < 22 Then
        iTabHeight = 22
    End If
    PropertyChanged "TabHeight"
    DrawTab
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = oBackColor
End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    oBackColor = AdjustToOLE_COLOR(NewBackColor)
    UserControl.BackColor = oBackColor
    PropertyChanged "BackColor"
    DrawTab
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = oForeColor
End Property

Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
    oForeColor = AdjustToOLE_COLOR(NewForeColor)
    UserControl.ForeColor = oForeColor
    PropertyChanged "ForeColor"
    DrawTab
End Property

Public Property Get ForeColorActive() As OLE_COLOR
    ForeColorActive = oActiveForeColor
End Property

Public Property Let ForeColorActive(ByVal NewForeColorActive As OLE_COLOR)
    oActiveForeColor = AdjustToOLE_COLOR(NewForeColorActive)
    PropertyChanged "ForeColorActive"
    DrawTab
End Property

Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = oForeColorDisabled
End Property

Public Property Let ForeColorDisabled(ByVal NewDissColor As OLE_COLOR)
    oForeColorDisabled = AdjustToOLE_COLOR(NewDissColor)
    PropertyChanged "ForeColorDisabled"
    DrawTab
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = oFrameColor
End Property

Public Property Let FrameColor(ByVal NewFrameColor As OLE_COLOR)
    oFrameColor = AdjustToOLE_COLOR(NewFrameColor)
    PropertyChanged "FrameColor"
    DrawTab
End Property

Public Property Get ScrollArrowColor() As OLE_COLOR
    ScrollArrowColor = oScrollColor
End Property

Public Property Let ScrollArrowColor(ByVal NewScrollColor As OLE_COLOR)
    oScrollColor = AdjustToOLE_COLOR(NewScrollColor)
    PropertyChanged "ScrollArrowColor"
    DrawTab
End Property

Public Property Get TabWidth(ByVal Index As Long) As Long
    TabWidth = Tabs(Index).TabWidth
End Property

Public Property Let TabWidth(ByVal Index As Long, ByVal NewTabWidth As Long)
    If NewTabWidth >= UserControl.ScaleWidth Then
        NewTabWidth = UserControl.ScaleWidth - 20
    End If
    Tabs(Index).TabWidth = NewTabWidth
    PropertyChanged "TabWidth"
    DrawTab
End Property

Public Property Get AutoSize(ByVal Index As Long) As Boolean
    AutoSize = Tabs(Index).TabAutoSize
End Property

Public Property Let AutoSize(ByVal Index As Long, ByVal NewAutoSize As Boolean)
    Tabs(Index).TabAutoSize = NewAutoSize
    PropertyChanged "AutoSize"
    DrawTab
End Property

Public Property Get TabEnabled(ByVal Index As Long) As Boolean
    TabEnabled = Tabs(Index).TabEn
End Property

Public Property Let TabEnabled(ByVal Index As Long, ByVal NewTabEnabled As Boolean)
    Tabs(Index).TabEn = NewTabEnabled
    PropertyChanged "TabEnabled"
    DrawTab
End Property

Public Property Get TabCaption(ByVal Index As Long) As String
    TabCaption = Tabs(Index).TabCaption
End Property

Public Property Let TabCaption(ByVal Index As Long, ByVal NewTabCaption As String)
    Tabs(Index).TabCaption = NewTabCaption
    SetTabAccessKeys
    PropertyChanged "TabCaption"
    DrawTab
End Property

Public Property Get TabPicture(ByVal Index As Long) As StdPicture
    Set TabPicture = Tabs(Index).TabIcon
End Property

Public Property Set TabPicture(ByVal Index As Long, ByVal NewTabPicture As StdPicture)
    Set Tabs(Index).TabIcon = NewTabPicture
    PropertyChanged "TabPicture"
    DrawTab
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal newFont As Font)
    Set UserControl.Font = newFont
    PropertyChanged "Font"
    DrawTab
End Property

Public Property Get MaskColor() As OLE_COLOR
    MaskColor = oMaskColor
End Property

Public Property Let MaskColor(ByVal NewMaskColor As OLE_COLOR)
    oMaskColor = AdjustToOLE_COLOR(NewMaskColor)
    PropertyChanged "MaskColor"
    DrawTab
End Property

Public Property Get NumberOfTabs() As Long
    NumberOfTabs = iNumberOfTabs
End Property

Public Property Get Alignment() As eTabAlignment
    Alignment = eTab
End Property

Public Property Let Alignment(ByVal NewAlignment As eTabAlignment)
    eTab = NewAlignment
    PropertyChanged "Alignment"
    DrawTab
End Property

Public Property Get ForeColorHot() As OLE_COLOR
    ForeColorHot = oForeColorHot
End Property

Public Property Let ForeColorHot(ByVal NewForeColorHot As OLE_COLOR)
    oForeColorHot = AdjustToOLE_COLOR(NewForeColorHot)
    PropertyChanged "ForeColorHot"
End Property

Public Property Get SelectedTab() As Long
    SelectedTab = iSelectedTab
End Property

Public Property Let SelectedTab(ByVal NewSelectedTab As Long)
    iSelectedTab = NewSelectedTab
    RaiseEvent TabPressed(iPrevTab)
    iPrevTab = iSelectedTab
    PropertyChanged "SelectedTab"
    DrawTab
End Property

Public Property Get BackColorScroll() As OLE_COLOR
    BackColorScroll = oScrollBackColor
End Property

Public Property Let BackColorScroll(ByVal NewScrollColor As OLE_COLOR)
    oScrollBackColor = AdjustToOLE_COLOR(NewScrollColor)
    DrawTab
    PropertyChanged "BackColorScroll"
End Property

Public Property Get TabHotStripColor() As OLE_COLOR
    TabHotStripColor = oTabHotStripColor
End Property

Public Property Let TabHotStripColor(ByVal NewStripColor As OLE_COLOR)
    oTabHotStripColor = AdjustToOLE_COLOR(NewStripColor)
    PropertyChanged "TabHotStripColor"
    DrawTab
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long
    'On Error GoTo EH:
    With PropBag
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        Alignment = .ReadProperty("Alignment", eTab)
        TabHeight = .ReadProperty("TabHeight", 22)
        BackColor = .ReadProperty("BackColor", RGB(240, 240, 224))
        BackColorScroll = .ReadProperty("BackColorScroll", oScrollBackColor)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorActive = .ReadProperty("ForeColorActive", RGB(56, 80, 152))
        ForeColorHot = .ReadProperty("ForeColorHot", RGB(0, 0, 255))
        ForeColorDisabled = .ReadProperty("ForeColorDisabled", oForeColorDisabled)
        FrameColor = .ReadProperty("FrameColor", RGB(152, 160, 160))
        ScrollArrowColor = .ReadProperty("ScrollArrowColor", oScrollColor)
        MaskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
        TabHotStripColor = .ReadProperty("TabHotStripColor", RGB(232, 144, 40))
        SelectedTab = .ReadProperty("SelectedTab", 1)
        iNumberOfTabs = .ReadProperty("NumberOfTabs", 1)
        ReDim Tabs(iNumberOfTabs) As New cTabs
    End With
    For i = 1 To iNumberOfTabs
        With Tabs(i)
            .TabAutoSize = PropBag.ReadProperty("AutoSize" & i)
            .TabWidth = PropBag.ReadProperty("TabWidth" & i)
            .TabCaption = PropBag.ReadProperty("TabText" & i)
            .TabEn = PropBag.ReadProperty("TabEnabled" & i)
            Set .TabIcon = PropBag.ReadProperty("TabPicture" & i, Nothing)
        End With
    Next i
    SetTabAccessKeys
'Exit Sub
'EH:
'Err.Clear
'Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
    With PropBag
        .WriteProperty "Alignment", eTab
        .WriteProperty "TabHeight", iTabHeight
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "BackColorScroll", oScrollBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorActive", oActiveForeColor
        .WriteProperty "ForeColorHot", oForeColorHot
        .WriteProperty "ForeColorDisabled", oForeColorDisabled
        .WriteProperty "FrameColor", oFrameColor
        .WriteProperty "ScrollArrowColor", oScrollColor
        .WriteProperty "MaskColor", oMaskColor
        .WriteProperty "TabHotStripColor", oTabHotStripColor
        .WriteProperty "SelectedTab", iSelectedTab
        .WriteProperty "Font", UserControl.Font
        .WriteProperty "NumberOfTabs", iNumberOfTabs
    End With
    For i = 1 To iNumberOfTabs
        With Tabs(i)
            PropBag.WriteProperty "TabPicture" & i, .TabPicture, 0
            PropBag.WriteProperty "AutoSize" & i, .TabAutoSize
            PropBag.WriteProperty "TabWidth" & i, .TabWidth
            PropBag.WriteProperty "TabText" & i, .TabCaption
            PropBag.WriteProperty "TabEnabled" & i, .TabEn
        End With
    Next i
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++'
'End Properties.
'+++++++++++++++++++++++++++++++++++++++++++++++++++'

'+++++++++++++++++++++++++++++++++++++++++++++++++++'
'Begin Control Procedures.
'+++++++++++++++++++++++++++++++++++++++++++++++++++'
Public Sub DrawTab()
    PlaceTabs
End Sub

Public Function AddTab(Optional iTabWidth As Long = 60, _
                        Optional sTabText As String = "", _
                        Optional bTabEn As Boolean = True, _
                        Optional bAutoSize As Boolean = True, _
                        Optional pTabPicture As StdPicture = Nothing) As Long
    
    iNumberOfTabs = iNumberOfTabs + 1
    ReDim Preserve Tabs(iNumberOfTabs) As New cTabs
        With Tabs(iNumberOfTabs)
            .TabWidth = iTabWidth
            .TabAutoSize = bAutoSize
            .TabEn = bTabEn
            If sTabText = "" Then
                sTabText = "Tab" & iNumberOfTabs
                .TabCaption = sTabText
            Else
                .TabCaption = sTabText
            End If
            Set .TabIcon = pTabPicture
        End With
        PropertyChanged "TabPicture"
        PropertyChanged "NumberOfTabs"
        AddTab = iNumberOfTabs
        DrawTab
End Function

Public Function DeleteTab()
    If iNumberOfTabs > 1 Then
        iNumberOfTabs = iNumberOfTabs - 1
    End If
    PropertyChanged "NumberOfTabs"
    DrawTab
End Function

Public Sub GetPictureSize(pPicture As StdPicture, picX As Long, picY As Long)
    picX = 0
    picY = 0
    If pPicture Is Nothing Then Exit Sub
    picX = ScaleX(pPicture.Width, 8, 3)
    picY = ScaleY(pPicture.Height, 8, 3)
End Sub

Private Sub SetTabAccessKeys()
Dim i As Long
    sAccessKeys = ""
    For i = 1 To iNumberOfTabs
        sAccessKeys = sAccessKeys & SetAccessKey(Tabs(i).TabCaption)
    Next i
    UserControl.AccessKeys = sAccessKeys
End Sub

Public Function GetTabAccessKey(iKey As Integer) As Long
Dim i As Long
Dim sChr As String
    For i = 1 To iNumberOfTabs
        sChr = SetAccessKey(Tabs(i).TabCaption)
        If sChr <> "" Then
            If Asc(sChr) = iKey Then
                GetTabAccessKey = i
                Exit For
            End If
        End If
    Next i
End Function

Private Sub PlaceTabs()
Dim rc                      As RECT             'Temp RECT
Dim i                       As Long             'Loop and tab counter
Dim x                       As Single           'Tab Sizes used for creating a RECT.
Dim y                       As Single           '
Dim X1                      As Single           '
Dim Y1                      As Single           '
Dim iOldTabWidth            As Long             'Used for totaling up the tab widths or heights
Dim iTotalTabWidthHeight    As Long             'Total width or height of the tabs while counting
Dim ShowScroll              As Boolean          'Flag to show the scroll buttons
Dim iPicLeft                As Long             'Scroll box left
Dim iPicTop                 As Long             'Scroll box top
Dim iTabRight               As Long             'Tab right position
Dim iTabBottom              As Long             'Tab bottom position

If iFirstVisibleTab = 0 Then
    iFirstVisibleTab = 1
End If

If iLastVisibleTab = 0 Then
    iLastVisibleTab = 1
End If

'Calculate the total width or height of all the tabs.
'This will determine if we need to display
'the scroll buttons or not.
'** Scroll Box is actually "picBlank" (PictureBox Control)
    
    picBlank.Visible = False
    Select Case eTab
        Case 0, 3
        'Top and Bottom config.
            'Measure the length of all the tabs.
            ShowScroll = False
            'This allows for the little spaces between the tabs.
            iTotalTabWidthHeight = 2
            For i = 1 To iNumberOfTabs
                'If tab is set to AutoSize then calc the width
                'of the tab, otherwise just get the tab width.
                If Tabs(i).TabAutoSize = True Then
                    iOldTabWidth = CalculateTabWidth(i)
                Else
                    iOldTabWidth = Tabs(i).TabWidth
                End If
                '//
                'Add up what we have so far, not forgetting
                'to allow for the little spaces between the tabs.
                iTotalTabWidthHeight = iTotalTabWidthHeight + iOldTabWidth + 2
                If iTotalTabWidthHeight > UserControl.ScaleWidth Then
                    'The tab widths so far are greater than the
                    'control itself so we better show the scroll buttons.
                    ShowScroll = True
                    '//
                End If
                '//
            Next i
            If ShowScroll Then
            'Now measure what we can see, to find the last visible tab.
                'Clear these first
                iOldTabWidth = 0
                iTotalTabWidthHeight = 0
                '//
                'Show the scroll box
                DisplayScrollButtons
                '//
                For i = iFirstVisibleTab To iNumberOfTabs
                    'Yes we do this again to find the last visible tab.
                    If Tabs(i).TabAutoSize = True Then
                        iOldTabWidth = CalculateTabWidth(i)
                    Else
                        iOldTabWidth = Tabs(i).TabWidth
                    End If
                    iTotalTabWidthHeight = iTotalTabWidthHeight + iOldTabWidth + 2
                    '//
                    'iPicLeft is the left side of the scrollbox
                    iPicLeft = picBlank.Left
                    '//
                    'iTabRight is the right side of the currently
                    'counted tab.
                    iTabRight = iTotalTabWidthHeight
                    '//
                    'if iTabRight is greater than iPicLeft then
                    'This means that the tabs right side is either
                    'past the edge of the control or underneath
                    'the scroll box, so we call this one our iLastVisibleTab.
                    If iTabRight > iPicLeft Then
                        iLastVisibleTab = i
                        'Flag bRightIsDown to True so the user
                        'Can still scroll the buttons.
                        bRightIsDown = True
                        StopScrolling = False
                        '//
                        Exit For
                    Else
                        'If there is no tab under the scroll box
                        'then we must be at the end so set
                        'bRightIsDown to False so the user dosent keep scrolling.
                        'See the pRightDown_MouseDown sub to understand more.
                        bRightIsDown = False
                        StopScrolling = True
                        '//
                    End If
                Next i
            Else
                'No need for a scroll box, so set the
                'iLastVisibleTab to the greatest
                'number of tabs that we have.
                iLastVisibleTab = iNumberOfTabs
                iFirstVisibleTab = 1
                '//
            End If
            
        Case 1, 2
        'Left and Right config.
            iTotalTabWidthHeight = 2
            iMaxTabWidth = 0
            'We dont know it yet, but just incase we
            'need to display a scroll box we have to do this
            'loop to find the greatest width of the tabs,
            'so we know how wide to set the scroll box.
            For i = 1 To iNumberOfTabs
                If Tabs(i).TabAutoSize = True Then
                    iOldTabWidth = CalculateTabWidth(i)
                Else
                    iOldTabWidth = Tabs(i).TabWidth
                End If
                If iOldTabWidth >= iMaxTabWidth Then
                    iMaxTabWidth = iOldTabWidth
                End If
                'Add up the heights of the tabs so far
                'Allowing for the little space inbetween the tabs.
                iTotalTabWidthHeight = iTotalTabWidthHeight + iTabHeight + 1
                '//
            Next i
            '//
            If iTotalTabWidthHeight > UserControl.ScaleHeight Then
                'If the iTotalTabWidthHeight is greater than
                'the height of the control then flag ShowScroll to True.
                ShowScroll = True
                '//
            End If
            If ShowScroll Then
            'Show the scroll box
            DisplayScrollButtons
            '//
            'Now measure what we can see, to find the last visible tab.
                'Clear these first
                iOldTabWidth = 0
                iTotalTabWidthHeight = 0
                '//
                For i = iFirstVisibleTab To iNumberOfTabs
                    iTotalTabWidthHeight = iTotalTabWidthHeight + iTabHeight + 1
                    'Scroll box to edge position.
                    iPicTop = picBlank.Top
                    '//
                    'Current tab bottom edge position.
                    iTabBottom = iTotalTabWidthHeight
                    '//
                    If iTabBottom > iPicTop Then
                        'If we are here then the bottom edge of the tab
                        'is under the Scroll Box, so this is our
                        'Last tab
                        iLastVisibleTab = i
                        '//
                        'Flag bRightIsDown to True so the user
                        'can still scroll the tabs.
                        bRightIsDown = True
                        StopScrolling = False
                        '//
                        Exit For
                    Else
                        'We must be at the end of the tab strip
                        'So dont allow the user to keep on scrolling.
                        bRightIsDown = False
                        StopScrolling = True
                        '//
                    End If
                Next i
            Else
                'No need for a scroll box so set our last tab
                iLastVisibleTab = iNumberOfTabs
                '//
            End If
    End Select
'//
    
    'Setup the body of the Tab Control
    rcBody = GetRect(UserControl.hwnd)
    Select Case eTab
        Case 0
            'Top config.
            rcBody.Top = rcBody.Top + iTabHeight
        Case 1
            'Left config.
            For i = 1 To iNumberOfTabs
                If Tabs(i).TabAutoSize = True Then
                    iOldTabWidth = CalculateTabWidth(i)
                Else
                    iOldTabWidth = Tabs(i).TabWidth
                End If
                If iOldTabWidth >= iMaxTabWidth Then
                    iMaxTabWidth = iOldTabWidth
                End If
            Next i
            If iMaxTabWidth < 20 Then
                iMaxTabWidth = 20
            End If
            rcBody.Left = rcBody.Left + iMaxTabWidth
        Case 2
            'Right
            For i = 1 To iNumberOfTabs
                If Tabs(i).TabAutoSize = True Then
                    iOldTabWidth = CalculateTabWidth(i)
                Else
                    iOldTabWidth = Tabs(i).TabWidth
                End If
                If iOldTabWidth >= iMaxTabWidth Then
                    iMaxTabWidth = iOldTabWidth
                End If
            Next i
            If iMaxTabWidth < 20 Then
                iMaxTabWidth = 20
            End If
            rcBody.Right = rcBody.Right - iMaxTabWidth
        Case 3
            'Bottom
            rcBody.Bottom = rcBody.Bottom - iTabHeight
    End Select
    '//
    
    'Clear the control and draw a square
    'for the body of the tab control.
    Cls
    DrawASquare hdc, rcBody, oFrameColor
    '//
    
    If Not ShowScroll Then
        iFirstVisibleTab = 1
    End If
    
    'Ok now lets draw some Tabs.
    'Create the RECT areas for the tabs.
    Select Case eTab
        Case 0
        'Tabs At The Top
            x = 2
            y = 2
            Y1 = TabHeight
            For i = iFirstVisibleTab To iLastVisibleTab
                With Tabs(i)
                'Position the Tabs.
                    .TabLeft = x
                    .TabTop = y
                    If .TabAutoSize = False Then
                        X1 = .TabWidth
                    Else
                        'Have to figure out the width before drawing tab
                        X1 = CalculateTabWidth(i)
                    End If

                    .TabHeight = Y1
                '//
                
                    'Create a RECT area using the above dimentions to draw the tab.
                    With rc
                        .Left = x
                        .Top = y
                        .Right = .Left + X1
                        .Bottom = Y1
                    '//
                    
                    'Keep this RECT
                        ReDim Preserve rcTabs(i)
                        CopyTheRect rcTabs(i), rc
                    '//
                    
                    'Draw the tabs
                        If i <> iSelectedTab Then
                            ShadeTab i
                            DrawUnSelectedTab i
                        Else
                            DrawSelectedTab i
                        End If
                    '//
                    End With
                'Move across to the next tab
                    x = (x + 2) + X1
                '//

                End With
            Next i
        Case 1
        'Tabs On The Left Hand Side
            x = 2
            y = 2
            For i = iFirstVisibleTab To iLastVisibleTab
                With Tabs(i)
                
                'Position the Tabs.
                    .TabLeft = x
                    .TabTop = y
                    X1 = iMaxTabWidth
                    Y1 = y + TabHeight
                '//
                
                    'Create a RECT area using the above dimentions to draw the tab.
                    With rc
                        .Left = x
                        .Top = y
                        .Right = (.Left + X1) - 2
                        .Bottom = Y1
                    '//
                
                    'Keep this RECT
                        ReDim Preserve rcTabs(i)
                        CopyTheRect rcTabs(i), rc
                    '//
                    
                    'Draw the tabs
                        If i <> iSelectedTab Then
                            ShadeTab i
                            DrawUnSelectedTab i
                        Else
                            DrawSelectedTab i
                        End If
                    '//
                    End With
                    
                End With
                'Move down to the next tab
                    y = (y + 1) + TabHeight
                '//
            Next i
        Case 2
        'Tabs On The Right Hand Side
            x = rcBody.Right
            y = 2
            For i = iFirstVisibleTab To iLastVisibleTab
                With Tabs(i)
                'Position the Tabs.
                    .TabLeft = x
                    .TabTop = y
                    X1 = iMaxTabWidth
                    Y1 = y + TabHeight
                '//
                    'Create a RECT area using the above dimentions to draw into.
                    With rc
                        .Left = x
                        .Top = y
                        .Right = (.Left + X1) - 2
                        .Bottom = Y1
                    '//
                
                    'Keep this RECT
                        ReDim Preserve rcTabs(i)
                        CopyTheRect rcTabs(i), rc
                    '//
                    
                    'Draw the tabs
                        If i <> iSelectedTab Then
                            ShadeTab i
                            DrawUnSelectedTab i
                        Else
                            DrawSelectedTab i
                        End If
                    '//
                    End With
                    'Move down to the next tab
                    y = (y + 1) + TabHeight
                    '//
                    
                End With
            Next i
        Case 3
        'Tabs At The Bottom
            x = 2
            y = rcBody.Bottom - 2
            Y1 = TabHeight + y
            For i = iFirstVisibleTab To iLastVisibleTab
                With Tabs(i)
                
                'Position the Tabs.
                    .TabLeft = x
                    .TabTop = y
                    If .TabAutoSize = False Then
                        X1 = .TabWidth
                    Else
                        'Have to figure out the width before drawing tab
                        X1 = CalculateTabWidth(i)
                    End If
                    .TabHeight = Y1
                '//
                
                'Create a RECT area using the above dimentions to draw into.
                    With rc
                        .Left = x
                        .Top = y
                        .Right = .Left + X1
                        .Bottom = Y1
                    'Keep this RECT
                        ReDim Preserve rcTabs(i)
                        CopyTheRect rcTabs(i), rc
                    '//
                    
                    'Draw the tabs
                        If i <> iSelectedTab Then
                            ShadeTab i
                            DrawUnSelectedTab i
                        Else
                            DrawSelectedTab i
                        End If
                    '//
                        
                    End With
                    'Move across to the next tab.
                    x = (x + 2) + X1
                    '//
                    
                End With
            Next i
    End Select
    
    'Draw the pictures and captions.
    For i = iFirstVisibleTab To iLastVisibleTab
        DrawPictureAndCaption i
    Next i
    '//
    
    Refresh
End Sub

Public Sub DrawUnSelectedTab(Index As Long)
'Here we draw the unselected tabs.
    Select Case eTab
        Case 0
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left, .Top + 2, .Left, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left + 2, .Top, .Right - 1, .Top, oFrameColor
                'Right
                DrawALine UserControl.hdc, .Right, .Top + 2, .Right, .Bottom, oFrameColor
                'Left Corner
                DrawADot UserControl.hdc, .Left + 1, .Top + 1, oFrameColor
                'Right Corner
                DrawADot UserControl.hdc, .Right - 1, .Top + 1, oFrameColor
            
            'Highlights HotTab
                If iHotTab = Index And MouseInTab = True Then
                    DrawALine UserControl.hdc, .Left + 2, .Top, .Right - 1, .Top, oTabHotStripColor
                    DrawALine UserControl.hdc, .Left + 2, .Top + 1, .Right - 1, .Top + 1, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                    DrawALine UserControl.hdc, .Left + 1, .Top + 2, .Right, .Top + 2, AdjustToOLE_COLOR(oTabHotStripColor, 40)
                End If
            End With
            '//
        Case 1
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left, .Top + 2, .Left, .Bottom - 2, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left + 2, .Top, .Right, .Top, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left + 2, .Bottom - 1, .Right, .Bottom - 1, oFrameColor
                'Left Top Corner
                DrawADot UserControl.hdc, .Left + 1, .Top + 1, oFrameColor
                'Bottom Left Corner
                DrawADot UserControl.hdc, .Left + 1, .Bottom - 2, oFrameColor
            
            'Highlights HotTab
                If iHotTab = Index And MouseInTab = True Then
                    DrawALine UserControl.hdc, .Left, .Top + 2, .Left, .Bottom - 2, oTabHotStripColor
                    DrawALine UserControl.hdc, .Left + 1, .Top + 2, .Left + 1, .Bottom - 2, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                    DrawALine UserControl.hdc, .Left + 2, .Top + 1, .Left + 2, .Bottom - 1, AdjustToOLE_COLOR(oTabHotStripColor, 40)
                End If
            '//
            End With
        Case 2
            With rcTabs(Index)
                'Right
                DrawALine UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 2, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left, .Top, .Right - 2, .Top, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left, .Bottom - 1, .Right - 2, .Bottom - 1, oFrameColor
                'Left Top Corner
                DrawADot UserControl.hdc, .Right - 2, .Top + 1, oFrameColor
                'Bottom Left Corner
                DrawADot UserControl.hdc, .Right - 2, .Bottom - 2, oFrameColor
            
            'Highlights HotTab
                If iHotTab = Index And MouseInTab = True Then
                    DrawALine UserControl.hdc, .Right - 1, .Top + 2, .Right - 1, .Bottom - 2, oTabHotStripColor
                    DrawALine UserControl.hdc, .Right - 2, .Top + 2, .Right - 2, .Bottom - 2, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                    DrawALine UserControl.hdc, .Right - 3, .Top + 1, .Right - 3, .Bottom - 1, AdjustToOLE_COLOR(oTabHotStripColor, 40)
                End If
            '//
            
            End With
        Case 3
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left, .Top + 2, .Left, .Bottom - 2, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left + 2, .Bottom - 1, .Right - 1, .Bottom - 1, oFrameColor
                'Right
                DrawALine UserControl.hdc, .Right, .Top + 2, .Right, .Bottom - 2, oFrameColor
                'Left Corner
                DrawADot UserControl.hdc, .Left + 1, .Bottom - 2, oFrameColor
                'Right Corner
                DrawADot UserControl.hdc, .Right - 1, .Bottom - 2, oFrameColor
            
            'Highlights HotTab
                If iHotTab = Index And MouseInTab = True Then
                    DrawALine UserControl.hdc, .Left + 2, .Bottom - 1, .Right - 1, .Bottom - 1, oTabHotStripColor
                    DrawALine UserControl.hdc, .Left + 2, .Bottom - 2, .Right - 1, .Bottom - 2, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                    DrawALine UserControl.hdc, .Left + 1, .Bottom - 3, .Right, .Bottom - 3, AdjustToOLE_COLOR(oTabHotStripColor, 40)
                End If
            '//
            
            End With
    End Select
End Sub

Public Sub DrawSelectedTab(Index As Long)
'Here we draw the selected tab
    Select Case eTab
        Case 0
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left - 2, .Top, .Left - 2, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left, .Top - 2, .Right + 1, .Top - 2, oFrameColor
                'Right
                DrawALine UserControl.hdc, .Right + 2, .Top, .Right + 2, .Bottom, oFrameColor
                'Left Corner
                DrawADot UserControl.hdc, .Left - 1, .Top - 1, oFrameColor
                'Right Corner
                DrawADot UserControl.hdc, .Right + 1, .Top - 1, oFrameColor
                'Blank line (rem it out to see why i done it.)
                DrawALine UserControl.hdc, .Left - 1, .Bottom, .Right + 2, .Bottom, oBackColor
            'End With
            
            'Highlights
            If hasFocus = True Then
                DrawALine UserControl.hdc, .Left, .Top - 2, .Right + 1, .Top - 2, oTabHotStripColor
                DrawALine UserControl.hdc, .Left, .Top - 1, .Right + 1, .Top - 1, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                DrawALine UserControl.hdc, .Left - 1, .Top, .Right + 2, .Top, AdjustToOLE_COLOR(oTabHotStripColor, 40)
            End If
            
            End With
            '//
        Case 1
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left - 2, .Top, .Left - 2, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left, .Top - 2, .Right, .Top - 2, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left, .Bottom + 1, .Right, .Bottom + 1, oFrameColor
                'Left Top Corner
                DrawADot UserControl.hdc, .Left - 1, .Top - 1, oFrameColor
                'Bottom Left Corner
                DrawADot UserControl.hdc, .Left - 1, .Bottom, oFrameColor
                'Blank line
                DrawALine UserControl.hdc, .Right, .Top - 1, .Right, .Bottom + 1, oBackColor
            
            'Highlights
            If hasFocus = True Then
                DrawALine UserControl.hdc, .Left - 2, .Top, .Left - 2, .Bottom, oTabHotStripColor
                DrawALine UserControl.hdc, .Left - 1, .Top, .Left - 1, .Bottom, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                DrawALine UserControl.hdc, .Left, .Top - 1, .Left, .Bottom + 1, AdjustToOLE_COLOR(oTabHotStripColor, 40)
            End If
            
            End With
        Case 2
            With rcTabs(Index)
                'Right
                DrawALine UserControl.hdc, .Right + 1, .Top, .Right + 1, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hdc, .Left, .Top - 2, .Right, .Top - 2, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left, .Bottom + 1, .Right, .Bottom + 1, oFrameColor
                'Left Top Corner
                DrawADot UserControl.hdc, .Right, .Top - 1, oFrameColor
                'Bottom Left Corner
                DrawADot UserControl.hdc, .Right, .Bottom, oFrameColor
                'Balnk line
                DrawALine UserControl.hdc, .Left - 1, .Top - 1, .Left - 1, .Bottom + 1, oBackColor
            
            'Highlights
            If hasFocus = True Then
                DrawALine UserControl.hdc, .Right + 1, .Top, .Right + 1, .Bottom, oTabHotStripColor
                DrawALine UserControl.hdc, .Right, .Top, .Right, .Bottom, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                DrawALine UserControl.hdc, .Right - 1, .Top - 1, .Right - 1, .Bottom + 1, AdjustToOLE_COLOR(oTabHotStripColor, 40)
            End If
            
            End With
        Case 3
            With rcTabs(Index)
                'Left
                DrawALine UserControl.hdc, .Left - 2, .Top + 2, .Left - 2, .Bottom, oFrameColor
                'Bottom
                DrawALine UserControl.hdc, .Left, .Bottom + 1, .Right + 1, .Bottom + 1, oFrameColor
                'Right
                DrawALine UserControl.hdc, .Right + 2, .Top + 2, .Right + 2, .Bottom, oFrameColor
                'Left Corner
                DrawADot UserControl.hdc, .Left - 1, .Bottom, oFrameColor
                'Right Corner
                DrawADot UserControl.hdc, .Right + 1, .Bottom, oFrameColor
                'Blank line
                DrawALine UserControl.hdc, .Left - 1, .Top + 1, .Right + 2, .Top + 1, oBackColor
            
            'Highlights
            If hasFocus = True Then
                DrawALine UserControl.hdc, .Left, .Bottom + 1, .Right + 1, .Bottom + 1, oTabHotStripColor
                DrawALine UserControl.hdc, .Left, .Bottom, .Right + 1, .Bottom, AdjustToOLE_COLOR(oTabHotStripColor, 20)
                DrawALine UserControl.hdc, .Left - 1, .Bottom - 1, .Right + 2, .Bottom - 1, AdjustToOLE_COLOR(oTabHotStripColor, 40)
            End If
            
            End With
    End Select
End Sub

Private Function CalculateTabWidth(Index As Long) As Long
Dim rc              As RECT
Dim iTextWidth      As Long
Dim pY              As Long
'Use this when the AutoSize property is set to True.

    With Tabs(Index)
        'Use GetTextRect to create a RECT the same size
        'as the Tab Caption.
        GetTextRect UserControl.hdc, Tabs(Index).TabCaption, Len(Tabs(Index).TabCaption), rc
        '//
        'Calculate the width of the text.
        iTextWidth = rc.Left + rc.Right
        '//
        'If it has one get the width of the picture.
        GetPictureSize Tabs(Index).TabIcon, 0, pY
        '//
    End With
    'Now add the length of the text,the width
    'of the picture and add a little bit
    'and this will be the width of our tab.
    CalculateTabWidth = iTextWidth + pY + 10
    '//
End Function

Private Sub DisplayScrollButtons(Optional JustDrawButtons As Boolean = False)
Dim rc                  As RECT
Dim iSquare             As Long
Dim iBttnCenter         As Long
Dim i                   As Long
Dim j                   As Long
    picBlank.BackColor = oBackColor
    Select Case eTab
        Case 0
        'Top
            'iSquare will always be the height of
            'the tabs -4. This is so the scroll buttons
            'will be in proportion to the height of the tabs.
            iSquare = iTabHeight - 4
            '//
            'Position and size the picBlank and scroll buttons.
            With picBlank
                .Left = (UserControl.ScaleWidth - iSquare * 2) - 4
                .Top = 0
                .Height = iTabHeight
                .Width = iSquare * 2 + 4
                .Visible = True
            End With
            
            'The scroll buttons
            With pLeftUp
                .Height = iSquare
                .Width = iSquare
                .Left = 2
                .Top = (picBlank.Height - iSquare) / 2
            End With
            With pRightDown
                .Height = iSquare
                .Width = iSquare
                .Left = picBlank.ScaleWidth - iSquare
                .Top = (picBlank.Height - iSquare) / 2
            End With
            '//
            '//
        Case 1
        'Left
            'We use the variable iMaxTabWidth here to
            'size the width of the scroll buttons and
            'picBlank so its in proportion to the width
            'of the tabs.
            'But the height is always at a constant 46
            'from the bottom of the tab control.
            With picBlank
                .Left = 0
                .Top = UserControl.ScaleHeight - 46
                .Height = 46
                .Width = iMaxTabWidth
                .Visible = True
            End With
            
            With pLeftUp
                .Height = 20
                .Width = iMaxTabWidth - 4
                .Left = 2
                .Top = 4
            End With
            With pRightDown
                .Height = 20
                .Width = iMaxTabWidth - 4
                .Left = 2
                .Top = picBlank.Height - 20
            End With
            '//
        Case 2
        'Right
            'Works in reverse to Case 1:
            With picBlank
                .Left = UserControl.ScaleWidth - iMaxTabWidth
                .Top = UserControl.ScaleHeight - 46
                .Height = 46
                .Width = UserControl.ScaleWidth
                .Visible = True
            End With
            
            With pLeftUp
                .Height = 20
                .Width = iMaxTabWidth - 4
                .Left = 2
                .Top = 4
            End With
            With pRightDown
                .Height = 20
                .Width = iMaxTabWidth - 4
                .Left = 2
                .Top = picBlank.Height - 20
            End With
        Case 3
            'Works in reverse to Case 0:
            iSquare = iTabHeight - 4
            With picBlank
                .Left = (UserControl.ScaleWidth - iSquare * 2) - 4
                .Top = UserControl.ScaleHeight - iTabHeight
                .Height = iTabHeight
                .Width = iSquare * 2 + 4
                .Visible = True
            End With
            
            With pLeftUp
                .Height = iSquare
                .Width = iSquare
                .Left = 2
                .Top = (picBlank.Height - iSquare) / 2
            End With
            With pRightDown
                .Height = iSquare
                .Width = iSquare
                .Left = picBlank.ScaleWidth - iSquare
                .Top = (picBlank.Height - iSquare) / 2
            End With
    End Select
    
    'Draw the scroll buttons
    
    Select Case eTab
        Case 0, 3
        'Top and Bottom tab config.
        'Find the center of the scroll button
        'ready to draw the arrows into.
        iBttnCenter = iSquare / 2
        '//
            With pLeftUp
                'Shade the buttons and draw a hot border if required.
                ClearRect rc
                .Cls
                rc = GetRect(.hwnd)
                DrawGradient .hdc, AdjustToOLE_COLOR(oScrollBackColor, 30), AdjustToOLE_COLOR(oScrollBackColor, -30), rc, .ScaleHeight
                rc = GetRect(.hwnd)
                DrawASquare .hdc, rc, oFrameColor
                If mMouseOverLScroll = True Then
                    ResizeRect rc, -1, -1
                    DrawASquare .hdc, rc, oTabHotStripColor
                End If
                '//
                'Draw The Arrows
                j = 3
                For i = 1 To 3
                    DrawALine .hdc, iBttnCenter - j, iBttnCenter, iBttnCenter + i, iBttnCenter - 4, oScrollColor
                    DrawALine .hdc, iBttnCenter - j, iBttnCenter, iBttnCenter + i, iBttnCenter + 4, oScrollColor
                    j = j - 1
                Next i
            
                DrawADot .hdc, iBttnCenter + 1, iBttnCenter - 4, oScrollColor
                DrawADot .hdc, iBttnCenter + 1, iBttnCenter + 4, oScrollColor
                '//
                
                .Visible = True
            End With
            With pRightDown
                'Shade the buttons and draw a hot border if required.
                ClearRect rc
                .Cls
                rc = GetRect(.hwnd)
                DrawGradient .hdc, AdjustToOLE_COLOR(oScrollBackColor, 30), AdjustToOLE_COLOR(oScrollBackColor, -30), rc, .ScaleHeight
                rc = GetRect(.hwnd)
                DrawASquare .hdc, rc, oFrameColor
                If mMouseOverRScroll = True Then
                    ResizeRect rc, -1, -1
                    DrawASquare .hdc, rc, oTabHotStripColor
                End If
                '//
                'Draw The Arrows
                j = 3
                For i = 1 To 3
                    DrawALine .hdc, iBttnCenter + j, iBttnCenter, iBttnCenter - i, iBttnCenter - 4, oScrollColor
                    DrawALine .hdc, iBttnCenter + j, iBttnCenter, iBttnCenter - i, iBttnCenter + 4, oScrollColor
                    j = j - 1
                Next i
            
                DrawADot .hdc, iBttnCenter - 1, iBttnCenter - 4, oScrollColor
                DrawADot .hdc, iBttnCenter - 1, iBttnCenter + 4, oScrollColor
                '//
                
                .Visible = True
            End With
        Case 1, 2
        iBttnCenter = (iMaxTabWidth / 2) - 2
        'Left and Right
            With pLeftUp
                'Shade the buttons and draw a hot border if required.
                ClearRect rc
                .Cls
                rc = GetRect(.hwnd)
                DrawGradient .hdc, AdjustToOLE_COLOR(oScrollBackColor, 30), AdjustToOLE_COLOR(oScrollBackColor, -30), rc, .ScaleHeight
                rc = GetRect(.hwnd)
                DrawASquare .hdc, rc, oFrameColor
                If mMouseOverLScroll = True Then
                    ResizeRect rc, -1, -1
                    DrawASquare .hdc, rc, oTabHotStripColor
                End If
                '//
                
                'Draw The Arrows
                For i = 0 To 2
                    DrawALine .hdc, iBttnCenter - 4, 9 + i, iBttnCenter, 5 + i, oScrollColor
                Next i
                
                For i = 0 To 2
                    DrawALine .hdc, iBttnCenter - 1, 6 + i, iBttnCenter + 3, 10 + i, oScrollColor
                Next i

                DrawADot .hdc, iBttnCenter - 5, 10, oScrollColor
                DrawADot .hdc, iBttnCenter + 3, 10, oScrollColor
                '//
                
                .Visible = True
            End With
            With pRightDown
                'Shade the buttons and draw a hot border if required.
                ClearRect rc
                rc = GetRect(.hwnd)
                .Cls
                rc = GetRect(.hwnd)
                DrawGradient .hdc, AdjustToOLE_COLOR(oScrollBackColor, 30), AdjustToOLE_COLOR(oScrollBackColor, -30), rc, .ScaleHeight
                rc = GetRect(.hwnd)
                DrawASquare .hdc, rc, oFrameColor
                If mMouseOverRScroll = True Then
                    ResizeRect rc, -1, -1
                    DrawASquare .hdc, rc, oTabHotStripColor
                End If
                '//
                
                'Draw The Arrows
                For i = 0 To 2
                    DrawALine .hdc, iBttnCenter - 4, 7 + i, iBttnCenter, 11 + i, oScrollColor
                Next i
                
                For i = 0 To 2
                    DrawALine .hdc, iBttnCenter - 1, 10 + i, iBttnCenter + 3, 6 + i, oScrollColor
                Next i

                DrawADot .hdc, iBttnCenter - 5, 8, oScrollColor
                DrawADot .hdc, iBttnCenter + 3, 8, oScrollColor
                '//
                
                .Visible = True
            End With
    End Select
End Sub

Private Sub ShadeTab(Index As Long)
Dim rc As RECT
Dim i As Long
Dim iHeight As Long
    CopyTheRect rc, rcTabs(Index)
    If eTab <> tRight Then
        rc.Left = rc.Left + 1
    End If
    
    rc.Top = rc.Top
    If eTab = tBottom Then
        iHeight = rc.Bottom - rc.Top - 3
    Else
        iHeight = rc.Bottom - rc.Top - 2
    End If
    DrawGradient UserControl.hdc, AdjustToOLE_COLOR(oBackColor, 20), AdjustToOLE_COLOR(oBackColor, -20), rc, iHeight
End Sub

Private Sub DrawPictureAndCaption(Index As Long)
Dim rc                  As RECT
Dim rcText              As RECT
Dim pX                  As Long
Dim pY                  As Long
Dim wTabText            As Long
Dim hTabText            As Long
Dim WidthOfTab          As Long
Dim HeightOfTab         As Long
Dim HalfTabHeight       As Long
Dim HalfTabTextHeight   As Long
Dim HalfPicHeight       As Long
Dim iDrawTextFlag       As Long
Dim iNewTabWidth        As Long
Dim iTextLeft           As Long
Dim iTextTop            As Long
Dim oTextColor          As OLE_COLOR
    'Make a copy of the tab RECT
    'so we dont alter the original tab RECT.
    CopyTheRect rc, rcTabs(Index)
    '//
    
    'Make a copy of the tab RECT to draw some text into.
    CopyTheRect rcText, rcTabs(Index)
    '//
    
    'Get the size of the picture even if there is not one set.
    GetPictureSize Tabs(Index).TabIcon, pX, pY
    If pY > 1 Then
        HalfPicHeight = pY / 2
    End If
    '//
    
    'Resize the rcText rect to be the size of the captopn text
    'using the DT_CALCRECT constant.
    'NB, this does not draw the text.
    GetTextRect UserControl.hdc, Tabs(Index).TabCaption, Len(Tabs(Index).TabCaption), rcText
    '//
    
    'Store the Width and Height of the newly resized rcText
    'in the variables wTabText and hTabText.
    With rcText
        wTabText = .Right - .Left
        hTabText = .Bottom - .Top
    End With
    '//
    
    'Store the Width, Height and Half Height of the Tab
    'in the variables WidthOfTab and HeightOfTab.
    WidthOfTab = rc.Right - rc.Left
    HeightOfTab = rc.Bottom - rc.Top
    HalfTabHeight = HeightOfTab / 2
    '//
    
    'Store the half text height, in the variable HalfTabTextHeight.
    HalfTabTextHeight = (rcText.Bottom - rcText.Top) / 2
    '//

    'If the user has turned off the AutoSize property
    'check to see if the picture and caption can fit into the tab.
    If wTabText + pX + 4 > WidthOfTab Then
    'Use Ellipsis (...)
    'The Re-Calculate the rcText RECT.
        iDrawTextFlag = DT_WORD_ELLIPSIS
        iNewTabWidth = (wTabText + pX + 4) - WidthOfTab
        rcText.Right = rcText.Right - iNewTabWidth
    Else
    'Fits in.
        iDrawTextFlag = DT_CENTER
    End If
    
    'Setup variables to Position the rcText
    'left and top positions
    'to match the Alignment property.
    Select Case eTab
        Case 0
        'Top
            If Index = iSelectedTab Then
                iTextLeft = 4
                iTextTop = 1
            Else
                iTextLeft = 4
                iTextTop = 0
            End If
        Case 1
        'Left
            If Index = iSelectedTab Then
                iTextLeft = 2
                iTextTop = 0
            Else
                iTextLeft = 4
                iTextTop = 0
            End If
        Case 2
        'Right
            If Index = iSelectedTab Then
                iTextLeft = 4
                iTextTop = 0
            Else
                iTextLeft = 2
                iTextTop = 0
            End If
        Case 3
        'Bottom
            If Index = iSelectedTab Then
                iTextLeft = 4
                iTextTop = -2
            Else
                iTextLeft = 4
                iTextTop = 0
            End If
    End Select
    '//
    
    'Select the text color
    If Tabs(Index).TabEn = False Then
        'Tab Disabled.
        oTextColor = oForeColorDisabled
    Else
    If Index = iHotTab Then
        'Mouse over tab.
        oTextColor = oForeColorHot
    Else
    If Index = iSelectedTab Then
        'Current selected tab
        oTextColor = oActiveForeColor
    Else
        'Enabled tab, no mouse over non selected tab.
        oTextColor = oForeColor
    End If
    End If
    End If
    '//
    
    'Position and draw the text.
    SetTheTextColor UserControl.hdc, oTextColor
    PositionRect rcText, iTextLeft + pX, HalfTabHeight - HalfTabTextHeight - iTextTop
    DrawTheText UserControl.hdc, Tabs(Index).TabCaption, Len(Tabs(Index).TabCaption), rcText, iDrawTextFlag
    '//
    
    'Draw the pictures.
    If Not Tabs(Index).TabIcon Is Nothing Then
        'If Tabs(Index).TabIcon.Type = 1 Then
            Set cPic = New cImageManipulation
            cPic.PaintTransparentPicture UserControl.hdc, Tabs(Index).TabIcon, rcTabs(Index).Left + iTextLeft - 2, rcTabs(Index).Top + HalfTabHeight - HalfPicHeight - iTextTop, pX, pY, 0, 0, oMaskColor
        'Else
        'If Tabs(Index).TabIcon.Type = 3 Then
            
            'UserControl.PaintPicture Tabs(Index).TabIcon, rcTabs(Index).Left + iTextLeft, rcTabs(Index).Top + HalfTabHeight - HalfPicHeight - iTextTop, pX, pY, 0, 0, pX, pY
        'End If
        'End If
        Set cPic = Nothing
    End If
    '//
End Sub



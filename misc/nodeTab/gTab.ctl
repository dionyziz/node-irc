VERSION 5.00
Begin VB.UserControl gnTab 
   Alignable       =   -1  'True
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   ScaleHeight     =   780
   ScaleWidth      =   4905
   Begin VB.CommandButton cmdScroll 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton cmdScroll 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pbTop 
      Align           =   1  'Align Top
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   4905
      TabIndex        =   4
      Top             =   0
      Width           =   4905
   End
   Begin VB.PictureBox pbBottom 
      Align           =   2  'Align Bottom
      Height          =   50
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   4905
      TabIndex        =   3
      Top             =   735
      Width           =   4905
   End
   Begin VB.PictureBox pbTabs 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1335
   End
   Begin VB.Timer tmrMouseOut 
      Left            =   0
      Top             =   240
   End
End
Attribute VB_Name = "gnTab"
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

'-- Error constants
Private Const TAB_EXISTS = 8601
Private Const TAB_DOESNT_EXIST = 8602

' Border Constraints
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BDR_RAISED = &H5
Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_THIN = BDR_RAISEDINNER

Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_SOFT = &H1000
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private mBorderType As Long

'-- Constants for API calls
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_NOPREFIX = &H800
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const WINDING = 2
Private Const PS_SOLID = 0
Private Const m_def_HoverColour = &H80000012
Private Const m_def_SelectedTabHoverColor = 0
Private Const m_def_SeperatorLineColor = &H80000012
Private Const m_def_AllignBottom = 0
Private Const m_def_HoverLineColour = 0
Private Const m_def_BorderType = 1

'-- Types for API calls
Private Type RECT
    left    As Long
    top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    X       As Long
    Y       As Long
End Type

'-- Type to store information for each tab
Private Type gnTab
    Key             As String
    Caption         As String
    ToolTipText     As String
    left            As Long
    Width           As Long
    Active          As Boolean
    Hovering        As Boolean
End Type

Public Enum enmBorderType
    [enmNone] = 0
    [enmRaised] = 1
    [enmSunken] = 2
    [enmBump] = 3
    [enmEtched] = 4
    [enmThin] = 5
End Enum


'-- API Declares
Private Declare Function SetBkColor Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal crColor As Long) As Long
    
Private Declare Function SetTextColor Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal crColor As Long) As Long
    
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
    ByVal hDC As Long, _
    ByVal lpStr As String, _
    ByVal nCount As Long, _
    lpRect As RECT, _
    ByVal wFormat As Long) As Long
    
Private Declare Function OleTranslateColor Lib "olepro32.dll" ( _
    ByVal Color As Long, _
    ByVal hPal As Long, _
    ByRef pClrRef As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" _
    (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function GetCursorPos Lib _
"user32" (lpPoint As POINTAPI) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
    ByVal hObject As Long) As Long
    
'-- Events
Public Event Change(Index As Long, TabCaption As String, TabKey As String)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove

'-- Private variables
Private mTabBackColor       As OLE_COLOR
Private mTabTextColor       As OLE_COLOR
Private mTabLineColor       As OLE_COLOR
Private mActiveTabBackColor As OLE_COLOR
Private mActiveTabTextColor As OLE_COLOR
Private mActiveTabLineColor As OLE_COLOR
Private mBackColor          As OLE_COLOR
Private m_HoverColour       As OLE_COLOR
Private m_SeperatorLineColor As OLE_COLOR
Private m_SelectedTabHoverColor As OLE_COLOR

Private m_BorderType As enmBorderType

Private m_AllignBottom      As Boolean
Private mFlip               As Boolean
Private mbNeedButtons       As Boolean
Private mLeftMostTab        As Long
Private mTabs()             As gnTab
Private mTabCount           As Long
Private mActiveTab          As Long





'-- Exposed properties

Public Property Set Font(NewFont As StdFont)
    Set pbTabs.Font = NewFont
    PropertyChanged "Font"
    UserControl_Resize
    CalcTabPos
    DrawTabs
End Property

Public Property Get Font() As StdFont
    Set Font = pbTabs.Font
End Property

Public Property Let TabBackColor(val As OLE_COLOR)
    mTabBackColor = val
    PropertyChanged "TabBackColor"
    DrawTabs
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get TabBackColor() As OLE_COLOR
    TabBackColor = mTabBackColor
End Property

Public Property Let TabTextColor(val As OLE_COLOR)
    mTabTextColor = val
    PropertyChanged "TabTextColor"
    DrawTabs
End Property

Public Property Get TabTextColor() As OLE_COLOR
    TabTextColor = mTabTextColor
End Property


Public Property Let ActiveTabBackColor(val As OLE_COLOR)
    mActiveTabBackColor = val
    PropertyChanged "ActiveTabBackColor"
    DrawTabs
End Property

Public Property Get ActiveTabBackColor() As OLE_COLOR
    ActiveTabBackColor = mActiveTabBackColor
End Property

Public Property Let ActiveTabTextColor(val As OLE_COLOR)
    mActiveTabTextColor = val
    PropertyChanged "ActiveTabTextColor"
    DrawTabs
End Property

Public Property Get ActiveTabTextColor() As OLE_COLOR
    ActiveTabTextColor = mActiveTabTextColor
End Property


Public Property Let BackColor(val As OLE_COLOR)
    mBackColor = val
    PropertyChanged "BackColor"
    DrawTabs
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Public Property Let SelectedTab(vData As String)
    Dim i As Long
    Dim r As Long
    
    r = TabFromKey(vData)
    
    If r = 0 Then ErrRaise TAB_DOESNT_EXIST
    mActiveTab = r
    
    For i = 1 To mTabCount
        mTabs(i).Active = (i = r)
    Next
    
    DrawTabs
    
    Debug.Print "ChangeTab " & mTabs(r).Key
    
    RaiseEvent Change(r, mTabs(r).Caption, mTabs(r).Key)

End Property

Public Property Get SelectedTab() As String
    
    SelectedTab = mTabs(mActiveTab).Key
    
End Property

Public Property Set MouseIcon(ByVal vData As StdPicture)
    pbTabs.MouseIcon = vData

End Property

Public Property Let MouseIcon(ByVal vData As StdPicture)
    pbTabs.MouseIcon = vData

End Property

Public Property Let MousePointer(ByVal vData As Long)
    
    pbTabs.MousePointer = vData

End Property

Public Property Get HoverColour() As OLE_COLOR
    HoverColour = m_HoverColour
End Property

Public Property Let HoverColour(ByVal New_HoverColour As OLE_COLOR)
    m_HoverColour = New_HoverColour
    PropertyChanged "HoverColour"
    DrawTabs
End Property


Public Property Get AllignBottom() As Boolean
    AllignBottom = m_AllignBottom
End Property

Public Property Let AllignBottom(ByVal New_AllignBottom As Boolean)
    m_AllignBottom = New_AllignBottom
    PropertyChanged "AllignBottom"
    DrawTabs
End Property


Public Property Get SeperatorLineColor() As OLE_COLOR
    SeperatorLineColor = m_SeperatorLineColor
End Property

Public Property Let SeperatorLineColor(ByVal New_SeperatorLineColor As OLE_COLOR)
    m_SeperatorLineColor = New_SeperatorLineColor
    PropertyChanged "SeperatorLineColor"
    DrawTabs
End Property

Public Property Get SelectedTabHoverColor() As OLE_COLOR
    SelectedTabHoverColor = m_SelectedTabHoverColor
End Property

Public Property Let SelectedTabHoverColor(ByVal New_SelectedTabHoverColor As OLE_COLOR)
    m_SelectedTabHoverColor = New_SelectedTabHoverColor
    PropertyChanged "SelectedTabHoverColor"
End Property

Public Property Get TabBorderType() As enmBorderType
    
    TabBorderType = m_BorderType
   
End Property

Public Property Let TabBorderType(ByVal New_BorderType As enmBorderType)
    
    '-- Initiate Tab Dorder Style...
    Select Case New_BorderType
        Case Is = 0
            mBorderType = enmNone
            m_BorderType = 0
        Case Is = 1
            mBorderType = BDR_RAISED
            m_BorderType = 1
        Case Is = 2
            mBorderType = EDGE_SUNKEN
            m_BorderType = 2
        Case Is = 3
            mBorderType = EDGE_BUMP
            m_BorderType = 3
        Case Is = 4
           mBorderType = EDGE_ETCHED
           m_BorderType = 4
        Case Is = 5
           mBorderType = EDGE_THIN
           m_BorderType = 5
    
    End Select
    
    PropertyChanged "BorderType"

    DrawTabs

End Property


'-- Exposed methods...

Public Sub RemoveTab(Key As String)
'-- Remove the tab specified by Key...
    
    Dim i       As Long
    Dim SelTab  As Long
    
    '-- Find the specified tab...
    SelTab = TabFromKey(Key)
    If SelTab = 0 Then ErrRaise TAB_DOESNT_EXIST
    
    '-- If it's not the last tab then move others down to fill gap...
    If SelTab < mTabCount Then
        For i = SelTab To mTabCount - 1
            LSet mTabs(i) = mTabs(i + 1)
        Next
    End If
    
    '-- Reduce the tab count...
    mTabCount = mTabCount - 1
    
    '-- If active tab was deleted select a new one...
    
    If SelTab = mActiveTab Then
        
        If SelTab <= mTabCount Then
            mActiveTab = SelTab
        Else
            mActiveTab = mTabCount
        End If
        
        If mTabCount > 0 Then
            mTabs(mActiveTab).Active = True
            RaiseEvent Change(mActiveTab, mTabs(mActiveTab).Caption, mTabs(mActiveTab).Key)
        End If
    
    End If
    
    ReDim Preserve mTabs(mTabCount)
        
    UserControl_Resize
    DrawTabs

End Sub

Public Sub RemoveAllTabs()
'-- Remove all tabs...

    Erase mTabs()
    mTabCount = 0
    mLeftMostTab = 1
    UserControl_Resize
    DrawTabs
    tmrMouseOut.Interval = 1
    
End Sub

Public Sub AddTab(Caption As String, _
        Key As String, _
        Optional ToolTip As String = vbNullString, _
        Optional InsertBefore As String = vbNullString)

'-- Add a new tab.  Specify a tab to insert before by putting
'-- the Key of the desired tab in InsertBefore, defaults to
'-- after the last tab

    Dim i   As Long
    Dim Pos As Long
    
    '-- Check for duplicate keys...
    If mTabCount > 0 Then
        For i = 1 To mTabCount
            If Key = mTabs(i).Key Then ErrRaise TAB_EXISTS
        Next
    End If
    
    '-- Find position to put tab...
    If (mTabCount > 0) And (InsertBefore <> vbNullString) Then
        
        '-- locate correct tab, and move up...
        Pos = TabFromKey(InsertBefore)
        If Pos = 0 Then ErrRaise TAB_DOESNT_EXIST
        
        mTabCount = mTabCount + 1
        
        ReDim Preserve mTabs(mTabCount)
        
        For i = mTabCount - 1 To Pos Step -1
            LSet mTabs(i + 1) = mTabs(i)
        Next
        
        '-- Move up the active tab if necessary...
        If Pos <= mActiveTab Then
            mActiveTab = mActiveTab + 1
        End If
    Else
        '-- Either no position is specified, or there are no tabs
        '-- so insert at the end...
        If InsertBefore <> vbNullString Then ErrRaise TAB_DOESNT_EXIST
        
        Pos = mTabCount + 1
        mTabCount = Pos
        
        ReDim Preserve mTabs(mTabCount)
    
    End If
    
    '-- And add it...
    With mTabs(Pos)
        .Caption = Caption
        .Key = Key
        .ToolTipText = ToolTip
        .Active = False
    End With
    
    '-- If this is the first tab, it needs to be active...
    If mTabCount = 1 Then
        mActiveTab = 1
        mTabs(Pos).Active = True
    End If
    
    UserControl_Resize
    DrawTabs
    
End Sub

Private Sub pbTabs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '-- Only Draw The Tab If The Mouse Is
    '-- Actually Over The Tab And Not The Active One...
  '  If y > 18 Then Exit Sub
     
     ActivateNewTab Button, Shift, X, Y

    
End Sub

Private Sub tmrMouseOut_Timer()
    Dim P As POINTAPI
    Dim i As Long
    Dim rct As RECT
    
    GetCursorPos P
    
    If WindowFromPoint(P.X, P.Y) <> pbTabs.hWnd Then
            
            If mTabCount > 0 Then
                For i = 1 To mTabCount
                  If mTabs(i).Hovering = True Then
                    mTabs(i).Hovering = False
                 End If
                Next
            End If
     
    End If
    
      
 DrawTabs
       
End Sub

Private Sub UserControl_Initialize()
On Error GoTo ExitMe

'    If Ambient.UserMode = True Then
        tmrMouseOut.Interval = 1
'    End If

Exit Sub
ExitMe:
        tmrMouseOut.Interval = 1


End Sub




'
'-- Private methods...

Private Sub UserControl_InitProperties()
'-- Init default properties...
    On Error Resume Next
    
    Set Font = UserControl.Parent.Font
    
    TabBackColor = vbButtonFace
    
    TabTextColor = vbButtonText
    
    ActiveTabBackColor = &H8000000F
    
    ActiveTabTextColor = &H404040
    
    BackColor = &H8000000C
    
    tmrMouseOut.Enabled = True
    
    m_HoverColour = m_def_HoverColour

    m_AllignBottom = m_def_AllignBottom
    
    m_SeperatorLineColor = m_def_SeperatorLineColor
    
    m_SelectedTabHoverColor = m_def_SelectedTabHoverColor

    UserControl_Load

    m_BorderType = m_def_BorderType
    
    mBorderType = BDR_RAISED
    
    m_BorderType = m_def_BorderType
    
        
End Sub




Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-- Read stored properties...
    On Error Resume Next
    
    Set Font = PropBag.ReadProperty("Font")
    
    TabBackColor = PropBag.ReadProperty("TabBackColor", vbButtonFace)
    
    TabTextColor = PropBag.ReadProperty("TabTextColor", vbButtonText)
    
    ActiveTabBackColor = PropBag.ReadProperty("ActiveTabBackColor", vbWindowBackground)
    
    ActiveTabTextColor = PropBag.ReadProperty("ActiveTabTextColor", vbButtonText)
    
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    
    m_HoverColour = PropBag.ReadProperty("HoverColour", m_def_HoverColour)

    m_AllignBottom = PropBag.ReadProperty("AllignBottom", m_def_AllignBottom)
    
    m_SeperatorLineColor = PropBag.ReadProperty("SeperatorLineColor", m_def_SeperatorLineColor)
    
    m_SelectedTabHoverColor = PropBag.ReadProperty("SelectedTabHoverColor", m_def_SelectedTabHoverColor)

    UserControl_Load
    
    m_BorderType = PropBag.ReadProperty("BorderType", m_def_BorderType)
    
    m_BorderType = PropBag.ReadProperty("BorderType", m_def_BorderType)
    
        
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-- Store design time settings...
    
    PropBag.WriteProperty "Font", Font
    
    PropBag.WriteProperty "TabBackColor", TabBackColor, vbButtonFace
    
    PropBag.WriteProperty "TabTextColor", TabTextColor, vbButtonText
      
    PropBag.WriteProperty "ActiveTabBackColor", ActiveTabBackColor, vbWindowBackground
    
    PropBag.WriteProperty "ActiveTabTextColor", ActiveTabTextColor, vbButtonText
    
    PropBag.WriteProperty "BackColor", mBackColor, vbButtonFace


    Call PropBag.WriteProperty("HoverColour", m_HoverColour, m_def_HoverColour)
    
    Call PropBag.WriteProperty("AllignBottom", m_AllignBottom, m_def_AllignBottom)
    
    Call PropBag.WriteProperty("SeperatorLineColor", m_SeperatorLineColor, m_def_SeperatorLineColor)

    Call PropBag.WriteProperty("SelectedTabHoverColor", m_SelectedTabHoverColor, m_def_SelectedTabHoverColor)

    Call PropBag.WriteProperty("BorderType", m_BorderType, m_def_BorderType)
    
    Call PropBag.WriteProperty("BorderType", m_BorderType, m_def_BorderType)


End Sub

Private Sub UserControl_Load()

    mLeftMostTab = 1
    
    '-- Have design time view...
    If Ambient.UserMode = False Then
        
        tmrMouseOut.Interval = 0
        
        AddTab "GN Tab", "a", vbNullString
    
    End If
    
    
End Sub

Private Sub cmdScroll_Click(Index As Integer)
'-- Scroll the tabs...
    
    '-- No need to scroll if there are less than 2 tabs...
    If mTabCount > 1 Then
    
        If Index = 0 Then
            
            '-- Scroll view left...
            If mLeftMostTab > 1 Then mLeftMostTab = mLeftMostTab - 1
        
        Else
            '-- Scroll view right...
            If (mLeftMostTab < mTabCount) And _
                (((mTabs(mTabCount).left + mTabs(mTabCount).Width) _
                + pbTabs.left \ Screen.TwipsPerPixelX) > UserControl.Width \ Screen.TwipsPerPixelX) _
                Then mLeftMostTab = mLeftMostTab + 1
        
        End If
        
        '-- Redraw tabs...
        DrawTabs
    
    End If
    
    pbTabs.SetFocus
    UserControl.SetFocus
    
    
End Sub

Private Sub cmdScroll_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'-- Eliminate the focus rectangle on the scroll buttons...
    
    UserControl.SetFocus
    
End Sub

Private Sub ActivateNewTab(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-- Determine which tab has been selected...

    Dim i           As Long
    Dim PrevActive  As Long
    
    '-- Bail if there are no tabs...
    If mTabCount > 0 Then
    
        '-- Only respond to left mouse button...
        If Button = vbLeftButton Then
            
            '-- Store previously selected tab...
            PrevActive = mActiveTab
        
            '-- Find the clicked tab...
            For i = 1 To mTabCount
                
                If (X >= mTabs(i).left) And (X <= mTabs(i).left + mTabs(i).Width) Then
                    
                    mActiveTab = i
                
                End If
            
            Next
            
            '-- Activate it, setting other to inactive...
            For i = 1 To mTabCount
                
                mTabs(i).Active = (i = mActiveTab)
            
            Next
            
            '-- If active tab is off the left side, then move it back on...
            If mActiveTab < mLeftMostTab Then
                
                mLeftMostTab = mActiveTab
                
                CalcTabPos
            
            End If
            
            '-- If active tab is off the right side, then move it back on...
            If mbNeedButtons Then
                
                If (mTabs(mActiveTab).left + mTabs(mActiveTab).Width) * Screen.TwipsPerPixelX _
                        + pbTabs.left > UserControl.Width Then
                    
                    mLeftMostTab = mLeftMostTab + 1
                    
                    CalcTabPos
                
                End If
            
            End If
            
            '-- If active tab is different than last then raise the Change event
            '-- If mActiveTab <> PrevActive Then RaiseEvent Change(mTabs(mActiveTab).Key)...
            RaiseEvent Change(mActiveTab, mTabs(mActiveTab).Caption, mTabs(mActiveTab).Key)
            
        End If
        
        '-- Redraw the tabs...
        DrawTabs
    
    End If

            
End Sub

Private Sub pbTabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-- Set tooltip based on which tab we are hovering over...
    Dim i As Long
  
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If mTabCount > 0 Then
         
            
        For i = 1 To mTabCount
            If (X > mTabs(i).left) And (X < mTabs(i).left + mTabs(i).Width) Then
                   pbTabs.ToolTipText = mTabs(i).ToolTipText
                   mTabs(i).Hovering = True
                
                Else
                   
                   mTabs(i).Hovering = False
                   
            End If
        Next
    
    End If
    
      
End Sub

Private Sub UserControl_Resize()
    
   '-- Resize Events...
    CalcTabPos
    
        ' Default Height
    If UserControl.Height <> 350 Then
        
        UserControl.Height = 350
    
    End If
    
    If mTabCount < 1 Then pbTabs.Width = 1
    
    ' Show Scroll buttons If Needed..
    If UserControl.Width > pbTabs.Width Then
        
        cmdScroll(0).Visible = False
        
        cmdScroll(1).Visible = False
        
        mbNeedButtons = False
        
        mLeftMostTab = 1
        
        MsgBox "Scroll Bars Inivisible, now."
        
        DrawTabs
        
        pbTabs.SetFocus
        
        UserControl.SetFocus
    
    Else
        
        If mTabCount > 1 Then
            
            cmdScroll(0).Visible = True
            
            cmdScroll(1).Visible = True
            
            cmdScroll(0).Move 0, pbTop.Height, cmdScroll(0).Width, UserControl.Height - (pbTop.Height + pbBottom.Height)
            
            cmdScroll(1).Move cmdScroll(0).Width, pbTop.Height, cmdScroll(0).Width, UserControl.Height - (pbTop.Height + pbBottom.Height)
            
            mbNeedButtons = True
        
        End If
    
    End If
    
    pbTabs.Height = UserControl.Height
    
    pbTabs.top = 0
    
    CalcTabPos
    
    DrawTabs
    
    
End Sub

Private Sub CalcTabPos()
'-- Figure size and position of all tabs...
    
    Dim i As Long
    Dim X As Long
    
    '-- Bail if there are no tabs...
    If mTabCount <> 0 Then
    
        '--  Set position of first tab...
        mTabs(1).left = 5
        
        '-- Get width as position of tabs...
        For i = 1 To mTabCount
            
            CalculateWidth (i)
            
            If i > 1 Then mTabs(i).left = (mTabs(i - 1).Width + mTabs(i - 1).left) + 3
        
        Next
            
            
        '-- Set width of pbTabs...
        If mTabCount > 0 Then
            
            pbTabs.Width = (mTabs(mTabCount).Width + mTabs(mTabCount).left + 10) * Screen.TwipsPerPixelX
        
        Else
            
            pbTabs.Width = 1
        
        End If
        
        
        If mLeftMostTab > 0 And mLeftMostTab <= mTabCount Then
            
            X = mTabs(mLeftMostTab).left - 4
        
        End If
        
        
        '-- Set position of pbTabs...
        If mbNeedButtons Then
            
            pbTabs.left = (cmdScroll(0).Width * 2) - (X * Screen.TwipsPerPixelX)
        
        Else
            
            pbTabs.left = X * Screen.TwipsPerPixelX
        
        End If
    
    End If
    
End Sub

Private Sub DrawTabs()
'-- Draw all the tabs...

    Dim i       As Long
    Dim Active  As Long
    
    '-- Figure tab positions and sizes...
    CalcTabPos
    
    UserControl.BackColor = mBackColor
    UserControl.Cls
    
    pbTabs.BackColor = mBackColor
    pbTabs.ForeColor = mTabLineColor
    
    ' Change The Top And Bottom Colour Of The Control..
    If AllignBottom Then
       
       pbTop.BackColor = ActiveTabBackColor
       
       pbBottom.BackColor = mBackColor
    
    Else
       
       pbTop.BackColor = mBackColor
       
       pbBottom.BackColor = ActiveTabBackColor
    
    End If
    
    
    
    pbTabs.Cls
    
    '-- If there are tabs to draw...
    If mTabCount > 0 Then
    
        '-- Draw the tabs, so that rightmost ones are on top...
        For i = mTabCount To 1 Step -1
                      
            If mTabs(i).Active = True Then
                Active = i
            Else
               
                DrawTab i
            End If
        Next
        
        '-- Draw the active one last, so that it is always on top...
        DrawTab Active
        
    End If
    
    pbTabs.Refresh
    
End Sub

Private Sub DrawTab(TabNum As Long)
'-- Draw an individual tab...

    Dim TabColor    As Long
    Dim hRGN        As Long
    Dim hBRSH       As Long
    Dim hOldPen     As Long
    Dim hPen        As Long
    Dim rct         As RECT
    Dim bWasBold    As Boolean
    
    
    CalculateWidth TabNum
    
    '-- Set up our RECT structure for DrawText...
    rct.left = mTabs(TabNum).left
    
    rct.Right = mTabs(TabNum).left + mTabs(TabNum).Width
    
    rct.top = 1
    
    rct.Bottom = pbTabs.ScaleHeight - 1
    
    
    
    '-- Correct The background color...
    If mTabs(TabNum).Active Then
        
        TabColor = ConvertColor(mActiveTabBackColor)
    
    Else
        
        TabColor = ConvertColor(mTabBackColor)
    
    End If
    
   
    '-- If the tab is active, we want a bold font...
    If Not mTabs(TabNum).Hovering Then
        
        If mTabs(TabNum).Active Then
            
            bWasBold = pbTabs.Font.Bold
            
            pbTabs.Font.Bold = True
            
            pbTabs.Font.Underline = False
        
        Else
            
            pbTabs.Font.Underline = False
            
            pbTabs.Font.Bold = False
        
        End If
    
    
    Else
        
        If mTabs(TabNum).Active Then
            
            bWasBold = pbTabs.Font.Bold
            
            pbTabs.Font.Bold = True
            
            pbTabs.Font.Underline = False
        
        Else
            
            pbTabs.Font.Underline = True
            
            pbTabs.Font.Bold = False
        
        
        End If
    
    
    End If
    
    
    '-- Draw The Appropriate Tab Back Color
    '-- And Set The Rct For The Text And Border Style...
    
    If Not AllignBottom Then
        
        If Not mTabs(TabNum).Hovering Then
    
                If mTabs(TabNum).Active Then
                    
                    'Draw The Tab...
                    pbTabs.Line (mTabs(TabNum).left + mTabs(TabNum).Width, 4)-((mTabs(TabNum).Width + mTabs(TabNum).left - mTabs(TabNum).Width), pbTabs.Height + 315), Me.ActiveTabBackColor, BF
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(mActiveTabTextColor))
                
                Else
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(mTabTextColor))
                
                End If '-- If mTabs(TabNum).Active Then
            
            
            Else
                
                If mTabs(TabNum).Active Then
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(mActiveTabTextColor))
                    
                    pbTabs.Line (mTabs(TabNum).left + mTabs(TabNum).Width, 4)-((mTabs(TabNum).Width + mTabs(TabNum).left - mTabs(TabNum).Width), pbTabs.Height + 315), Me.ActiveTabBackColor, BF
                    
                    '-- Readjust Rectangle...
                    rct.top = rct.top + 3
                    
                    rct.Right = rct.Right + 2
                    
                    '-- Draw The Edge Of The Tab...
                    DrawTabEdge rct
                    
                    '-- Readjust Rectangle For Text...
                    rct.top = rct.top - 2
                    
                    rct.Right = rct.Right - 2
                
                Else
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(HoverColour))
                
                End If '-- If mTabs(TabNum).Active Then
            
            End If '-- If Not mTabs(TabNum).Hovering Then
        
        Else
               
            If Not mTabs(TabNum).Hovering Then
               
               If mTabs(TabNum).Active Then
                    
                    pbTabs.Line (mTabs(TabNum).left + mTabs(TabNum).Width, 0)-((mTabs(TabNum).Width + mTabs(TabNum).left - mTabs(TabNum).Width), pbTabs.Height - 327), Me.ActiveTabBackColor, BF
                    
                    rct.top = rct.top - 2
                    
                    rct.left = rct.left '+ 1
     
                    Call SetTextColor(pbTabs.hDC, ConvertColor(mActiveTabTextColor))
                
                Else
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(mTabTextColor))
                
                End If '--  If mTabs(TabNum).Active Then
            
            Else
                
                If mTabs(TabNum).Active Then
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(SelectedTabHoverColor))
                    
                    pbTabs.Line (mTabs(TabNum).left + mTabs(TabNum).Width, 0)-((mTabs(TabNum).Width + mTabs(TabNum).left - mTabs(TabNum).Width), pbTabs.Height - 100), Me.ActiveTabBackColor, BF
                    
                    '-- Readjust Rectangle...
                    rct.top = rct.top - 5
                    
                    rct.left = rct.left
                    
                    rct.Bottom = rct.Bottom - 2.5
                    
                    rct.Right = rct.Right + 1
                    
                    '-- Draw The Edge Of The Tab...
                    DrawTabEdge rct
                    
                    '-- Readjust Rectangle For Text...
                    rct.Right = rct.Right - 1
                    
                    rct.top = rct.top + 3
                
                Else
                    
                    Call SetTextColor(pbTabs.hDC, ConvertColor(HoverColour))
                
                End If '-- If mTabs(TabNum).Active Then
            
            End If '-- If Not mTabs(TabNum).Hovering Then
            
            
    
    End If  '-- If Not AllignBottom Then
    
    
    
    '-- Draw The Actual Text Of The Tab...
    If Not AllignBottom Then
        
        Call DrawText(pbTabs.hDC, mTabs(TabNum).Caption, Len(mTabs(TabNum).Caption), _
                rct, DT_CENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_VCENTER)
    
    Else
        
        rct.top = rct.top + 4
        
        If Not mTabs(TabNum).Active Then
            rct.Bottom = rct.Bottom - 3
        Else
            rct.Bottom = rct.Bottom - 5
        End If
        
        Call DrawText(pbTabs.hDC, mTabs(TabNum).Caption, Len(mTabs(TabNum).Caption), _
                rct, DT_CENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_VCENTER)
    
    End If '--     If Not AllignBottom Then
    
    
    '-- Set font back to what it was...
    '-- And Set The Seperator Line
    If mTabs(TabNum).Active Then
        
        pbTabs.Font.Bold = bWasBold
    
    Else
        
        pbTabs.Line (mTabs(TabNum).left + mTabs(TabNum).Width + 4, 5.5)-(mTabs(TabNum).Width + mTabs(TabNum).left + 4, UserControl.ScaleHeight - 330), SeperatorLineColor, BF
    
    End If
    
   ' MsgBox mBorderType

End Sub


Private Function DrawTabEdge(rct As RECT)
' Draw The Edge Round The Tab...

    Select Case TabBorderType
        Case Is = 0
            Call DrawEdge(pbTabs.hDC, rct, 0, BF_RECT)
        Case Is = 1
            Call DrawEdge(pbTabs.hDC, rct, BDR_RAISED, BF_RECT)
        Case Is = 2
            Call DrawEdge(pbTabs.hDC, rct, EDGE_SUNKEN, BF_RECT)
        Case Is = 3
            Call DrawEdge(pbTabs.hDC, rct, EDGE_BUMP, BF_RECT)
        Case Is = 4
            Call DrawEdge(pbTabs.hDC, rct, EDGE_ETCHED, BF_RECT)
        Case Is = 5
            Call DrawEdge(pbTabs.hDC, rct, EDGE_THIN, BF_RECT)
    
    End Select

End Function
Private Function ConvertColor(OleColor As Long) As Long
' converts from the OLE_COLOR type to COLORREF

    Dim r As Long
    Call OleTranslateColor(OleColor, 0, r)
    ConvertColor = r
    
End Function

Private Sub CalculateWidth(TabNum)
'-- Calculates the width of an individual tab
    Dim rct         As RECT
    Dim bWasBold    As Boolean
    
    bWasBold = pbTabs.Font.Bold
    
    '-- Figure the size for a bold font...
    pbTabs.Font.Bold = True
    
    Call DrawText(pbTabs.hDC, mTabs(TabNum).Caption, Len(mTabs(TabNum).Caption), _
        rct, DT_CENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_CALCRECT)
        
    '-- Change the font back...
    pbTabs.Font.Bold = bWasBold
    
    '-- Set The width...
    mTabs(TabNum).Width = rct.Right + 4
    
End Sub

Private Function TabFromKey(Key As String) As Long
'-- Find the index of a tab based on it's key...
    
    Dim i As Long
    Dim r As Long
    
    If mTabCount > 0 Then
        For i = 1 To mTabCount
            If mTabs(i).Key = Key Then
                r = i
                Exit For
            End If
        Next
    End If
    
    TabFromKey = r
    
End Function

Private Sub ErrRaise(Num As Long)
'-- Error Handling - Raise appropriate errors...
    Select Case Num
        Case TAB_EXISTS
            Err.Raise TAB_EXISTS, "gnTabCtl", "Tab with specified key already exists"
        Case TAB_DOESNT_EXIST
            Err.Raise TAB_DOESNT_EXIST, "gnTabCtl", "Tab with specified key does not exist"
    End Select
End Sub





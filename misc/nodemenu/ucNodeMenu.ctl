VERSION 5.00
Begin VB.UserControl nMenu 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7710
   ScaleHeight     =   5070
   ScaleWidth      =   7710
   Begin VB.Timer tmrReUpdate 
      Interval        =   100
      Left            =   4920
      Top             =   240
   End
   Begin VB.Image imgShadow 
      Height          =   255
      Left            =   4320
      Picture         =   "ucNodeMenu.ctx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Shape shpCoverBorder 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Left            =   2640
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape shpActivePic 
      BackColor       =   &H00E8EEEE&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F79652&
      Height          =   300
      Index           =   0
      Left            =   3360
      Top             =   1200
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgSeperate 
      Height          =   30
      Left            =   2160
      Picture         =   "ucNodeMenu.ctx":010E
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgExpand 
      Height          =   255
      Index           =   0
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   720
      Width           =   255
   End
   Begin VB.Image imgExpandPic 
      Height          =   240
      Left            =   2040
      Picture         =   "ucNodeMenu.ctx":01B8
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCheck 
      Height          =   240
      Left            =   3000
      Picture         =   "ucNodeMenu.ctx":020A
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Index           =   0
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   720
      Width           =   240
   End
   Begin VB.Image imgMenuInactive 
      Height          =   285
      Left            =   1560
      Picture         =   "ucNodeMenu.ctx":025C
      Top             =   2160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgMenuBack 
      Height          =   285
      Index           =   0
      Left            =   600
      Picture         =   "ucNodeMenu.ctx":120E
      Stretch         =   -1  'True
      Tag             =   "0"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image imgArea 
      Height          =   255
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00422306&
      BackStyle       =   0  'Transparent
      Caption         =   "|"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00602B0B&
      Height          =   210
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Shape shpMenu 
      BackColor       =   &H00E4EBEB&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00EFEFEF&
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpBorder 
      BackColor       =   &H00370D02&
      BorderColor     =   &H00C06950&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Image imgBack 
      Height          =   375
      Left            =   0
      Picture         =   "ucNodeMenu.ctx":21C0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "nMenu"
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

Private m_Regular_Background As Long
Private m_Regular_Border As Long
Private m_Font_Active_Color As Long
Private m_Shape_Active_Color As Long
Private m_Shape_Active_Border As Long
Private m_Shape_Active_Horizontal_Color As Long
Private m_Shape_Active_Horizontal_Border As Long
Private m_Font_Inactive_Color As Long
Private m_Shape_Inactive_Color As Long
Private m_Shape_Inactive_Color_Horizontal As Long
Private m_Seperator_Color As Long
Private m_Check_Inactive_Back As Long
Private m_Seperator_BackColor As Long
Private m_Check_Active_Back As Long
Private m_Check_Active_Border As Long
Private m_Check_InActive_Border As Long


Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type tNodeMenu
    Caption As String
    SubMenu As Object
    IsSeperator As Boolean
    TooltipText As String
    Enabled As Boolean
    Checked As Boolean
End Type

Public Event MenuClick(SubMenuIndex As Integer)
Public Event PopupMove()
Public Event PopupFinish()

Public AutoResize As Boolean
Public Menus_Active As Boolean
Public Main_Active As Integer
Public SubLeft As Integer
Public SubTop As Integer
Public Creator As Object
Public CreatorWidth As Integer
Public Popup As Boolean
Private Started As Boolean
Private m_Horizontal As Boolean
Private m_CharSet As Integer
Private Menus() As tNodeMenu
Private Just_Clicked As Boolean
Public Sub LoadDefaultProperties()
    m_Regular_Background = 14215660
    m_Regular_Border = 8029834
    m_Font_Active_Color = &H602B0B
    m_Shape_Active_Color = 15651521 '&HF2D6C2
    m_Shape_Active_Border = &HF79652
    m_Shape_Active_Horizontal_Color = 15265518
    m_Shape_Active_Horizontal_Border = 8029834 '&HC06950
    m_Font_Inactive_Color = &HC06950 '&HF9D8C4
    m_Shape_Inactive_Color = &HFFFFFF
    m_Shape_Inactive_Color_Horizontal = 14215660 '&HE8EEEE
    m_Seperator_Color = &HC06950
    m_Check_Inactive_Back = &HE8EEEE
    m_Seperator_BackColor = 12108485
    m_Check_Active_Back = m_Shape_Active_Color
    m_Check_Active_Border = m_Shape_Active_Border
    m_Check_InActive_Border = m_Check_Active_Border
    'UpdateAppearence
End Sub
Public Property Let Color_Regular_Background(ByVal NewValue As Long)
    m_Regular_Background = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Regular_Border(ByVal NewValue As Long)
    m_Regular_Border = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Font_Active_Color(ByVal NewValue As Long)
    m_Font_Active_Color = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Shape_Active_Color(ByVal NewValue As Long)
    m_Shape_Active_Color = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Shape_Active_Horizontal_Border(ByVal NewValue As Long)
    m_Shape_Active_Horizontal_Border = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Font_Inactive_Color(ByVal NewValue As Long)
    m_Font_Inactive_Color = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Shape_Inactive_Color_Horizontal(ByVal NewValue As Long)
    m_Shape_Inactive_Color_Horizontal = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Seperator_Color(ByVal NewValue As Long)
    m_Seperator_Color = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Check_Inactive_Back(ByVal NewValue As Long)
    m_Check_Inactive_Back = NewValue
    UpdateAppearence
End Property
Public Property Let Color_Seperator_BackColor(ByVal NewValue As Long)
    m_Seperator_BackColor = NewValue
    UpdateAppearence
End Property
Public Property Let Shadow(ByVal FileName As String)
    On Error GoTo R_Error
    Set imgShadow.Picture = LoadPicture(FileName)
    Exit Property
R_Error:
    Err.Raise vbObjectError + 1, "Property Let Shadow", Err.Description
End Property
Private Sub UpdateAppearence()
    Dim CurrentCount As Integer
    
    shpBorder.BorderColor = m_Regular_Border
    UserControl.BackColor = m_Regular_Background
    For CurrentCount = 0 To shpMenu.Count - 1
        If m_Horizontal Then
            shpMenu(CurrentCount).BackColor = m_Shape_Inactive_Color_Horizontal
        Else
            shpMenu(CurrentCount).BackColor = m_Shape_Inactive_Color
            If Menus(CurrentCount).IsSeperator Then
                shpMenu(CurrentCount).BackColor = m_Seperator_BackColor
            End If
        End If
        lblCaption(CurrentCount).ForeColor = m_Font_Active_Color
    Next CurrentCount
End Sub
Public Property Let CharSet(intValue As Integer)
    m_CharSet = intValue
End Property
Public Property Let Horizontal(boolValue As Boolean)
    m_Horizontal = boolValue
End Property
Public Property Get Horizontal() As Boolean
    Horizontal = m_Horizontal
End Property
Public Property Let Checked(MenuIndex As Integer, boolValue As Boolean)
    Menus(MenuIndex + 1).Checked = boolValue
    If boolValue Then
        imgIcon(MenuIndex + 1).Visible = True
        If imgIcon(MenuIndex + 1).Picture = 0 Or imgIcon(MenuIndex + 1).Picture Is imgCheck.Picture Then
            Set imgIcon(MenuIndex + 1).Picture = imgCheck.Picture
        End If
    Else
        If imgIcon(MenuIndex + 1).Picture Is imgCheck.Picture Then
            imgIcon(MenuIndex + 1).Visible = False
        End If
    End If
    shpActivePic(MenuIndex + 1).Visible = boolValue
End Property
Public Property Get Checked(MenuIndex As Integer) As Boolean
    Checked = Menus(MenuIndex + 1).Checked
End Property
Public Sub AddMenu(strCaption As String, Optional SubMenu As Object = Nothing, Optional stdIcon As StdPicture, Optional TooltipText As String = "")
    Dim CurrentCount As Integer
    
    CurrentCount = lblCaption.Count
    ReDim Preserve Menus(CurrentCount)
    Menus(CurrentCount).Caption = strCaption
    Menus(CurrentCount).TooltipText = TooltipText
    If Not SubMenu Is Nothing Then
        Set Menus(CurrentCount).SubMenu = SubMenu
    End If
    Load lblCaption(CurrentCount)
    Load shpMenu(CurrentCount)
    Load imgArea(CurrentCount)
    Load imgMenuBack(CurrentCount)
    Load imgIcon(CurrentCount)
    Load imgExpand(CurrentCount)
    Load shpActivePic(CurrentCount)
    
    lblCaption(CurrentCount).Visible = True
    lblCaption(CurrentCount).Font.CharSet = m_CharSet
    lblCaption(CurrentCount).Font.Name = ""
    'shpMenu(CurrentCount).Visible = True
    imgArea(CurrentCount).Visible = True
    'imgMenuBack(CurrentCount).Visible = True
    
    If Not SubMenu Is Nothing And Not m_Horizontal Then
        imgExpand(CurrentCount).Visible = True
        imgExpand(CurrentCount).Picture = imgExpandPic.Picture
    End If
    If Not stdIcon Is Nothing Then
        imgIcon(CurrentCount).Visible = True
        Set imgIcon(CurrentCount).Picture = stdIcon
    End If
    
    lblCaption(CurrentCount).Caption = strCaption
    If CBool(Len(TooltipText)) Then
        imgArea(CurrentCount).TooltipText = TooltipText
    End If
    If m_Horizontal Then
        shpMenu(CurrentCount).BackColor = m_Shape_Inactive_Color_Horizontal
    Else
        shpMenu(CurrentCount).BackColor = m_Shape_Inactive_Color
    End If
        
    If m_Horizontal Then
        lblCaption(CurrentCount).Left = lblCaption(CurrentCount - 1).Left + lblCaption(CurrentCount - 1).Width + 100 * 2
        lblCaption(CurrentCount).AutoSize = False
        lblCaption(CurrentCount).Width = lblCaption(CurrentCount).Width + 50 * 2
        shpMenu(CurrentCount).Width = lblCaption(CurrentCount).Width + 50 * 2
        shpMenu(CurrentCount).Left = lblCaption(CurrentCount).Left - 30
        shpMenu(CurrentCount).Top = 20 '15
        shpMenu(CurrentCount).Height = 255 '330 '255
        lblCaption(CurrentCount).Top = shpMenu(CurrentCount).Top + 10
    Else
        If strCaption = "-" Then
            Menus(CurrentCount).IsSeperator = True
            lblCaption(CurrentCount).Caption = ""
            lblCaption(CurrentCount).Height = 30 'imgSeperate.Height
            'imgMenuBack(CurrentCount).Picture = imgSeperate.Picture
            shpMenu(CurrentCount).Height = lblCaption(CurrentCount).Height
            shpMenu(CurrentCount).BackColor = m_Seperator_BackColor
            shpMenu(CurrentCount).Left = imgBack.Left + imgBack.Width + 80
            shpMenu(CurrentCount).Visible = True
            shpMenu(CurrentCount).Top = shpMenu(CurrentCount - 1).Top + shpMenu(CurrentCount - 1).Height + 20
            shpMenu(CurrentCount).ZOrder 0
        Else
            shpMenu(CurrentCount).Height = 330
            shpMenu(CurrentCount).BackColor = m_Shape_Inactive_Color
            shpMenu(CurrentCount).Left = 30 'lblCaption(CurrentCount).Left - 90
            shpMenu(CurrentCount).Top = shpMenu(CurrentCount - 1).Top + shpMenu(CurrentCount - 1).Height
        End If
        lblCaption(CurrentCount).Left = lblCaption(0).Left + lblCaption(0).Width + imgBack.Width
        shpMenu(CurrentCount).Width = lblCaption(CurrentCount).Width + lblCaption(CurrentCount).Left - 20 + imgExpand(0).Width
        'If Not SubMenu Is Nothing Then
        '    shpMenu(CurrentCount).Width = shpMenu(CurrentCount).Width
        'End If
        shpMenu(CurrentCount).BorderStyle = 0
        If CurrentCount = 1 Then
            shpMenu(CurrentCount).Top = 25
        End If
        lblCaption(CurrentCount).Top = shpMenu(CurrentCount).Top + shpMenu(CurrentCount).Height / 2 - lblCaption(CurrentCount).Height / 2
    End If
    imgArea(CurrentCount).Left = 0 'shpMenu(CurrentCount).Left - imgIcon(CurrentCount).Width
    imgArea(CurrentCount).Top = shpMenu(CurrentCount).Top
    imgArea(CurrentCount).Width = shpMenu(CurrentCount).Width + imgIcon(CurrentCount).Width
    imgArea(CurrentCount).Height = shpMenu(CurrentCount).Height
    If Not m_Horizontal Then
        imgIcon(CurrentCount).Left = imgArea(CurrentCount).Left - 40
    Else
        imgIcon(CurrentCount).Left = imgArea(CurrentCount).Left + 40
    End If
    imgIcon(CurrentCount).Top = imgArea(CurrentCount).Top + imgArea(CurrentCount).Height / 2 - imgIcon(CurrentCount).Height / 2
    imgIcon(CurrentCount).Left = imgArea(CurrentCount).Height / 2 - imgIcon(CurrentCount).Height / 2
    imgExpand(CurrentCount).Left = imgIcon(CurrentCount).Left + imgArea(CurrentCount).Width - imgExpand(CurrentCount).Width
    imgExpand(CurrentCount).Top = imgIcon(CurrentCount).Top
    imgMenuBack(CurrentCount).ZOrder 0
    lblCaption(CurrentCount).ZOrder 0
    shpActivePic(CurrentCount).ZOrder 0
    imgIcon(CurrentCount).ZOrder 0
    imgExpand(CurrentCount).ZOrder 0
    imgArea(CurrentCount).ZOrder 0
    imgMenuBack(CurrentCount).Left = imgArea(CurrentCount).Left
    imgMenuBack(CurrentCount).Top = imgArea(CurrentCount).Height
    imgMenuBack(CurrentCount).Width = imgArea(CurrentCount).Width
    imgMenuBack(CurrentCount).Height = imgArea(CurrentCount).Height
    lblCaption(CurrentCount).ForeColor = m_Font_Inactive_Color
    shpBorder.ZOrder 1
    imgBack.ZOrder 1
    shpBack.ZOrder 1
End Sub
Public Property Let MenuEnabled(ByVal Index As Integer, ByVal NewValue As Boolean)
    Menus(Index).Enabled = NewValue
End Property
Public Sub EndMenu()
    Dim smallShape As Shape
    Dim maxWidth As Integer
    If Not m_Horizontal Then
        shpBorder.BorderStyle = 1
        For Each smallShape In shpMenu
            If maxWidth < smallShape.Width Then
                maxWidth = smallShape.Width
            End If
        Next smallShape
        maxWidth = maxWidth + 100
        For Each smallShape In shpMenu
            smallShape.Width = maxWidth
        Next smallShape
        UserControl.Height = shpMenu(shpMenu.Count - 1).Top + shpMenu(shpMenu.Count - 1).Height + 40
        UserControl.Width = shpMenu(0).Left + shpMenu(0).Width + 60 '+ 100 + imgBack.Width - 50
        shpBack.Visible = True
        shpBack.Height = UserControl.Height
        shpBack.Width = UserControl.Width
        shpBack.Left = imgBack.Left + imgBack.Width
        shpBack.Top = 0
    Else
        shpBorder.BorderStyle = 0
        If AutoResize Then
            UserControl.Height = shpMenu(shpMenu.Count - 1).Top + shpMenu(shpMenu.Count - 1).Height + 25
        End If
        shpBack.Visible = False
    End If
    For Each smallShape In shpMenu
        With imgArea(smallShape.Index)
            If m_Horizontal Then
                .Left = smallShape.Left
                If Not m_Horizontal Then
                    .Width = UserControl.Width
                Else
                    .Width = shpMenu(smallShape.Index).Width
                End If
            Else
                .Left = 0
                .Width = smallShape.Width
                If Menus(smallShape.Index).IsSeperator Then
                    'shpMenu(smallShape.Index).Width = UserControl.Width - shpMenu(smallShape.Index).Left - 20
                End If
            End If
            .Top = smallShape.Top
            .Height = smallShape.Height
        End With
        With imgMenuBack(smallShape.Index)
            .Left = smallShape.Left
            .Top = smallShape.Top
            .Width = smallShape.Width
            .Height = smallShape.Height
        End With
        With imgIcon(smallShape.Index)
            .Left = smallShape.Left + 20 + imgArea(smallShape.Index).Height / 2 - .Height / 2
            .Top = smallShape.Top + 20 + imgArea(smallShape.Index).Height / 2 - .Height / 2
        End With
        With shpActivePic(smallShape.Index)
            .Left = smallShape.Left + 20
            .Top = smallShape.Top + 20
        End With
        With imgExpand(smallShape.Index)
            .Left = smallShape.Left - .Width + smallShape.Width
            .Top = smallShape.Top + smallShape.Height / 2 - imgExpand(smallShape.Index).Height / 2
        End With
    Next smallShape
    Main_Active = -1
End Sub
Public Sub HideSubMenus()
    Dim i As Integer
    Dim CurrentSubMenu As Object
    Dim errln As Integer
    On Error Resume Next
    For i = 1 To UBound(Menus)
        Set CurrentSubMenu = Menus(i).SubMenu
        If Not CurrentSubMenu Is Nothing Then
            CurrentSubMenu.HideSubMenus
            CurrentSubMenu.Visible = False
        End If
    Next i
    Menus_Active = False
    DoUpdate
    shpCoverBorder.Visible = False
End Sub
Public Sub HideAllParents()
    HideSubMenus
    If Popup Then
        RaiseEvent PopupFinish
    End If
    If Not Creator Is Nothing Then
        Creator.HideAllParents
    End If
End Sub
Private Sub imgArea_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Menus_Active = Not Menus_Active
    If Menus(Index).IsSeperator Then
        Exit Sub
    End If
    If m_Horizontal And Menus(Index).SubMenu Is Nothing Then
        Menus_Active = False
        'GoTo Just_Fire
    End If
    If Not m_Horizontal And Not Menus(Index).SubMenu Is Nothing Then
        Menus_Active = True
        If Not Main_Active = Index Then
            Main_Active = Index
        End If
        HideSubMenus
        DoUpdate
        Just_Clicked = True
        Exit Sub
    End If
    If m_Horizontal And Not Menus_Active Then
        HideSubMenus
    End If
    Main_Active = Index
    DoUpdate
    If Menus(Index).SubMenu Is Nothing Then
        If Popup Then
            RaiseEvent PopupFinish
        End If
        If Not Creator Is Nothing Then
            Creator.HideAllParents
        End If
    End If
    RaiseEvent MenuClick(Index - 1)
End Sub
Private Sub imgArea_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Menus(Index).IsSeperator Then
        Exit Sub
    End If
    If Menus_Active Then
        If Not Main_Active = Index Then
            Main_Active = Index
            DoUpdate
        End If
    ElseIf m_Horizontal Then
        shpCoverBorder.Visible = False
        If Not Main_Active = Index Then
            Main_Active = Index
            DoUpdate
        End If
    End If
End Sub
Public Sub ShowCover()
    shpCoverBorder.ZOrder 0
    shpCoverBorder.Visible = True
    shpCoverBorder.Left = 10
    shpCoverBorder.Width = CreatorWidth - 40
    shpCoverBorder.Top = -20
    shpCoverBorder.Height = 30
End Sub
Private Sub imgBack_Click()
    Menus_Active = False
End Sub
Private Sub tmrReUpdate_Timer()
    Dim AreaPic As Image
    Dim CurrentPOS As POINTAPI
    If m_Horizontal And Not Menus_Active Then
        On Error GoTo Project_Running_Interuption
        GetCursorPos CurrentPOS
        CurrentPOS.x = ScaleX(CurrentPOS.x, vbPixels, vbTwips) - UserControl.Parent.Left - (UserControl.Parent.Width - UserControl.Parent.ScaleWidth)
        CurrentPOS.y = ScaleY(CurrentPOS.y, vbPixels, vbTwips) - UserControl.Parent.Top - (UserControl.Parent.Height - UserControl.Parent.ScaleHeight)
        For Each AreaPic In imgArea
            If AreaPic.Index > 0 And Main_Active = AreaPic.Index Then
                'check for mouseout
                If CurrentPOS.x < AreaPic.Left Or CurrentPOS.x > AreaPic.Left + AreaPic.Width Or _
                   CurrentPOS.y < AreaPic.Top - 50 Or CurrentPOS.y > AreaPic.Top + AreaPic.Height - 50 Then
                    Menus_Active = False
                    Main_Active = 0
                    DoUpdate
                    'UserControl.Parent.Line (CurrentPOS.x, CurrentPOS.y)-(CurrentPOS.x + 10, CurrentPOS.y + 10), RGB(0, 0, 0)
                End If
            End If
        Next AreaPic
    End If
Project_Running_Interuption:
    'if the project including the menus
    'breaks execution, we won't be able
    'to call the APIs.
End Sub
Private Sub UserControl_Initialize()
    AutoResize = True
    LoadDefaultProperties
    imgBack.Left = 0
    imgBack.Top = 0
    shpBorder.BorderColor = m_Regular_Border
    UserControl.BackColor = m_Regular_Background
End Sub
Private Sub UserControl_LostFocus()
    On Error Resume Next
    If Not (m_Horizontal Or Menus(Main_Active).SubMenu Is Nothing) Then
        Exit Sub
    End If
    If Just_Clicked Then
        Just_Clicked = False
    Else
        Menus_Active = False
        HideSubMenus
    End If
    DoUpdate
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    If m_Horizontal And AutoResize Then
        UserControl.Width = UserControl.Parent.Width
    End If
    shpBorder.Width = UserControl.Width - 5
    shpBorder.Height = UserControl.Height - 5
    'imgBack.Width = UserControl.Width
    imgBack.Height = UserControl.Height
    shpBack.Height = UserControl.Height
    shpBack.Width = UserControl.Width
End Sub
Public Sub RaisePopupMove()
    RaiseEvent PopupMove
    Menus_Active = True
    Main_Active = 0
    DoUpdate
End Sub
Public Sub FixPosition(objThis As Object, objCreator As Object)
    objThis.Left = objCreator.Left + objThis.Left
    objThis.Top = objCreator.Top + objThis.Top
    If objThis.Left + objThis.Width + objThis.Parent.Left > Screen.Width Then
        objThis.Left = objCreator.Left - objThis.Width + 25
    End If
    If objThis.Top + objThis.Height + objThis.Parent.Top > Screen.Height Then
        objThis.Top = objCreator.Top - objThis.Height + 25
    End If
End Sub
Public Sub Refresh()
    If Main_Active = -1 Then
        Main_Active = 0
    End If
    DoUpdate
End Sub
Private Sub DoUpdate()
    Dim InactiveLabel As Label
    Dim ThisSubmenu As Object
    
    If Not Menus_Active Then
        imgShadow.Visible = False
        shpCoverBorder.Visible = False
    End If
    
    If Menus_Active Or m_Horizontal Then
        lblCaption(Main_Active).ForeColor = m_Font_Active_Color
        If Main_Active > 0 Then
            If m_Horizontal And Menus_Active Then
                If Menus(Main_Active).SubMenu Is Nothing Then
                    'horizontal: no submenus
                    imgShadow.Visible = False
                    shpCoverBorder.Visible = False
                Else
                    'gray for:
                    'Horizontal, active(clicked & submenu visible)
                    shpMenu(Main_Active).BackColor = m_Shape_Active_Horizontal_Color
                    shpMenu(Main_Active).BorderColor = m_Shape_Active_Horizontal_Border
                    shpCoverBorder.ZOrder 0
                    'shpCoverBorder.Visible = True
                    shpCoverBorder.Left = shpMenu(Main_Active).Left + 10
                    shpCoverBorder.Width = shpMenu(Main_Active).Width - 40
                    shpCoverBorder.Top = shpMenu(Main_Active).Top + shpMenu(Main_Active).Height - 20
                    shpCoverBorder.Height = 2000
                    imgShadow.ZOrder 0
                    imgShadow.Visible = True
                    imgShadow.Left = shpMenu(Main_Active).Left + shpMenu(Main_Active).Width - 10
                    imgShadow.Top = shpMenu(Main_Active).Top
                    shpMenu(Main_Active).Height = 300
                    shpMenu(Main_Active).BackColor = m_Shape_Inactive_Color_Horizontal
                End If
            Else
                'blue shapes for:
                '1) Vertical, mouseover
                shpMenu(Main_Active).BackColor = m_Shape_Active_Color
                shpMenu(Main_Active).BorderColor = m_Shape_Active_Border
                If m_Horizontal Then
                    '2) Horizontal, mouseover (not clicked)
                    shpMenu(Main_Active).Height = 255
                End If
            End If
            shpMenu(Main_Active).BorderStyle = 1
            shpMenu(Main_Active).Visible = True
        
            If Menus(Main_Active).Checked Then
                shpActivePic(Main_Active).BackColor = m_Check_Active_Back
                shpActivePic(Main_Active).BorderColor = m_Check_Active_Border
            End If
        End If
        Set ThisSubmenu = Menus(Main_Active).SubMenu
        If Not ThisSubmenu Is Nothing And Menus_Active Then
            Set ThisSubmenu.Creator = Me
            ThisSubmenu.CreatorWidth = shpMenu(Main_Active).Width
            If m_Horizontal Then
                ThisSubmenu.Left = shpMenu(Main_Active).Left '+ 25
                ThisSubmenu.Top = shpMenu(Main_Active).Top + UserControl.Height - 50 'shpMenu(Main_Active).Height '- 25
                ThisSubmenu.ShowCover
            Else
                ThisSubmenu.Left = shpMenu(Main_Active).Left + shpMenu(Main_Active).Width - 25
                ThisSubmenu.Top = shpMenu(Main_Active).Top + 25
            End If
            ThisSubmenu.ZOrder 0
            ThisSubmenu.Main_Active = -1
            ThisSubmenu.Visible = True
            ThisSubmenu.RaisePopupMove
        ElseIf Not ThisSubmenu Is Nothing And Not Menus_Active Then
            ThisSubmenu.Visible = False
        End If
    End If
    
    On Error GoTo Unloaded_Menu
    For Each InactiveLabel In lblCaption
        If InactiveLabel.Index = 0 Or InactiveLabel.Index = -1 Then
            GoTo Continue
        End If
        If InactiveLabel.Index <> Main_Active Or (Not Menus_Active And Not m_Horizontal) Then
            InactiveLabel.ForeColor = m_Font_Inactive_Color
            If Not (Menus(InactiveLabel.Index).IsSeperator Or Menus(InactiveLabel.Index).Enabled) Then
                If m_Horizontal Then
                    shpMenu(InactiveLabel.Index).BackColor = m_Shape_Inactive_Color_Horizontal
                    shpMenu(InactiveLabel.Index).Height = 255
                Else
                    shpMenu(InactiveLabel.Index).BackColor = m_Shape_Inactive_Color
                End If
                shpMenu(InactiveLabel.Index).BorderStyle = 0
                shpMenu(InactiveLabel.Index).Visible = False
                If CBool(imgMenuBack(InactiveLabel.Index).Tag) Then
                    imgMenuBack(InactiveLabel.Index).Picture = imgMenuInactive.Picture
                    imgMenuBack(InactiveLabel.Index).Tag = False
                End If
                If Main_Active <> 0 Then
                    If Menus(InactiveLabel.Index).Checked Then
                        shpActivePic(InactiveLabel.Index).BackColor = m_Check_Inactive_Back
                        shpActivePic(InactiveLabel.Index).BorderColor = m_Check_InActive_Border
                    End If
                End If
            End If
            Set ThisSubmenu = Menus(InactiveLabel.Index).SubMenu
            If Not ThisSubmenu Is Nothing Then
                If ThisSubmenu.Visible Then
                    ThisSubmenu.Visible = False
                    ThisSubmenu.HideSubMenus
                End If
                On Error Resume Next
            End If
        End If

Unloaded_Menu:
Continue:
    Next InactiveLabel
End Sub

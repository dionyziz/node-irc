VERSION 5.00
Object = "{26AD3DAD-35EF-4D74-92B0-D106F68C32EC}#91.0#0"; "prjNodeMenu.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00925727&
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin NodeMenu.nMenu nmnuMain 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin NodeMenu.nMenu nmnuScript 
      Height          =   615
      Left            =   4320
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
   End
   Begin NodeMenu.nMenu nmnuDialogs 
      Height          =   615
      Index           =   0
      Left            =   5400
      TabIndex        =   1
      Top             =   2400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
   End
   Begin NodeMenu.nMenu nmnuFile 
      Height          =   735
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
   End
   Begin VB.Image Image2 
      Height          =   210
      Left            =   4080
      Picture         =   "test.frx":0000
      Top             =   3600
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1680
      Top             =   3000
      Width           =   330
   End
   Begin VB.Image imgIconOne 
      Height          =   240
      Left            =   1080
      Top             =   960
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Private Sub Form_Load()
    Me.Show
        
    nmnuMain.Horizontal = True
        nmnuFile(0).Horizontal = False
        nmnuFile(0).AddMenu "Σύνδεση..."
        nmnuFile(0).AddMenu "Disconnect"
        nmnuFile(0).AddMenu "-"
        nmnuFile(0).AddMenu "Options...", , Image2.Picture
        nmnuFile(0).AddMenu "-"
        nmnuFile(0).AddMenu "Exit"
        nmnuFile(0).EndMenu
    nmnuMain.AddMenu "File", nmnuFile(0)
        nmnuScript.Horizontal = False
        nmnuScript.AddMenu "Main"
            nmnuDialogs(0).Horizontal = False
            nmnuDialogs(0).AddMenu "Begin Dialog..."
            nmnuDialogs(0).AddMenu "Disconnect Dialog..."
            nmnuDialogs(0).EndMenu
        nmnuScript.AddMenu "Dialogs", nmnuDialogs(0)
        nmnuScript.AddMenu "-"
        nmnuScript.AddMenu "Scripts Folder"
        nmnuScript.AddMenu "Application Folder"
        nmnuScript.AddMenu "Sounds Folder"
        'nmnuScript.Popup = True
        nmnuScript.EndMenu
    nmnuMain.AddMenu "Script", nmnuScript
    nmnuMain.EndMenu
    
    nmnuFile(0).Color_Seperator_BackColor = 255
    nmnuFile(0).Color_Shape_Active_Color = 255
    nmnuFile(0).Color_Font_Inactive_Color = 255
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        With nmnuScript
            .Left = X
            .Top = Y
            .Menus_Active = True
            .Main_Active = 1
            .Visible = True
        End With
    End If
End Sub
Private Sub nmnuDialogs_PopupMove(Index As Integer)
    nmnuDialogs(Index).FixPosition nmnuDialogs(Index), nmnuScript
End Sub
Private Sub nmnuFile_MenuClick(Index As Integer, SubMenuIndex As Integer)
    Select Case SubMenuIndex
        Case 0 'connect
        Case 1 'disconnect
        Case 3 'options
            nmnuFile(Index).Checked(3) = Not nmnuFile(Index).Checked(3)
        Case 5 'exit
            End
    End Select
End Sub
Private Sub nmnuScript_PopupFinish()
    nmnuScript.Visible = False
End Sub

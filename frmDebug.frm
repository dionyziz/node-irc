VERSION 5.00
Begin VB.Form frmDebug 
   BackColor       =   &H00000000&
   Caption         =   "Debug Node"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   3495
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmDebug.frx":5F32
      Top             =   1560
      Width           =   6615
   End
End
Attribute VB_Name = "frmDebug"
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
Private MaxMode As Boolean
Private Sub Form_Load()
    txtDebug.Left = 0
    txtDebug.Top = 0
End Sub
Private Sub Form_Resize()
    Static PrivMaxMode As Boolean
    Static CodeCall As Boolean
    
    If CodeCall Then
        Exit Sub
    End If
    
    If MaxMode Then
        txtDebug.Width = Me.Width
        txtDebug.Height = Me.Height
    Else
        txtDebug.Width = Me.ScaleWidth
        txtDebug.Height = Me.ScaleHeight
    End If
    MaxMode = MaxMode Or WindowState = vbMaximized
    If MaxMode <> PrivMaxMode Then
        CodeCall = True
        PrivMaxMode = MaxMode
        If MaxMode Then
            BorderStyle = 0
            Caption = vbNullString
            If WindowState = vbMaximized Then
                Me.Visible = False
                WindowState = vbNormal
                Me.Visible = True
            End If
            Me.Left = 0
            Me.Top = 0
            Me.Width = Screen.Width
            Me.Height = Screen.Height
        Else
            Me.Left = Screen.Width \ 2
            Me.Width = Screen.Width \ 2
            Me.Top = 0
            Me.Height = Screen.Height
            BorderStyle = 2
            Caption = "Debug Node"
        End If
        CodeCall = False
    End If
End Sub
Private Sub txtDebug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If MaxMode = True Then
            MaxMode = False
            Form_Resize
        End If
    End If
End Sub

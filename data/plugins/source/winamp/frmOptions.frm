VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Winamp Plugin Options"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWinAmp 
      Interval        =   100
      Left            =   300
      Top             =   1380
   End
   Begin VB.CheckBox chkChannel 
      Caption         =   "Post active song to the Channels I am in"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5295
   End
   Begin VB.CheckBox chkStatus 
      Caption         =   "Display active song on my Status Window"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmOptions"
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
    LoadAll
End Sub
Private Sub cmdApply_Click()
    SaveAll
    LoadAll
End Sub
Private Sub cmdOK_Click()
    cmdApply_Click
    HideDialog
End Sub
Private Sub cmdCancel_Click()
    LoadAll
    HideDialog
End Sub
Private Sub HideDialog()
    Me.Hide
End Sub
Private Sub SaveAll()
     SaveSetting "Node.PlugIns", "Winamp", "ShowStatus", chkStatus.Value = vbChecked
     SaveSetting "Node.PlugIns", "Winamp", "PostChannel", chkChannel.Value = vbChecked
End Sub
Private Sub LoadAll()
    chkStatus.Value = IIf(GetSetting("Node.PlugIns", "Winamp", "ShowStatus", True), vbChecked, vbUnchecked)
    chkChannel.Value = IIf(GetSetting("Node.PlugIns", "Winamp", "PostChannel", True), vbChecked, vbUnchecked)
End Sub

Private Sub tmrWinAmp_Timer()
    WinampTimer
End Sub

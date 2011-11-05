VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   350
      Left            =   3840
      TabIndex        =   4
      Top             =   880
      Width           =   1150
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   540
      Width           =   2415
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   350
      Left            =   3840
      TabIndex        =   3
      Top             =   500
      Width           =   1150
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   1260
      Width           =   1150
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Height          =   350
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1150
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Search for"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   975
   End
End
Attribute VB_Name = "frmFind"
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
Public SearchString As String
Public SearchStart As Long
Public SearchCanceled As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub Form_Initialize()
    InitCommonControls
End Sub
Private Sub cmdCancel_Click()
    frmMain.SearchCanceled = True
    Me.Hide
End Sub
Private Sub cmdFind_Click()
    SearchCanceled = False
    If Not SearchString = txtFind Or SearchStart = -1 Then
        SearchString = txtFind
        SearchStart = 0
        frmMain.rtfText.Find SearchString, SearchStart
    Else
        SearchStart = frmMain.rtfText.Find(SearchString, SearchStart + Len(SearchString))
    End If
End Sub
Private Sub cmdReplace_Click()
    frmMain.rtfText.InsertString txtReplace.Text
    SearchStart = frmMain.rtfText.Find(SearchString, SearchStart + Len(txtReplace))
End Sub
Private Sub cmdReplaceAll_Click()
    Dim i As Integer
    i = frmMain.rtfText.Find(txtFind, 0)
    Do While i <> -1
        frmMain.rtfText.InsertString txtReplace.Text
        i = frmMain.rtfText.Find(txtFind, i + Len(txtReplace))
    Loop
End Sub


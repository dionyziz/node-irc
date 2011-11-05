VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTextReplacing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Text Styler"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   900
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   5280
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Top             =   2700
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Text to Find"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Text to Replace"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTextReplacing"
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

Private Sub cmdAdd_Click()
    Dim txtToCheck As String
    Dim txtToReplace As String
    txtToCheck = InputBox("Enter the string which you want to replace")
    If txtToCheck = "" Then Exit Sub
    txtToReplace = InputBox("With what do you want to replace '" & txtToCheck & "'")
    Dim temp As ListItem
    Set temp = ListView1.ListItems.Add(, , txtToCheck)
    temp.SubItems(1) = txtToReplace
End Sub

Private Sub cmdDelete_Click()
    ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
End Sub

Private Sub cmdEdit_Click()
    Dim txtToCheck As String
    Dim txtToReplace As String
    txtToCheck = InputBox("Enter the string which you want to replace", , ListView1.SelectedItem.Text)
    If txtToCheck = "" Then Exit Sub
    txtToReplace = InputBox("With what do you want to replace '" & txtToCheck & "'", , ListView1.SelectedItem.SubItems(1))
    ListView1.SelectedItem.Text = txtToCheck
    ListView1.SelectedItem.SubItems(1) = txtToReplace
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Resize()
    ListView1.ColumnHeaders(1).Width = (ListView1.Width - 300) / 3
    ListView1.ColumnHeaders(2).Width = ListView1.Width - 300 - ListView1.ColumnHeaders(1).Width
End Sub


VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Aliases"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CheckBox chkReplacing 
      Caption         =   "Use Aliases"
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Alias"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   4800
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Alias"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   4320
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Replace with this:"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Type This:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1935
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
Public Replacing As Boolean

Private Sub chkReplacing_Click()
    Replacing = chkReplacing.Value
End Sub

Private Sub Command1_Click()
    Dim newalias As String
    Dim newreplace As String
    newalias = InputBox("Enter what you would like to replace...", "Aliases")
    If newalias <> "" Then
        newreplace = InputBox("Enter what you would like to type for it...", "Aliases")
        If newreplace <> "" Then
            List1.AddItem newalias
            List2.AddItem newreplace
        End If
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            List1.RemoveItem i
            List2.RemoveItem i
        End If
    Next i
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) Then
            List1.RemoveItem i
            List2.RemoveItem i
        End If
    Next i
End Sub

Private Sub Command3_Click()
    Dim intFl As Integer
    Dim i As Integer
    
    intFl = FreeFile
    Open App.Path & "\..\..\conf\alias.dat" For Output As #intFl
        For i = 0 To List1.ListCount - 1
            Print #intFl, List1.List(i) & "," & List2.List(i)
        Next i
    Close #intFl
    Me.Hide
End Sub

Private Sub Command4_Click()
    Form_Load
    Me.Hide
End Sub

Private Sub Form_Load()
    Dim intFl As Integer
    Dim typetext As String
    Dim reptext As String
    
    List1.Clear
    List2.Clear
    
    intFl = FreeFile
    Open App.Path & "\..\..\conf\alias.dat" For Input As #intFl
        Do Until EOF(intFl)
            Line Input #intFl, typetext
            reptext = Strings.Right(typetext, Len(typetext) - InStr(1, typetext, ","))
            typetext = Replace(typetext, "," & reptext, "")
            frmOptions.List1.AddItem typetext
            frmOptions.List2.AddItem reptext
        Loop
    Close #intFl
End Sub


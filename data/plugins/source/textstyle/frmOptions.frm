VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Node Text Styler"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTextStyling 
      Caption         =   "Activate text styling"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Value           =   2  'Grayed
      Width           =   4395
   End
   Begin VB.CommandButton cmdReplaceList 
      Caption         =   "Replace List..."
      Height          =   350
      Left            =   3120
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox chkReplacing 
      Caption         =   "Activate Replacing"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   350
      Left            =   1620
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   350
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame frmTextStyling 
      Caption         =   "Type of style"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      Begin VB.OptionButton optLeet 
         Caption         =   "Leet"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   840
         Width           =   4275
      End
      Begin VB.OptionButton optCaps 
         Caption         =   "Partial Caps"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   4275
      End
      Begin VB.OptionButton optAscii 
         Caption         =   "Ascii Text"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   4275
      End
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
Public SelectedStyle As Integer
Public Replacing As Boolean

Private Sub chkReplacing_Click()
    Replacing = chkReplacing.Value
End Sub

Private Sub chkTextStyling_Click()
    frmTextStyling.Enabled = False Xor chkTextStyling.Value
End Sub

Private Sub cmdReplaceList_Click()
    frmTextReplacing.Show
End Sub

Private Sub Form_Load()
    LoadAll
    frmTextStyling.Enabled = False Xor chkTextStyling.Value
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
    SaveSetting "Node.PlugIns", "Text Styler", "StylingEnabled", chkTextStyling.Value = vbChecked
    SaveSetting "Node.PlugIns", "Text Styler", "Style", SelectedStyle
    SaveSetting "Node.PlugIns", "Text Styler", "ReplacingEnabled", chkReplacing.Value = vbChecked
    Dim temp1 As String
    Dim temp2 As String
    Dim i As Integer
    For i = 1 To frmTextReplacing.ListView1.ListItems.Count
        temp1 = temp1 & frmTextReplacing.ListView1.ListItems(i).Text & ","
        temp2 = temp2 & frmTextReplacing.ListView1.ListItems(i).SubItems(1) & ","
    Next i
    temp1 = Left(temp1, Len(temp1) - 1)
    temp2 = Left(temp2, Len(temp2) - 1)
    SaveSetting "Node.PlugIns", "Text Styler", "ReplacingTxtToChk", temp1
    SaveSetting "Node.PlugIns", "Text Styler", "ReplacingTxtToRplc", temp2
End Sub
Private Sub LoadAll()
    chkTextStyling.Value = IIf(GetSetting("Node.PlugIns", "Text Styler", "StylingEnabled", True), vbChecked, vbUnchecked)
    SelectedStyle = GetSetting("Node.PlugIns", "Text Styler", "Style", 1)
    Select Case SelectedStyle
        Case 1
            optAscii.Value = True
        Case 2
            optCaps.Value = True
        Case 3
            optLeet.Value = True
    End Select

    chkReplacing.Value = IIf(GetSetting("Node.PlugIns", "Text Styler", "ReplacingEnabled", True), vbChecked, vbUnchecked)

    Dim ReplacingTxtToChk() As String
    Dim ReplacingTxtToReplace() As String

    ReplacingTxtToChk = Split(GetSetting("Node.PlugIns", "Text Styler", "ReplacingTxtToChk", "lol,brb,wb,ty,np,bbl,ttyl,wtf,dont,cant,gn,i"), ",")
    ReplacingTxtToReplace = Split(GetSetting("Node.PlugIns", "Text Styler", "ReplacingTxtToRplc", "laughing out loud,be right back,welcome back,thank you,no problem,be back later,talk to you later,WHAT THE FUCK,don't,can't,good night,I"), ",")

    Dim i As Integer
    Dim temp As ListItem
    frmTextReplacing.ListView1.ListItems.Clear
    For i = 0 To UBound(ReplacingTxtToChk)
        Set temp = frmTextReplacing.ListView1.ListItems.Add(, , ReplacingTxtToChk(i))
        temp.SubItems(1) = ReplacingTxtToReplace(i)
    Next i
End Sub

Private Sub optAscii_Click()
    SelectedStyle = 1
End Sub

Private Sub optCaps_Click()
    SelectedStyle = 2
End Sub

Private Sub optLeet_Click()
    SelectedStyle = 3
End Sub

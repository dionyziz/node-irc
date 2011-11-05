VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{26AD3DAD-35EF-4D74-92B0-D106F68C32EC}#94.0#0"; "prjNodeMenu.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000010&
   Caption         =   "Node IRC Script Editor"
   ClientHeight    =   9915
   ClientLeft      =   1365
   ClientTop       =   1455
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   9915
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin sEdit.CodeEdit rtftext 
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2566
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   0
      LineNumbers     =   -1  'True
   End
   Begin VB.TextBox txtDesc 
      Height          =   9615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   4080
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0076
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0102
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0192
            Key             =   "Search"
         EndProperty
      EndProperty
   End
   Begin NodeMenu.nMenu nmnuFile 
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuMain 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
   End
   Begin MSComctlLib.ImageList imglOBrowser 
      Left            =   4800
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":020D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":05A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0941
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvObjects 
      Height          =   3255
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imglOBrowser"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComDlg.CommonDialog cdFile 
      Left            =   4200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin NodeMenu.nMenu nmnuView 
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuEdit 
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
   End
End
Attribute VB_Name = "frmMain"
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

Private Changed As Boolean
'Init XP Controls
'used to make the program display XP-style controls
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public SearchString As String
Public SearchStart As Long
Public SearchCanceled As Boolean
Private ObjectDescriptions() As String
Private Sub Form_Initialize()
    InitCommonControls
    If App.PrevInstance Then
        If MsgBox("Another instance of Node Script Editor is currently running. Would you like to start another instance?", vbYesNo Or vbQuestion, "Script Editor already running") = vbNo Then
            End
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim FS As FileSystemObject
    Dim strCommand As String
    
    Set FS = New FileSystemObject
    strCommand = Command
    If Left(strCommand, 1) = """" And Right(strCommand, 1) = """" Then
        strCommand = Mid(strCommand, 2, Len(strCommand) - 2)
    End If
    If strCommand <> "" Then
        If FS.FileExists(strCommand) Then
            cdFile.FileName = strCommand
            rtftext.LoadFile strCommand
        End If
    End If
    Changed = False
    LoadMenus
    
    'load the classes using node.exe TypeLib
    mdlObjList.BuildFromTypeLib

    'load the rest of the object tree from XML
    xmlLoadObjects
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    'too small size
    On Error Resume Next
    rtftext.Top = nmnuMain.Height
    rtftext.Width = Me.ScaleWidth - 120
    rtftext.Height = Me.ScaleHeight - rtftext.Top
    tvObjects.Top = nmnuMain.Height
    tvObjects.Left = Me.Width / 2
    tvObjects.Width = Me.Width / 2
    tvObjects.Height = Me.ScaleHeight - tvObjects.Top
    txtDesc.Top = nmnuMain.Height
    txtDesc.Width = Me.Width / 2
    txtDesc.Height = Me.ScaleHeight - txtDesc.Top
    
    nmnuMain.Width = Me.ScaleWidth
End Sub
Private Sub LoadMenus()
    nmnuMain.Horizontal = True
        nmnuFile.Horizontal = False
        nmnuFile.AddMenu "New", , ilMenu.ListImages(1).Picture
        nmnuFile.AddMenu "Open...", , ilMenu.ListImages(2).Picture
        nmnuFile.AddMenu "Save", , ilMenu.ListImages(3).Picture
        nmnuFile.AddMenu "Save As..."
        nmnuFile.AddMenu "-"
        nmnuFile.AddMenu "Exit"
        nmnuFile.EndMenu
        nmnuFile.ZOrder 0
    nmnuMain.AddMenu "File", nmnuFile
        nmnuView.Horizontal = False
        nmnuView.AddMenu "Editor"
        nmnuView.AddMenu "Object Browser"
        nmnuView.Checked(0) = True
        nmnuView.EndMenu
        nmnuView.ZOrder 0
    nmnuMain.AddMenu "View", nmnuView
        nmnuEdit.Horizontal = False
        nmnuEdit.AddMenu "Find and Replace...", , ilMenu.ListImages(4).Picture
        nmnuEdit.EndMenu
        nmnuEdit.ZOrder 1
    nmnuMain.AddMenu "Edit", nmnuEdit
    nmnuMain.EndMenu
    nmnuMain.ZOrder 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim SaveChanges As VbMsgBoxResult
    
    If Changed Then
        SaveChanges = MsgBox("The current code has been changed. Do you want to save these changes?", vbYesNoCancel Or vbQuestion)
        If SaveChanges = vbYes Then
            nmnuFile_MenuClick 2 'save
        ElseIf SaveChanges = vbCancel Then
            Cancel = True
        End If
    End If
    If Cancel = False Then
        Unload frmFind
        End
    End If
End Sub
Private Sub nMnuEdit_MenuClick(SubMenuIndex As Integer)
    Select Case SubMenuIndex
        Case 0 'find
            frmFind.Show 'vbModal
    End Select
End Sub
Private Sub nmnuFile_MenuClick(SubMenuIndex As Integer)
    Dim SaveChanges As VbMsgBoxResult
    Select Case SubMenuIndex
        Case 0 'new
            If Changed Then
                SaveChanges = MsgBox("The current code has been changed. Do you want to save these changes?", vbYesNoCancel Or vbQuestion)
                If SaveChanges = vbYes Then
                    nmnuFile_MenuClick 2 'save
                ElseIf SaveChanges = vbCancel Then
                    Exit Sub
                End If
            End If
            cdFile.FileName = ""
            rtftext.Text = ""
        Case 1 'open
            If Changed Then
                SaveChanges = MsgBox("The current code has been changed. Do you want to save these changes?", vbYesNoCancel Or vbQuestion)
                If SaveChanges = vbYes Then
                    nmnuFile_MenuClick 2 'save
                ElseIf SaveChanges = vbCancel Then
                    Exit Sub
                End If
            End If
            On Error GoTo Canceled
            cdFile.ShowOpen
            rtftext.LoadFile cdFile.FileName
            Changed = False
        Case 2 'save
            If cdFile.FileName = "" Then
                nmnuFile_MenuClick 3 'save as
            Else
                rtftext.SaveFile cdFile.FileName
            End If
            Changed = False
        Case 3 'save as
            On Error GoTo Canceled
            cdFile.ShowSave
            nmnuFile_MenuClick 2 'save
        Case 5 'exit
            'why here and in form_unload?
            'If Changed Then
            '    SaveChanges = MsgBox("The current code has been changed. Do you want to save these changes?", vbYesNoCancel Or vbQuestion)
            '    If SaveChanges = vbYes Then
            '        nmnuFile_MenuClick 2 'save
            '    ElseIf SaveChanges = vbCancel Then
            '        Exit Sub
            '    End If
            'End If
            Unload Me
    End Select
Canceled:
End Sub
Private Sub nmnuView_MenuClick(SubMenuIndex As Integer)
    Select Case SubMenuIndex
    Case 0 'editor
        rtftext.Visible = True
        tvObjects.Visible = False
        txtDesc.Visible = False
        nmnuView.Checked(0) = True
        nmnuView.Checked(1) = False
    Case 1 'object browser
        rtftext.Visible = False
        tvObjects.Visible = True
        txtDesc.Visible = True
        nmnuView.Checked(0) = False
        nmnuView.Checked(1) = True
    End Select
End Sub
Private Sub rtftext_Change()
    Changed = True
End Sub
Private Sub xmlLoadObjects()
    'New Function Objects are now in XML
        
    Dim XMLDoc As MSXML2.DOMDocument
    Dim XMLNode As MSXML2.IXMLDOMElement
    Dim MyNode As MSComctlLib.Node
    
    ReDim ObjectDescriptions(0)
    
    Set XMLDoc = New MSXML2.DOMDocument
    
    XMLDoc.async = True
    If Not XMLDoc.Load(App.Path & "\Objekts.xml") Then
        MsgBox "Failed to load object list from XML file.", vbCritical
        Exit Sub
    End If
    
    Set MyNode = tvObjects.Nodes.Item(1)
    
    For Each XMLNode In XMLDoc.documentElement.childNodes
        AddObjects XMLNode, MyNode
    Next XMLNode
End Sub
Public Sub AddObjects(ByRef ObjectGroup As IXMLDOMNode, ByRef ParentNode As MSComctlLib.Node)
    Dim ChildObject As IXMLDOMNode
    Dim MyNode As MSComctlLib.Node
    Dim intIconIndex As Integer
    Dim strName As String
    Dim strDescription As String
    Dim i As Integer
    
    For i = 0 To ObjectGroup.Attributes.Length - 1
        Select Case ObjectGroup.Attributes.Item(i).nodeName
            Case "Icon"
                On Error GoTo Invalid_Icon
                intIconIndex = ObjectGroup.Attributes.Item(i).nodeValue
                On Error GoTo 0
            Case "Name"
                strName = ObjectGroup.Attributes.Item(i).nodeValue
            Case "Description"
                strDescription = ObjectGroup.Attributes.Item(i).nodeValue
        End Select
    Next i
    
    Set MyNode = tvObjects.Nodes.Add(ParentNode, tvwChild, , strName, intIconIndex)
    
    ReDim Preserve ObjectDescriptions(MyNode.index)
    ObjectDescriptions(MyNode.index) = strDescription
    
    For Each ChildObject In ObjectGroup.childNodes
        AddObjects ChildObject, MyNode
    Next ChildObject
    
    Exit Sub
Invalid_Icon:
    MsgBox "The XML element " & ObjectGroup.nodeName & " has an invalid icon index."
End Sub
Private Sub tvObjects_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    txtDesc.Text = ObjectDescriptions(Node.index)
End Sub

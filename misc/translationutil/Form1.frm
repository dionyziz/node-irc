VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private langcharset As String
Private Sub LoadLanguage(ByVal FileName As String)
    'This sub is used to load the FileName language-file into the array Language
    Dim intFL As Integer 'file index
    Dim strCurrentLine As String 'the current line of the file
    Dim strStatement As String
    Dim i As Integer 'the key index for each language caption
    Dim objFile As File
    Dim objTextStream As TextStream
    Dim strLanguage As String
    Dim strTemp As String

    'get a free file index
    intFL = FreeFile
    'open the language file
    'Set objFile = fs.GetFile(FileName)
    'Set objTextStream = objFile.OpenAsTextStream(ForReading)
    strLanguage = DecodeFile(FileName, False)
    'Open FileName For Binary Access Read Shared As #intFl
    'skip the first two lines
    'strCurrentLine = objTextStream.ReadLine & objTextStream.ReadLine
    strCurrentLine = DecodedLineInput(strLanguage) & DecodedLineInput(strLanguage)
    i = 2
    
    'go through the file
    Do Until Len(strLanguage) = 0
        i = i + 1
        
        'get a line from the file(either a key or a comment)
        strCurrentLine = DecodedLineInput(strLanguage)
        'strCurrentLine = objTextStream.ReadLine
        'if it has a leading escape, remove it
        If Strings.Left(strCurrentLine, 1) = ">" Then
            strCurrentLine = Strings.Right(strCurrentLine, Len(strCurrentLine) - 1)
        End If
        'if it's not a comment add the current key to the array
        '                                        & Chr(0)
        If Strings.Left(strCurrentLine, 1) <> "#" And Len(strCurrentLine) <> 0 Then
            'this line is not a comment
            'error trap for language file bugs
            On Error GoTo Language_File_Error
            'key indexes can't be unicode
            strStatement = GetStatement(strCurrentLine) ', vbFromUnicode)
            If IsNumeric(strStatement) Then
                'TO DO
                'Language(strStatement) = GetParameter(StrConv(strCurrentLine, vbFromUnicode))
                'Language(strStatement) = GetParameter(strCurrentLine, , , True)
                List1.AddItem Replace(GetParameter(strCurrentLine, , , False), Chr(0), "")
            Else
                If strStatement = "@" Then
                    'setting char-set
                    'the charset index can't be unicode: it's a simple number
                    langcharset = GetParameter(strCurrentLine) 'vbFromUnicode))
                End If
            End If
        End If
    Loop
EOF_Return:
    'objTextStream.Close
    Close #intFL
    Exit Sub
Language_File_Error:
    If Err.Number = 62 Then
        'EOF
        Resume EOF_Return
    End If
    MsgBox "The language file " & FileName & " has an error in line " & i, vbCritical, "Language File Error"
End Sub
Private Sub Form_Load()
    MsgBox App.Path
    LoadLanguage App.Path & "\..\..\data\languages\english.lang"
End Sub

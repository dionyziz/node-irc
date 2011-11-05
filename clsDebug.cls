VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

#Const INSTANT_DEBUG = False

Option Explicit
Private frmDB As frmDebug
Private intFL As Integer
Private intFL2 As Integer
Private intIndentLevel As Integer
Private Last_t As Long
Private UniqueHTMLid As Long
Private Sub Class_Initialize()
    On Error Resume Next
    Kill App.Path & "\logs\debug.dat.html"
    StartLogging
    'Xhtm "<html>"
    'Xhtm "<head>"
    'Xhtm "<title>Node Debug Log</title>"
    'Xhtm "<script language=""javascript"" src=""" & App.Path & "/data/html/debug.js""> </script>"
    'Xhtm "</head>"
    'Xhtm "<body>"
    InitDebugWindow
    X App.EXEName & " " & VERSION_CODENAME & " " & App.Major & "." & App.Minor
    X "Logging Started"
End Sub
Public Sub StartLogging()
    intFL = FreeFile
    Open App.Path & "\logs\debug.dat" For Append As #intFL
    'intFL2 = FreeFile
    'Open App.Path & "\logs\debug.dat.html" For Append As #intFL2
End Sub
Public Sub StopLogging()
    Close #intFL
    Close #intFL2
End Sub
Public Sub SaveLogs()
    StopLogging
    StartLogging
End Sub
Public Sub InitDebugWindow()
    Set frmDB = New frmDebug
    frmDB.txtDebug.Text = vbNullString
    
    #If INSTANT_DEBUG Then
        frmDB.Show
    #End If
    
    frmDB.Left = Screen.Width \ 2
    frmDB.Top = 0
    frmDB.Width = Screen.Width \ 2
    frmDB.Height = Screen.Height
End Sub
Public Sub ShowDebugWindow()
    frmDB.Show
End Sub
Public Sub X(ByVal CodeLine As String, Optional DirectHTM As Boolean = True)
    'line of code executing
    Print #intFL, String$(intIndentLevel, vbTab) & CodeLine
    DebugWindowAddLog String$(intIndentLevel, vbTab) & CodeLine
    'If DirectHTM Then
    '    Xhtm XString & CodeLine & "<br>"
    'End If
    SaveLogs
End Sub
Public Function XString() As String
    Dim strIndent As String
    Dim i As Integer
    
    For i = 0 To intIndentLevel
        strIndent = strIndent & "&nbsp;&nbsp;&nbsp;&nbsp;"
    Next i
    XString = strIndent
End Function
'Public Sub Xhtm(ByVal HTMLCode As String)
'    Print #intFL2, HTMLCode & vbnewline
'End Sub
Public Sub Enter(ByVal ProcedureName As String, Optional ByVal Reason As String)
    'entering procedure
    UniqueHTMLid = UniqueHTMLid + 1
    X " -> Entering " & ProcedureName & IIf(Len(Reason) > 0, " (" & Reason & ")", vbNullString), False
    'Xhtm XString
    'Xhtm "<a href=""JavaScript:expand('debug_div_" & UniqueHTMLid & "');""> + </a>"
    'Xhtm "<font color=""navyblue""><i>" & ProcedureName & "</i></font>"
    'Xhtm "<div id=""debug_div_" & UniqueHTMLid & """ style=""display:none"">"
    intIndentLevel = intIndentLevel + 1
    Last_t = GetTickCount
End Sub
Public Sub Leave(ByVal ProcedureName As String, Optional ByVal Reason As String)
    'leaving procedure
    intIndentLevel = intIndentLevel - 1
    X " <- Leaving " & ProcedureName & IIf(Len(Reason) > 0, " (" & Reason & ")", vbNullString), False
    'Xhtm "</div><br>"
    X "(Since Last Enter: " & GetTickCount - Last_t & " ms)"
End Sub
Public Sub XWarning(ByVal Warning As String)
    X "Warning: " & Warning, False
    'Xhtm "<font color=""red""><strong>Warning:" & Warning & "</strong></font><br>"
End Sub
Public Sub DebugWindowAddLog(ByVal strText As String)
    If Len(frmDB.txtDebug.Text) > 10000 Then
        frmDB.txtDebug.Text = vbNullString
    End If
    frmDB.txtDebug.Text = frmDB.txtDebug.Text & strText & vbNewLine
    frmDB.txtDebug.SelStart = Len(frmDB.txtDebug.Text)
End Sub
Private Sub Class_Terminate()
    X "Logging Stopped"
    'Xhtm "</body></html>"
    StopLogging
End Sub

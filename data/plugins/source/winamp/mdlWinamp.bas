Attribute VB_Name = "mdlWinamp"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Public Const HTML_OPEN As String = "" 'HtmlOpen = Chr(27)
Public Const HTML_CLOSE As String = "" 'HtmlOpen = Chr(29)
Public Const HTML_BR As String = HTML_OPEN & "br" & HTML_CLOSE

Public objNode As prjNPIN.clsHostInterface
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private oldWinampSong As String

Function GetWindowTitle(TheHWnd As Long) As String
    Dim Title As String
    If IsWindow(TheHWnd) Then
        Title = Space$(GetWindowTextLength(TheHWnd) + 1)
        Call GetWindowText(TheHWnd, Title, Len(Title))
        Title = Left$(Title, Len(Title) - 1)
    End If
    GetWindowTitle = Title
End Function

Public Sub WinampTimer()
    Dim ThisWinampSong As String
    ThisWinampSong = Replace(GetWindowTitle(FindWindow("Winamp v1.x", vbNullString)), " - Winamp", "")
    If ThisWinampSong <> oldWinampSong Then
        If frmOptions.chkStatus.Value = 1 Then
            objNode.npinAddStatus HTML_OPEN & "b" & HTML_CLOSE & "You're listening to " & ThisWinampSong & HTML_OPEN & "/b" & HTML_CLOSE & vbCrLf
        End If
        If frmOptions.chkChannel.Value = 1 Then
            objNode.TxtSend "/ame is listening to """ & ThisWinampSong & """"
        End If
        oldWinampSong = ThisWinampSong
    End If
End Sub
Sub Main()
    'MsgBox "Loading Winamp Plugin", vbInformation, "Winamp Plugin"
End Sub


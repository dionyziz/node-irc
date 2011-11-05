VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugIn"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Private objNode As prjNPIN.clsHostInterface
Implements prjNPIN.clsLittleFinger
Private Sub clsLittleFinger_IRCAction(ByVal ActionID As prjNPIN.NodePlugInsNotifications)
    'MsgBox "Action Event " & ActionID & " fired!", vbInformation, "Node Action Event"
End Sub
Private Function clsLittleFinger_PluginInit(objHostInterface As prjNPIN.clsHostInterface) As Boolean
    Set objNode = objHostInterface
    'MsgBox "PlugIn Initialized", vbInformation
    frmOptions.optAscii.Caption = MakeAscii(frmOptions.optAscii.Caption)
    frmOptions.optCaps.Caption = MakePartialCaps(frmOptions.optCaps.Caption)
    frmOptions.optLeet.Caption = MakeLeet(frmOptions.optLeet.Caption)
End Function
Private Function clsLittleFinger_PluginOptions() As Boolean
    frmOptions.Show
End Function
Private Sub clsLittleFinger_Sending()
    Dim Text As String
    'get the text
    Text = objNode.GetTxtSendText
    
    If frmOptions.Replacing = True Then
        'replace some strings
        Dim i As Integer
        For i = 1 To frmTextReplacing.ListView1.ListItems.Count
            Text = ReplaceOnlyComplete(Text, frmTextReplacing.ListView1.ListItems(i).Text, frmTextReplacing.ListView1.ListItems(i).SubItems(1))
        Next i
    End If
    
    'if the 1st chararcter a / -> exit sub
    If Left(Text, 1) = "/" Then
        objNode.SetTxtSendText Text
        Exit Sub
    End If
    
    If frmOptions.chkTextStyling.Value = 1 Then
    Select Case frmOptions.SelectedStyle
    Case 1
        Text = MakeAscii(Text)
    Case 2
        Text = MakePartialCaps(Text)
    Case 3
        Text = MakeLeet(Text)
    End Select
    End If
    
    objNode.SetTxtSendText Text
End Sub
Private Function MakeAscii(Text As String) As String
    Text = UCase(Text)
    Text = Replace(Text, "AE", "Æ")
    Text = Replace(Text, "A", "å")
    Text = Replace(Text, "B", "ß")
    Text = Replace(Text, "C", "Ç")
    Text = Replace(Text, "D", "Ð")
    Text = Replace(Text, "E", "Ë")
    Text = Replace(Text, "F", "f")
    Text = Replace(Text, "G", "g")
    Text = Replace(Text, "H", "H")
    Text = Replace(Text, "I", "î")
    Text = Replace(Text, "J", "J")
    Text = Replace(Text, "K", "k")
    Text = Replace(Text, "L", "|")
    Text = Replace(Text, "M", "m")
    Text = Replace(Text, "N", "ñ")
    Text = Replace(Text, "O", "ö")
    Text = Replace(Text, "P", "þ")
    Text = Replace(Text, "Q", "¶")
    Text = Replace(Text, "R", "®")
    Text = Replace(Text, "S", "§")
    Text = Replace(Text, "T", "t")
    Text = Replace(Text, "U", "ü")
    Text = Replace(Text, "V", "V")
    Text = Replace(Text, "W", "w")
    Text = Replace(Text, "X", "×")
    Text = Replace(Text, "Y", "¥")
    Text = Replace(Text, "Z", "z")
    Text = Replace(Text, "1", "¹")
    Text = Replace(Text, "2", "²")
    Text = Replace(Text, "3", "³")
    MakeAscii = Text
End Function
Private Function MakePartialCaps(Text As String)
    Text = Replace(UCase(Text), "A", "a")
    Text = Replace(Text, "E", "e")
    Text = Replace(Text, "I", "i")
    Text = Replace(Text, "O", "o")
    Text = Replace(Text, "U", "u")
    MakePartialCaps = Text
End Function
Private Function MakeLeet(Text As String)
    Text = Replace(LCase(Text), "a", "4")
    Text = Replace(Text, "b", "8")
    Text = Replace(Text, "e", "3")
    Text = Replace(Text, "i", "1")
    Text = Replace(Text, "o", "0")
    Text = Replace(Text, "t", "7")
    Text = Replace(Text, "g", "6")
    Text = Replace(Text, "s", "5")
    MakeLeet = Text
End Function
Private Function ReplaceOnlyComplete(strExpression As String, strFind As String, strReplace As String) As String
    'the only text
    If Len(strExpression) = Len(strFind) Then
        ReplaceOnlyComplete = Replace(strExpression, strFind, strReplace)
        Exit Function
    End If
    
    'at the end
    If Right(strExpression, Len(strFind) + 1) = " " & strFind Then
        ReplaceOnlyComplete = Left(strExpression, Len(strExpression) - Len(strFind)) & strReplace
        Exit Function
    End If
    
    'at the beginning
    If Left(strExpression, Len(strFind & " ")) = strFind & " " Then
        ReplaceOnlyComplete = strReplace & Right(strExpression, Len(strExpression) - Len(strFind))
        Exit Function
    End If
    
    ReplaceOnlyComplete = Replace(strExpression, " " & strFind & " ", " " & strReplace & " ")
End Function
Private Sub clsLittleFinger_SomeoneJoined(ByVal UserName As String, ByVal Channel As String)
    '--
End Sub
Private Sub clsLittleFinger_Receiving(ByVal User As String, ByVal Text As String, Optional ByVal Channel As String)
    '--
End Sub


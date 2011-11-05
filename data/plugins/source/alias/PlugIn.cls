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
End Function
Private Function clsLittleFinger_PluginOptions() As Boolean
    frmOptions.Show
End Function
Private Sub clsLittleFinger_Sending()
    Dim Text As String, text2 As String
    Dim i As Integer, ii As Integer
    Dim allwords() As String
    
    'get the text
    Text = objNode.GetTxtSendText
    If Text <> "" Then
        If frmOptions.Replacing = True Then
            'replace some strings
            ReDim allwords(0)
            allwords() = Split(Text, " ")
            For i = 0 To UBound(allwords)
                For ii = 0 To frmOptions.List1.ListCount - 1
                    If allwords(i) = frmOptions.List2.List(ii) Then
                        allwords(i) = frmOptions.List1.List(ii)
                    End If
                Next ii
            Next i
        Else
            objNode.SetTxtSendText Text
            Exit Sub
        End If
    
        For i = 0 To UBound(allwords)
            text2 = text2 & allwords(i) & " "
        Next i
        text2 = Strings.Left(text2, Len(text2) - 1)
        objNode.SetTxtSendText text2
    End If
End Sub
Private Sub clsLittleFinger_Receiving(ByVal User As String, ByVal Text As String, Optional ByVal Channel As String)
    '--
End Sub
Private Sub clsLittleFinger_SomeoneJoined(ByVal UserName As String, ByVal Channel As String)
    '--
End Sub

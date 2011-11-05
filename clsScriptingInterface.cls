VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNPINInterface"
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
Implements prjNPIN.clsHostInterface

Private Function clsHostInterface_GetTxtSendText() As String
    clsHostInterface_GetTxtSendText = frmMain.txtSend.Text
End Function

Private Property Get clsHostInterface_MyNick() As String
    clsHostInterface_MyNick = CurrentActiveServer.myNick
End Property

Private Sub clsHostInterface_npinAddStatus(ByVal strMessage As String, Optional ByVal ChannelId As Integer)
    frmMain.AddStatus strMessage, CurrentActiveServer, ChannelId
End Sub

Private Function clsHostInterface_preExecute(ByVal strIRC As String) As Variant
    'TO DO:
    'Give mulipleserver access to the PlugIns
    CurrentActiveServer.preExecute strIRC
End Function

Private Function clsHostInterface_SetTxtSendText(ByVal strText As String) As Variant
    frmMain.txtSend.Text = strText
End Function

Private Function clsHostInterface_xLanguage(ByVal intLanguageKey As Integer) As Variant
    clsHostInterface_xLanguage = Language(intLanguageKey)
End Function

Private Function clsHostInterface_TxtSend(ByVal strText As String) As Variant
    Dim i As Integer
    CurrentActiveServer.preExecute strText
End Function

Private Function clsHostInterface_GetMsgRCV() As String
    clsHostInterface_GetMsgRCV = frmMain.MsgRcvTxt
End Function

Private Function clsHostInterface_xMultiLanguage(ByVal LangName As String, ByVal intLanguageKey As Integer) As Variant
    Dim i
    For i = 0 To UBound(LangMultiNames)
        If LCase$(LangMultiNames(i)) = LCase$(LangName) Then
            clsHostInterface_xMultiLanguage = LanguageMulti(i).Cell(1, intLanguageKey)
            GoTo foundlang
        End If
    Next i
    LoadMultiLanguage LangName
    clsHostInterface_xMultiLanguage = LanguageMulti(UBound(LangMultiNames)).Cell(1, intLanguageKey)
foundlang:
End Function


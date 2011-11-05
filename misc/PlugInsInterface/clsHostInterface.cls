VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHostInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Public Sub npinAddStatus(ByVal strMessage As String, Optional ByVal ChannelID As Integer)
End Sub
Public Function xLanguage(ByVal intLanguageKey As Integer)
End Function
Public Function preExecute(ByVal strIRC As String)
End Function
Public Function GetTxtSendText() As String
End Function
Public Function SetTxtSendText(ByVal strText As String)
End Function
Public Function TxtSend(ByVal strText As String)
End Function
Public Function GetMsgRCV() As String
End Function
Public Function xMultiLanguage(ByVal LangName As String, ByVal intLanguageKey As Integer)
End Function
Public Property Get MyNick() As String
    'get local nickname
End Property

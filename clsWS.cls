VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWS"
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
Public WithEvents WS As Winsock
Attribute WS.VB_VarHelpID = -1
Public WSCollection_Parent As clsWSArray
Public MyIndex As Integer
Private Sub Class_Initialize()
    On Error GoTo Could_Not_Load_WinSock
    Set WS = New Winsock
    Exit Sub
Could_Not_Load_WinSock:
    DB.X "Warning: WinSock object could not be created at clsWs.Class_Initialize()"
End Sub
Private Sub WS_Close()
    WSCollection_Parent.CallBack_Close MyIndex
End Sub
Private Sub WS_Connect()
    WSCollection_Parent.CallBack_Connect MyIndex
End Sub
Private Sub WS_ConnectionRequest(ByVal requestID As Long)
    WSCollection_Parent.CallBack_ConnectionRequest MyIndex, requestID
End Sub
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
    WSCollection_Parent.CallBack_DataArrival MyIndex, bytesTotal
End Sub
Private Sub WS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    WSCollection_Parent.CallBack_Error MyIndex, Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay
End Sub
Private Sub WS_SendComplete()
    WSCollection_Parent.CallBack_SendComplete MyIndex
End Sub
Private Sub WS_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    WSCollection_Parent.CallBack_SendProgress MyIndex, bytesSent, bytesRemaining
End Sub

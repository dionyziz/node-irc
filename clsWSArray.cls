VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWSArray"
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

Public Event Connect(ByVal Index As Integer)
Public Event ConnectionClosed(ByVal Index As Integer)
Public Event ConnectionRequest(ByVal Index As Integer, ByVal requestID As Long)
Public Event DataArrival(ByVal Index As Integer, ByVal bytesTotal As Long)
Public Event Error(ByVal Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete(ByVal Index As Integer)
Public Event SendProgress(ByVal Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

Public WSCollection As Collection
Private Sub Class_Initialize()
    Set WSCollection = New Collection
End Sub
Public Property Get Item(ByVal Index As Integer) As Winsock
    'We need to add 1 here, as the Collection object
    'is one-based, while we want clsWSArray to be zero-based
    On Error GoTo eh_bad_glue
    Set Item = WSCollection.Item(Index + 1).WS
    Exit Property
eh_bad_glue:
    DB.X "Warning: Wrong call to clsWSArray.Item(). Index " & Index & " is not accessible."
End Property
Public Function LoadNew() As Integer
    Dim NewItem As clsWS
    
    Set NewItem = New clsWS
    Set NewItem.WSCollection_Parent = Me
    NewItem.MyIndex = WSCollection.Count
    
    WSCollection.Add NewItem
    
    LoadNew = NewItem.MyIndex
End Function
Public Sub Remove(ByVal Index As Integer)
    WSCollection.Remove Index + 1
End Sub
Public Property Get Count() As Integer
    Count = WSCollection.Count
End Property
Friend Sub CallBack_Connect(ByVal Index As Integer)
    RaiseEvent Connect(Index)
End Sub
Friend Sub CallBack_Close(ByVal Index As Integer)
    RaiseEvent ConnectionClosed(Index)
End Sub
Friend Sub CallBack_ConnectionRequest(ByVal Index As Integer, ByVal requestID As Long)
    RaiseEvent ConnectionRequest(Index, requestID)
End Sub
Friend Sub CallBack_DataArrival(ByVal Index As Integer, ByVal bytesTotal As Long)
    RaiseEvent DataArrival(Index, bytesTotal)
End Sub
Friend Sub CallBack_Error(ByVal Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RaiseEvent Error(Index, Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub
Friend Sub CallBack_SendComplete(ByVal Index As Integer)
    RaiseEvent SendComplete(Index)
End Sub
Friend Sub CallBack_SendProgress(ByVal Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    RaiseEvent SendProgress(Index, bytesSent, bytesRemaining)
End Sub

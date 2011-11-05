VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Name As String
Public isOnline As Boolean
Private Server() As String

Public Property Let Servers(ByVal index As Integer, servername As String)
    Server(index) = servername
End Property

Public Property Get Servers(ByVal index As Integer) As String
    Servers = Server(index)
End Property

Public Property Get ServerCount() As Integer
    ServerCount = UBound(Server)
End Property

Public Function AddServer(ByVal servername As String)
    ReDim Preserve Server(UBound(Server) + 1)
    Server(UBound(Server)) = servername
End Function

Private Sub Class_Initialize()
    ReDim Server(0)
    isOnline = False
End Sub

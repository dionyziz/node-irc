Attribute VB_Name = "mdlNetwork"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'
'The code in this module is based on the code provided by Richard Deeming, richard (at) trinet.co.uk
'

Option Explicit

Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Private Type HOSTENT
    hName      As Long
    hAliases   As Long
    hAddrType  As Integer
    hLen       As Integer
    hAddrList  As Long
End Type

Private Type WSADATA
    wVersion      As Integer
    wHighVersion  As Integer
    szDescription(0 To MAX_WSADescription)   As Byte
    szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
    wMaxSockets   As Integer
    wMaxUDPDG     As Integer
    dwVendorInfo  As Long
End Type

Private Declare Function gethostbyaddr Lib "wsock32.dll" (ByRef dwHost As Long, ByVal hLen As Integer, ByVal aType As Integer) As Long
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Function gethostname Lib "wsock32.dll" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32.dll" (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H1 And &HFF&
End Function
Private Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Private Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        App.LogEvent "Socket error occurred in Cleanup.", _
        vbLogEventTypeError
    End If
End Sub
Private Function SocketsInitialize() As Boolean
    Dim WSAD As WSADATA, sLoByte As String, sHiByte As String
        
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        DB.X "Warning: The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        DB.X "Warning: This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
    
        DB.X "Warning: Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
    
        SocketsInitialize = False
        Exit Function
    End If
    SocketsInitialize = True
End Function
Public Function GetIPAddress(Optional ByVal sHost As String, Optional serrmsg As String) As String
    'Resolves the host-name (or current machine if balnk) to an IP address
    Dim sHostName   As String * 256
    Dim lpHost      As Long
    Dim HOST        As HOSTENT
    Dim dwIPAddr    As Long
    Dim tmpIPAddr() As Byte
    Dim i           As Integer
    Dim sIPAddr     As String
    Dim werr        As Long
    
    If Not SocketsInitialize() Then
        GetIPAddress = vbNullString
        Exit Function
    End If
    
    If LenB(sHost) = 0 Then
        If gethostname(sHostName, 256) = SOCKET_ERROR Then
            werr = WSAGetLastError()
            GetIPAddress = vbNullString
            DB.X "Warning: Windows Sockets error " & Str(werr) & " occurred. Unable to successfully get Host Name."
            
            SocketsCleanup
            Exit Function
        End If
        sHostName = Trim$(sHostName)
    Else
        sHostName = Trim$(sHost) & ChrW$(0)
    End If
    
    lpHost = gethostbyname(sHostName)

    If lpHost = 0 Then
        werr = WSAGetLastError()
        GetIPAddress = vbNullString
        serrmsg = "Windows Sockets error " & Str$(werr) & _
                " has occurred. Unable to successfully get Host Name." & vbNewLine
        GetIPAddress = vbNullString
        
        SocketsCleanup
        Exit Function
    End If

    CopyMemory HOST, lpHost, Len(HOST)
    CopyMemory dwIPAddr, HOST.hAddrList, 4

    ReDim tmpIPAddr(1 To HOST.hLen)
    CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen

    For i = 1 To HOST.hLen
        sIPAddr = sIPAddr & tmpIPAddr(i) & "."
    Next

    GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)

    SocketsCleanup
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsActiveServer"
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

Private m_HostName As String
Private m_Description As String
Private m_ConnectTime As Long
Private m_Connected As Boolean
Private m_myStatus As NodeStatus
Private m_Port As Long

'while we are exchanging SOCKS5 messages we will need to know whether we have sent and recieved the method selection packages, and the connection packages
'Phase = 0 : Not Connected
'Phase = 1 : Connected, but nothing has been sent or recieved
'Phase = 2 : Sent version identifier/method selection message
'Phase = 3 : Recieved METHOD selection message
'Phase = 4 : Sent SOCKS request
'Phase = 5 : Recieved SOCKS reply
'Phase = 6 : SOCKS connection established - we can now send IRC data :-)
Private m_bProxyPhase As Byte

Private DCCFiles() As DCCFile 'the files being sent over DCC
Private DCCFilesRCV() As DCCFile 'the file being received over DCC
Private DCCChats() As DCCChatINFO 'the info for each DCC Chat session
Private txtStatus() As String 'buffer storing all the active conversations
Private ChannelList() As ndChannelListEntry 'the channel list
Private LastConnectionFailure As Long

Public NumberOfChannels As Long
Public ConnectRetry As Boolean
Public AltNickVal As Integer
Public IPSender As String 'ip of someone attempting to send a file over DCC
Public DCCTotal As Integer 'the total number of DCC attempts
Public intPreviousTabIndex As Integer
Public intCurrentlySelected As Integer
Public bCLSorting As Byte 'channel list sorting key
Public DoSessionPerform As Boolean
Public SessionPerformFile As String
Public TabsIndex As Integer 'the index of frmMain.tsTabs() we are using
Public ServerNode As ComctlLib.Node

Public strTempMemoParse As String
Public npmpData As nodePendingMemoParse

Public Enum nodePendingMemoParse
    nodeMemoParseNone
    nodeMemoParseRead
    nodeMemoParseList
    nodeMemoParseNoIdentify
    nodeMemoParseNoRegister
End Enum

Public Type nodeServerAttributes005
    Map As Boolean
    Knock As Boolean
    Safelist As Boolean
    HCN As Boolean
    WallChOps As Boolean
    Modes As Long
    ChanLimit As String
    Silence As Integer
    Watch As Integer
    MaxChannels As Long
    MaxBans As Integer
    MaxTargets As Integer
    AwayLen As Integer
    NickLen As Integer
    UserLen As Integer
    HostLen As Integer
    TopicLen As Integer
    KickLen As Integer
    ChannelLen As Integer
    KeyLen As Integer
    ChanTypes As String
    ChanModes As String
    PreFix As String
    CaseMapping As String
    Network As String
    IRCd As String
End Type

Private ServerAttributes As nodeServerAttributes005

'User's Current Nickname
Public myNick As String 'this variable stores the user's nickname and is updated when the nickname changes

'Nicks Array
Private m_NickList() As NodeNickList 'array that contains all the nicklists of the currently open channels

Public TabInfo As Collection 'collection containing the connection indeses between tabs and log-textboxes(for channels, privates and the status) or webbrowsers(for web site tabs)
Public TabType As Collection 'collection containing the type of every tab(channel, private, website or status)
Public Tabs As ComctlLib.TabStrip
Public WithEvents WinSockConnection As Winsock
Attribute WinSockConnection.VB_VarHelpID = -1
Public Function Self() As clsActiveServer
    Set Self = Me
End Function
Public Property Get SupportsHalfop() As Boolean
    SupportsHalfop = InStr(1, ServerAttributes.PreFix, "%") > 0
End Property
Public Property Get SupportsVoice() As Boolean
    SupportsVoice = InStr(1, ServerAttributes.PreFix, "+") > 0
End Property
Public Property Get SupportsOp() As Boolean
    SupportsOp = InStr(1, ServerAttributes.PreFix, "@") > 0
End Property
Private Sub Class_Initialize()
    Dim NewTabsIndex As Integer
    
    DB.Enter "clsActiveServer.Class_Initialize"
    
    DB.X "Loading WinSock object for this server"
    
    On Error GoTo Could_Not_Load_WinSock
    Set WinSockConnection = New Winsock

    DB.X "Initializing TabInfo and TabType"
    
    'create two new collections for tabinfo and tabtype
    Set TabInfo = New Collection
    Set TabType = New Collection
    
    NewTabsIndex = frmMain.tsTabs.Count
    
    DB.X "Loading new Tabs collection"
    Load frmMain.tsTabs(NewTabsIndex)
    Set Tabs = frmMain.tsTabs(NewTabsIndex)
    Tabs.ZOrder 0
    TabsIndex = NewTabsIndex
    Set Tabs.ImageList = frmMain.ilTabs

    DB.X "Preserving Memory"
    ReDim txtStatus(0)
    ReDim m_NickList(0)
    ReDim DCCFiles(0)
    ReDim DCCFilesRCV(0)
    ReDim DCCChats(0)

    'each nicklist object stores its current width in the Size property.
    NickList_Size(0) = NickList_DefaultSize
    
    DB.X "Creating Status Window"
    GetStatusID

    'set status to "not set"
    'fix: just set it to online
    m_myStatus = 0
    
    DB.Leave "clsActiveServer.Class_Initialize"
    Exit Sub
Could_Not_Load_WinSock:
    DB.X "Warning: Could not load WinSock for NewServer!"
    CriticalError
End Sub
Private Sub Class_Terminate()
    If WinSockConnection.State = sckConnected Then
        WinSockConnection.Close
    End If
    Set Tabs = Nothing
End Sub
Public Property Get ChannelList_Channel(ByVal index As Integer) As String
    ChannelList_Channel = ChannelList(index).Channel
End Property
Public Property Get ChannelList_Users(ByVal index As Integer) As Integer
    ChannelList_Users = ChannelList(index).Users
End Property
Public Property Get ChannelList_Topic(ByVal index As Integer) As String
    ChannelList_Topic = ChannelList(index).Topic
End Property
Public Property Let ChannelList_Channel(ByVal index As Integer, NewValue As String)
    ChannelList(index).Channel = NewValue
End Property
Public Property Let ChannelList_Users(ByVal index As Integer, NewValue As Integer)
    ChannelList(index).Users = NewValue
End Property
Public Property Let ChannelList_Topic(ByVal index As Integer, NewValue As String)
    ChannelList(index).Topic = NewValue
End Property
Public Sub ChannelList_ReDim(ByVal NewDimention As Integer, Optional ByVal bolPreserve As Boolean = True)
    If bolPreserve Then
        ReDim Preserve ChannelList(NewDimention)
    Else
        ReDim ChannelList(NewDimention)
    End If
End Sub
Public Function ChannelList_UBound()
    On Error Resume Next
    If UBound(ChannelList) And False Then
        ChannelList_UBound = 0
    Else
        ChannelList_UBound = UBound(ChannelList)
    End If
End Function
Public Property Get NickList_List(ByVal index As Integer) As Collection
    Set NickList_List = m_NickList(index).List
End Property
Public Property Get NickList_Size(ByVal index As Integer) As String
    NickList_Size = m_NickList(index).Size
End Property
Public Property Get NickList_Topic(ByVal index As Integer) As String
    NickList_Topic = m_NickList(index).Topic
End Property
Public Property Get NickList_Topic_Parsed(ByVal index As Integer) As String 'readonly
    NickList_Topic_Parsed = m_NickList(index).Topic_Parsed
End Property
Public Property Let NickList_List(ByVal index As Integer, ByRef NewValue As Collection)
    Set m_NickList(index).List = NewValue
End Property
Public Property Get NickList_Modes(ByVal index As Integer) As String
    NickList_Modes = m_NickList(index).Modes
End Property
Public Property Let NickList_Modes(ByVal index As Integer, ByRef NewValue As String)
    m_NickList(index).Modes = NewValue
End Property
Public Property Let NickList_Size(ByVal index As Integer, ByVal NewValue As String)
    m_NickList(index).Size = NewValue
End Property
Public Property Let NickList_Topic(ByVal index As Integer, ByVal NewValue As String)
    m_NickList(index).Topic = NewValue
    m_NickList(index).Topic_Parsed = CreateMainText(NewValue, Me, Options.UseSmileys, False)
End Property
Public Sub NickList_List_SetToNothing(ByVal index As Integer)
    Set m_NickList(index).List = Nothing
End Sub
Public Sub NickList_List_LoadNewCollection(ByVal index As Integer)
    Set m_NickList(index).List = New Collection
End Sub
Public Sub NickList_ReDim(ByVal NewDimention As Integer)
    ReDim Preserve m_NickList(NewDimention)
End Sub
Public Property Get MyStatus() As Integer
    MyStatus = m_myStatus
End Property
Public Property Let MyStatus(ByVal NewValue As Integer)
    Dim i As Integer
    
    m_myStatus = NewValue
    
    If WinSockConnection.State = sckConnected Then
        'set IRC status
        If NewValue = Status_Online Then
            preExecute "/away"
        Else
            preExecute "/away :" & Replace(Language(409), "%1", Language(409 + NewValue))
        End If
    End If
    'send status over NDC
    For i = 1 To UBound(NDCConnections)
        frmMain.NDCSendStatus i
    Next i
    'store your status
    SaveSetting App.EXEName, "Remember", "Status", CurrentActiveServer.MyStatus
End Property
Public Property Get bProxyPhase() As Byte
    bProxyPhase = m_bProxyPhase
End Property
Public Property Let bProxyPhase(ByVal NewValue As Byte)
    m_bProxyPhase = NewValue
End Property
Public Property Get HostName() As String
    HostName = m_HostName
End Property
Public Property Let HostName(NewValue As String)
    m_HostName = NewValue
End Property
Public Property Get Port() As Long
    Port = m_Port
End Property
Public Property Let Port(NewValue As Long)
    m_Port = NewValue
End Property
Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Get ConnectTime() As Long
    ConnectTime = m_ConnectTime
End Property
Public Property Get Connected() As Boolean
    Connected = m_Connected
End Property
Public Sub Parse005Attributes(ByVal strAttributes As String)
    Dim strAllAttributes() As String
    Dim strThisAttribute() As String
    Dim i As Integer
        
    strAllAttributes = Split(strAttributes, " ")
    For i = 0 To UBound(strAllAttributes)
        ReDim strThisAttribute(0)
        strThisAttribute = Split(strAllAttributes(i), "=")
        If UBound(strThisAttribute) <> -1 Then
            Select Case LCase$(strThisAttribute(0))
                Case "modes"
                    ServerAttributes.Modes = strThisAttribute(1)
                Case "watch"
                    ServerAttributes.Watch = strThisAttribute(1)
                Case "silence"
                    ServerAttributes.Silence = strThisAttribute(1)
                Case "chanlimit"
                    ServerAttributes.ChanLimit = strThisAttribute(1)
                Case "nicklen"
                    ServerAttributes.NickLen = strThisAttribute(1)
                Case "hostlen"
                    ServerAttributes.HostLen = strThisAttribute(1)
                Case "topiclen"
                    ServerAttributes.TopicLen = strThisAttribute(1)
                Case "kicklen"
                    ServerAttributes.KickLen = strThisAttribute(1)
                Case "channellen"
                    ServerAttributes.ChannelLen = strThisAttribute(1)
                Case "keylen"
                    ServerAttributes.KeyLen = strThisAttribute(1)
                Case "awaylen"
                    ServerAttributes.AwayLen = strThisAttribute(1)
                Case "maxbans"
                    ServerAttributes.MaxBans = strThisAttribute(1)
                Case "maxchannels"
                    ServerAttributes.MaxChannels = strThisAttribute(1)
                Case "maxtargets"
                    ServerAttributes.MaxTargets = strThisAttribute(1)
                Case "chantypes"
                    ServerAttributes.ChanTypes = strThisAttribute(1)
                Case "chanmodes"
                    ServerAttributes.ChanModes = strThisAttribute(1)
                Case "prefix"
                    ServerAttributes.PreFix = strThisAttribute(1)
                Case "casemapping"
                    ServerAttributes.CaseMapping = strThisAttribute(1)
                Case "network"
                    ServerAttributes.Network = strThisAttribute(1)
                Case "ircd"
                    ServerAttributes.IRCd = strThisAttribute(1)
                Case "map"
                    ServerAttributes.Map = True
                Case "knock"
                    ServerAttributes.Knock = True
                Case "wallchops"
                    ServerAttributes.WallChOps = True
                Case "safelist"
                    ServerAttributes.Safelist = True
                Case "hcn"
                    ServerAttributes.HCN = True
                Case Else
                    If UBound(strThisAttribute) = 1 Then
                        DB.X "005 Channel Attribute `" & strThisAttribute(0) & " = " & strThisAttribute(1) & "' could not be handled!"
                    Else
                        DB.X "005 Channel Attribute `" & strThisAttribute(0) & "' could not be handled!"
                    End If
            End Select
        End If
    Next i
End Sub
Public Sub Reset005Attributes()
    With ServerAttributes
        .Modes = -1
        .ChanLimit = -1
        .NickLen = -1
        .HostLen = -1
        .TopicLen = -1
        .KickLen = -1
        .ChannelLen = -1
        .KeyLen = -1
        .AwayLen = -1
        .MaxBans = -1
        .MaxChannels = -1
        .MaxTargets = -1
        .ChanModes = -1
        .Watch = -1
        .Silence = -1
        .ChanTypes = vbNullString
        .PreFix = vbNullString
        .CaseMapping = vbNullString
        .IRCd = vbNullString
        .Network = vbNullString
        .ChanModes = vbNullString
        .Map = False
        .Knock = False
        .Safelist = False
        .HCN = False
        .WallChOps = False
    End With
End Sub
Public Property Get IRCData(ByVal ChannelId As Integer) As String
    'zero-based array
    'the rest of the code (e.g. frmMain) is
    'one-based
    IRCData = txtStatus(ChannelId)
End Property
Public Property Let IRCData(ByVal ChannelId As Integer, ByVal NewValue As String)
    txtStatus(ChannelId) = NewValue
End Property
Public Property Get IRCData_Count()
    IRCData_Count = UBound(txtStatus) + 1
End Property
Public Property Let DCCFile_Progress(ByVal index As Integer, ByVal RCV As Boolean, ByVal NewValue As Byte)
    If RCV Then
        DCCFilesRCV(index).Progress = NewValue
    Else
        DCCFiles(index).Progress = NewValue
    End If
End Property
Public Property Get DCCFile_Progress(ByVal index As Integer, ByVal RCV As Boolean) As Byte
    If RCV Then
        DCCFile_Progress = DCCFilesRCV(index).Progress
    Else
        DCCFile_Progress = DCCFiles(index).Progress
    End If
End Property
Public Property Let DCCFile_StartTime(ByVal index As Integer, ByVal RCV As Boolean, ByVal NewValue As Long)
    If RCV Then
        DCCFilesRCV(index).StartTime = NewValue
    Else
        DCCFiles(index).StartTime = NewValue
    End If
End Property
Public Property Get DCCChats_Count() As Integer
    DCCChats_Count = UBound(DCCChats) + 1
End Property
Public Property Get DCCChats_TabIndex(ByVal index As Integer) As Integer
    DCCChats_TabIndex = DCCChats(index).TabIndex
End Property
Public Property Get DCCChats_UserName(ByVal index As Integer) As String
    DCCChats_UserName = DCCChats(index).UserName
End Property
Public Property Get DCCChats_WinSockIndex(ByVal index As Integer) As Integer
    DCCChats_WinSockIndex = DCCChats(index).WinSockIndex
End Property
Public Property Get DCCFile_WinSockIndex(ByVal index As Integer, RCV As Boolean) As Integer
    If RCV Then
        DCCFile_WinSockIndex = DCCFilesRCV(index).WinSockIndex
    Else
        DCCFile_WinSockIndex = DCCFiles(index).WinSockIndex
    End If
End Property
Public Property Get DCCFile_FileName(ByVal index As Integer, RCV As Boolean) As String
    If RCV Then
        DCCFile_FileName = DCCFilesRCV(index).FileName
    Else
        DCCFile_FileName = DCCFiles(index).FileName
    End If
End Property
Public Sub DCCFile_CancelTransfer(ByVal index As Integer, RCV As Boolean)
    If RCV Then
        DCCFilesRCV(index).Port = -1
        DCCFilesRCV(index).TabIndex = -1
        DCCFilesRCV(index).bytesSent = 0
        On Error Resume Next 'connection not initialized
        frmMain.wsDCC.Item(DCCFilesRCV(index).WinSockIndex).Close
        DCCFilesRCV(index).WinSockIndex = -1
        DCCFilesRCV(index).ETL = 0
        DCCFilesRCV(index).Speed = 0
        DCCFilesRCV(index).IPAddress = -1
    Else
        DCCFiles(index).Port = -1
        DCCFiles(index).TabIndex = -1
        DCCFiles(index).bytesSent = 0
        frmMain.wsDCCSend.Item(DCCFiles(index).WinSockIndex).Close
        DCCFiles(index).WinSockIndex = -1
        DCCFiles(index).ETL = 0
        DCCFiles(index).Speed = 0
        DCCFiles(index).IPAddress = -1
    End If
End Sub
Public Sub DCCChats_Disconnect(ByVal Nickname As String)
    Dim i As Integer, i2 As Integer
    Dim connum As Integer, nicksconn As Integer
    Dim strNickname As String, nickname2 As String
    
    connum = 1
    nickname2 = Nickname
    If InStr(1, Nickname, "(") > 0 Then
        connum = Val(Mid$(Nickname, InStr(1, Nickname, "(") + 1, Len(Nickname) - (InStr(1, Nickname, ")") - 1)))
        nickname2 = Left$(Nickname, InStr(1, Nickname, "(") - 1)
    End If
    
    For i = 1 To DCCChats_Count - 1
        If LCase$(DCCChats(i).UserName) = strNickname Then nicksconn = nicksconn + 1
    Next i
        
    i2 = 1
    strNickname = LCase$(nickname2)
    For i = 1 To DCCChats_Count - 1
        With DCCChats(i)
            If DCCChats(i).TabIndex > 0 Then
                If Me.Tabs.Tabs.Item(DCCChats(i).TabIndex).Caption = Nickname Then
                    If LCase$(.UserName) = strNickname Then
                    'And i2 = connum Then
                        frmMain.wsDCCChat.Item(.WinSockIndex).Close
                        frmMain.AddStatus "User has closed the connection" & vbNewLine, Me, GetChanID(Nickname)
                        .Port = -1
                        .TabIndex = -1
                        .UserName = vbNullString
                        .WinSockIndex = -1
                        Exit For
                    'ElseIf LCase$(.UserName) = strNickname And i2 <> connum Then
                    '    i2 = i2 + 1
                    End If
                End If
            End If
        End With
    Next i
End Sub
Public Sub IRCData_ReDim(ByVal NewDimention As Integer)
    ReDim Preserve txtStatus(NewDimention)
End Sub
Public Sub RegularOperations()
    Dim t As Long
    
    If ConnectRetry Then
        If WinSockConnection.State <> sckConnected Then
            t = GetTickCount
            If t > LastConnectionFailure + Options.ConnectRetryDelay * 1000 Then
                frmMain.AddStatus Language(780) & vbNewLine, Me
                frmMain.executeCommand "connect " & HostName & " " & Port & " """ & Description & """", GetServerIndexFromActiveServer(Me)
            End If
        End If
    End If
End Sub
Public Sub WinSockConnection_Close()
    m_Connected = False
    
    frmMain.AddNews Language(124)
    'the connection was lost or reset
    'let the user know
    frmMain.AddStatus "- " & Language(124) & " -" & vbNewLine, Me
    'there may has been an error;
    'close the connection in the case it wasn't already closed
    WinSockConnection.Close
    WinSockConnection.RemoteHost = vbNullString
    DoSessionPerform = False
    'clear the nicknames in every nicklist on every channel
    ClearAllNicklists
    If CurrentActiveServer Is Me Then
        'refresh the currently selected tab
        frmMain.tmrMakeItQuicker_Timer
    End If
    'we are not connected to any proxy
    m_bProxyPhase = 0
    
    ServerNode.Text = HostName & " (" & Language(772) & ")"
    ThisSoundSchemePlaySound "disconnect"
    
    'the user close the connection
    'so we're not going to bring it
    'back...
    ConnectRetry = False
End Sub
Private Sub WinSockConnection_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim strCancelLink As String
    Dim strNScript As String
    
    DB.Enter "clsActiveServer.WinSockConnection_Error"
    
    DB.X "WinSock Error: " & Description & " ( " & Number & " )"
    
    If Options.UseProxy Then
        frmMain.AddStatus Language(927) & vbNewLine, Me
    Else
        frmMain.AddStatus Language(69) & vbNewLine, Me
    End If
    frmMain.AddStatus WinSockErrorIDToLangString(Number) & vbNewLine, Me
    
    WinSockConnection.Close
    ThisSoundSchemePlaySound "error"
    
    ClearAllNicklists
    
    frmMain.tmrMakeItQuicker_Timer
    
    ServerNode.Text = HostName & " (" & Language(772) & ")"
    
    'we are not connected to any proxy
    bProxyPhase = 0
    
    LastConnectionFailure = GetTickCount
    ConnectRetry = Options.ConnectRetry
    If ConnectRetry Then
        GlobalLinkID = GlobalLinkID + 1
        strNScript = "NodeScript:/cancel_connect_retry " & GetServerIndexFromActiveServer(Me) & " " & GlobalLinkID
        strCancelLink = HTML_OPEN & "a href=""" & strNScript & """ id='cancel_connection_link_" & GlobalLinkID & "'" & HTML_CLOSE & Language(121) & HTML_OPEN & "/a" & HTML_CLOSE
        frmMain.AddStatus Replace(Replace(Language(779), "%1", CStr(Options.ConnectRetryDelay)), "%2", strCancelLink) & vbNewLine, Me
    End If

    DB.Leave "clsActiveServer.WinSockConnection_Error"
End Sub
Private Sub WinSockConnection_Connect()
    'connection with the server was established
    
    DB.Enter "clsActiveServer.WinSockConnection_Connect"
    
    Reset005Attributes
    ThisSoundSchemePlaySound "connect"
    
    'connect successful, we won't need to retry anything
    ConnectRetry = False
    
    If Options.UseProxy Then
        DB.X "We're using a proxy"
        'if we are using a proxy we will have to send
        'connection data to it in order to establish a
        'connection with the server on the other side
        'of the proxy
        
        'Connection with proxy established
        frmMain.AddStatus Language(576) & vbNewLine, Me
        'we are connected to the proxy
        m_bProxyPhase = 1
        
        'Send a version identifier/method selection message
        '
        'TO DO: NO AUTHENTICATION REQUIRED (DONE)
        'TO DO: GSAPI
        'TO DO: USERNAME/PASSWORD
        '
        'How to form a method selection/version identifier request:
        '                   +----+----------+----------+
        '                   |VER | NMETHODS | METHODS  |
        '                   +----+----------+----------+
        '                   | 1  |    1     | 1 to 255 |
        '                   +----+----------+----------+

        'wsIRC.SendData      X'05'   X'01'    X'00'
        '
        '...which basically means:
        '"Connect to SOCKS *5* proxy. We only support *one* method, this is method number *zero*
        '(no authentication required)
        
        DB.X "Proxy: Requesting SOCKS 5 with no authentication"
        
        WinSockConnection.SendData ChrW$(5) & ChrW$(1) & ChrW$(0)
        
        '
        'Other possible methods:
        '          o X '00' NO AUTHENTICATION REQUIRED
        '          o X '01' GSSAPI
        '          o X '02' USERNAME/PASSWORD
        '          o X '03' to X'7F' IANA ASSIGNED
        '          o X '80' to X'FE' RESERVED FOR PRIVATE METHODS
        '          o X 'FF' NO ACCEPTABLE METHODS
        
        'we have sent a version identifier/method selection message
        m_bProxyPhase = 2
    Else
        m_Connected = True
        DB.X "We're not using a Proxy"
        DB.X "Calling RealConnect"
        RealConnect
    
        ServerNode.Text = HostName
    End If
   
    DB.Leave "clsActiveServer.WinSockConnection_Connect"
End Sub
Private Sub RealConnect()
    Dim frfile As Integer 'variable used to store a free file index
    Dim mail As String, realn As String '
    Dim info As String, strTemp As String
    Dim strPerform As String
    Dim strPerformData() As String
    Dim i As Integer
    
    DB.Enter "clsActiveServer.RealConnect"
    
    With frmMain
        'connection with... established
        .AddStatus Language(58) & " " & CurrentActiveServer.WinSockConnection.RemoteHost & " " & Language(59) & vbNewLine, Me
        'let windows update
        DoEvents
        'if the file with the info of the user doesn't exist
        If Not FS.FileExists(App.Path & "\conf\info.dat") Then
            'display warning
            MsgBox Language(60), vbInformation, Language(61)
            'and close connection
            WinSockConnection.Close
            'let the user enter this information
            'select the `User Info' category
            For i = 1 To frmOptions.tvOptions.Nodes.Count
                If frmOptions.tvOptions.Nodes.Item(i).Key = "k00" Then
                    Set frmOptions.tvOptions.SelectedItem = frmOptions.tvOptions.Nodes.Item(i)
                    frmOptions.tvOptions_NodeClick frmOptions.tvOptions.SelectedItem
                    frmOptions.Show
                    frmOptions.tvOptions.SetFocus
                    Exit For
                End If
            Next i
            'update the frame with the options
            'frmOptions.lvList_Click
            'and show options dialog
            frmOptions.Show vbModal
            'quit this event, as we aren't connected
            Exit Sub
        End If
        
        DB.X "Reading User Data"

        'open the file with the user information in order to read it
        'and send it to the server
        'get a free file index
        frfile = FreeFile
        'open the file with the info
        Open App.Path & "\conf\info.dat" For Input As #frfile
        'the connection may not have been allowed, so create an error trap
        On Error GoTo Unable_To_SendInfo
        'read the user's nickname, store it to myNick varible,
        'and send it to the server using the NICK irc-command
        If AltNickVal = 0 Then
            myNick = xLineInput(frfile)
            'skip Alternative and Alternative 2 fields
            xLineInput frfile
            xLineInput frfile
        ElseIf AltNickVal = 1 Then
            xLineInput frfile
            myNick = xLineInput(frfile)
            xLineInput frfile
        ElseIf AltNickVal = 2 Then
            xLineInput frfile
            xLineInput frfile
            myNick = xLineInput(frfile)
        End If
        
        DB.X "Sending User Data"
        
        SendData "NICK " & myNick & vbNewLine
        
        'read the email and realname attributes and send
        'them and the local ip to the server using the USER irc-command
        'TO DO: Resolve external local IP! (in case of Routers/Firewalls) (DONE)
        SendData "USER " & xLineInput(frfile) & _
                       " """" """ & WinSockConnection.LocalIP & """ :" & _
                       xLineInput(frfile) & vbNewLine
        'close the user information file
        Close #frfile
        
        'read (M)essage (O)f (T)he (D)ay using MOTD irc-command
        'Not really necessary; MOTD command is automatically
        'executed by all modern IRC servers anyway
        'SendData "MOTD" & vbnewline
        
        'send USERHOST request
        SendData "USERHOST " & myNick & vbNewLine
        LocalLookupNick = myNick
        
        'inform the user that we sent his/her info.
        .AddStatus "- " & Language(62) & " -" & vbNewLine, Me
        
        'let windows update again
        DoEvents
        
        DB.X "Checking Perform"
        If DoSessionPerform Then
            DB.X "Session Perform Enabled"
            DB.X "Session Performing from SessionPerformFile := " & SessionPerformFile
            
            Perform SessionPerformFile
                        
            DoSessionPerform = False
            
        'execute the IRC commands in the file perform.dat
        'check to see if the perform file exists
        ElseIf FS.FileExists(App.Path & "/conf/perform.xml") And Options.EnablePerform Then
            DB.X "Perform Enabled"
            DB.X "XML Reading from File " & App.Path & "\conf\perform.xml"
            
            If Options.PerformSingle Then
                strPerform = XMLPerformReadServer(" default")
            Else
                strPerform = XMLPerformReadServer(HostName)
            End If
            
            DB.X "XML Perform read successful:"
            
            DB.Enter "Debug: Show Current Perform", "not a real procedure"
            DB.X strPerform
            DB.Leave "Debug: Show Current Perform"
            
            'this could simply be
            'WinScokConnection.SendData strPerform
            'but this would cause
            'Node commands not to be
            'executed (only IRC commands)
            'so we have to parse through
            'the string
            
            strPerformData = Split(strPerform, vbLf)
            
            'go through the file
            For i = 0 To UBound(strPerformData)
                'if it is a command...
                If Strings.Left$(strPerformData(i), 1) = "/" Then
                    '...execute it
                    preExecute strPerformData(i)
                End If
            'go to the next line
            Next i
        End If
        If Options.JoinPanel Then
            DB.X "Loading Join Panel"
            .LoadPanel "join"
        End If
        'don't show any error messages
EOF_Perform:
        If GetSetting(App.EXEName, "InfoTips", "connect", "0") = "0" Then
            DB.X "First-time connection; showing InfoTip"
            .ShowInfoTip Connected
            SaveSetting App.EXEName, "InfoTips", "connect", "1"
        End If
        'set IRC status
        DB.X "Setting Status to " & MyStatus
        If MyStatus = Status_Online Then
            preExecute "/away"
        Else
            preExecute "/away " & Replace(Language(409), "%1", Language(410 + MyStatus))
        End If
        DB.Leave "clsActiveServer.RealConnect"
        Exit Sub
Unable_To_SendInfo:
        'there was an error while trying to send USER, NICK, or MOTD commands. Let the user know.
        .AddStatus IMPORTANT_PREFIX & Language(63) & IMPORTANT_SUFFIX & vbNewLine, Me
    End With
    
    DB.Leave "clsActiveServer.RealConnect", "Unable to Send User Info"
End Sub
Public Sub Perform(ByVal CommandsFile As String)
    Dim intFL As Integer
    Dim strTemp As String
    
    intFL = FreeFile
    Open CommandsFile For Input As #intFL
    'go through the file
    Do Until EOF(intFL)
        'create an error trap in case
        'we are at the end of the file
        'without EOF having us informed
        On Error GoTo EOF_Perform
        'read a line
        strTemp = xLineInput(intFL)
        'if it is a command...
        If Strings.Left$(strTemp, 1) = "/" Then
            '...execute it
            preExecute strTemp
        End If
    'go to the next line
    Loop
    'close the perform file
    Close #intFL
EOF_Perform:
End Sub
Public Sub preExecute(ByVal strStatement, Optional ByVal ClearText As Boolean = True)
    'this is an important sub; it is called when the user types any text
    'or command, or when a command needs to be executed by the program
    'It either sends the text to the selected channel or private
    'or calls ExecuteCommand to execute the statement typed
    'If the user used the double slash it executes a VBScript statement
    Dim strExactStatement As String 'what is the exact command(without the leading slashes)
    Dim bolExecuted As Boolean 'was the command successfully executed?
    Dim strChan As String 'what's the current channel caption
    Dim i As Integer 'used in for loops
        
    If Left$(strStatement, 2) = vbNewLine Then
        strStatement = Right$(strStatement, Len(strStatement) - 2)
    End If
    
    With frmMain
        DoEvents 'let windows refresh
        strExactStatement = strStatement 'the ExactStatement will initially be as is
        'if the user typed two leading slashes(//)...
        If Strings.Left$(strExactStatement, 2) = "//" Then
            '...run the comand as a script
            'load sphere variables
            SetSphereVariables
            'the script may have an error; create an error trap
            On Error GoTo lbError
            'use ndScript to execute this script command
            .ndScript.ExecuteStatement Strings.Right$(strExactStatement, Len(strExactStatement) - 2)
        '...if the statement starts with '/'
        ElseIf Strings.Left$(strExactStatement, 1) = "/" Then
            'execute this IRC or client command
            'the exact statement will be the text typed without the leading slash
            strExactStatement = Strings.Right$(strExactStatement, Len(strExactStatement) - 1)
            'If it's an IRC command, we may not be connected
            'If it's a client command, it may contain invalid parameters
            'so, create an error trap
            On Error GoTo lbError
            'call executeCommand function to execute this command
            bolExecuted = .executeCommand(strExactStatement)
            'if it wasn't executed(so it's not a client command)...
            If Not bolExecuted Then
                'pass it to the server we are connected to
                SendData strExactStatement & vbNewLine
            End If
        Else
            'else send the message to the current query or channel
            'if the status tab is selected show the message "you are not on a channel"
            'the current channel or private is the caption of the selected tab
            strChan = Tabs.SelectedItem.Caption
            If TabType(Tabs.SelectedItem.index) = TabType_Channel Or TabType(Tabs.SelectedItem.index) = TabType_Private Then
                'we may not be connected
                On Error GoTo Not_Connected
                If DCCChats_Count >= 1 Then
                    For i = 1 To DCCChats_Count - 1
                        If DCCChats_TabIndex(i) = Tabs.SelectedItem.index Then
                            'check if the DCC chat has been unloaded
                            'to-do: fix the objectcollectionitemexists
                            'If ObjectCollectionItemExists(.wsDCCChat, DCCChats_WinSockIndex(i)) Then
                                'it's not unloaded
                                'check to see if a connection is present
                                If .wsDCCChat.Item(DCCChats_WinSockIndex(i)).State = 7 Then
                                    .wsDCCChat.Item(DCCChats_WinSockIndex(i)).SendData strStatement & vbNewLine
                                    .AddStatus "&lt; " & myNick & " &gt; " & strStatement & vbNewLine, Me, TabInfo(Tabs.SelectedItem.index)
                                    If CurrentActiveServer Is Me Then
                                        .buildStatus
                                    End If
                                    .txtSend.Text = vbNullString
                                    Exit Sub
                                Else
                                    preExecute "/close"
                                    GoTo Not_Connected
                                End If
                            'End If
                        End If
                    Next i
                End If
                'The current message can be either...
                '...a Private Message
                'use PRIVMSG to send it to the server to the current channel
                '
                '...or a Channel Message
                'use PRIVMSG again to send it to the server to the current private window(query)
                'Add text to the channel/private window
                SendData "privmsg " & strChan & " :" & strStatement & vbNewLine
                'display what the user just said
                '(either in a private window or on the channel)
                If TabType(Tabs.SelectedItem.index) = TabType_Private Then
                    .IRCAction ndPrivMsg, strChan, , myNick, strStatement, GetServerIndexFromActiveServer(Me)
                Else
                    .IRCAction ndPrivMsg, myNick, strChan, , strStatement, GetServerIndexFromActiveServer(Me)
                End If
                '.AddStatus "&lt; " & myNick & " &gt; " & strStatement & vbnewline, Me, TabInfo(Tabs.SelectedItem.Index)
                'and update the view
                .buildStatus
            Else
                'this is not a channel or query
                'You are not on a channel
                .AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(34) & EVENT_SUFFIX & vbNewLine, Me
                'exit sub; don't remove the text the user has typed. He or she may wanted to send it to a channel, but forgot to select it.
                Exit Sub
            End If
        End If
        If ClearText Then
            'Remove the sent text
            frmMain.txtSend.Text = vbNullString
        End If
        'skip the error-handle code(traps code)
        Exit Sub
lbError:
        If Err.Number = 40006 Then
            'not connected
            .AddStatus EVENT_PREFIX & Language(125) & EVENT_SUFFIX & vbNewLine, Me
            Exit Sub
        End If
        'there was an error: show it
        .AddStatus EVENT_PREFIX & Language(35) & " " & REASON_PREFIX & strStatement & REASON_SUFFIX & ": " & REASON_PREFIX & Err.Description & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, Me
        'skip the not-connected error handler code
        Exit Sub
Not_Connected:
        'show the connection error
        .AddStatus EVENT_PREFIX & Language(125) & ": " & REASON_PREFIX & Err.Description & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, Me
    End With
End Sub
Private Sub WinSockConnection_DataArrival(ByVal bytesTotal As Long)
    'recieving data from the winsock control...
    
    'If we are using a proxy...
    If Options.UseProxy Then
        'and we have not
        'yet recieved method/socks replies...
        If bProxyPhase = 2 Or bProxyPhase = 4 Then
            'we need
            'to check the data and handle it correctly
            'pass it to RealDataArrival_Proxy
            RealDataArrival_Proxy bytesTotal
            'the data has been handled correctly
            'using the SOCKS methods
            'we don't need to pass it to
            'the IRC message handling procedures
            'so exit sub
            Exit Sub
        End If
    End If
    'either we are not using a proxy
    'or the SOCKS handshacking procedure
    'has been completed
    'pass the data to the IRC message handling
    'procedures
    frmMain.RealDataArrival Me, bytesTotal
End Sub
Public Sub SendData(ByVal Data As String)
    DB.X "RAW-out:: " & Data
    On Error GoTo Failed_To_SendData
    WinSockConnection.SendData Data
    Exit Sub
Failed_To_SendData:
    WinSockConnection_Error -1, "Failed to SendData", 0, "SendData", vbNullString, 0, False
End Sub
Public Sub ClearAllNicklists()
    Dim i As Integer
    Dim i2 As Integer
    Dim intTabInfo As Integer
    Dim intTemp As Integer
    Dim childNode As ComctlLib.Node
        
    i = 0
    Do
        i = i + 1
        If i > Tabs.Tabs.Count Then
            Exit Do
        End If
        If TabType(i) = TabType_Channel Then
            ClearList NickList_List(TabInfo(i))
            If Not Options.KeepChannelsOpen Then
                'Close all channel tabs
                intTabInfo = TabInfo(i)
                'Unload txtStatus(intTabInfo)
                Set m_NickList(intTabInfo).List = Nothing
                intTemp = GetTab(intTabInfo)
                TabInfo.Remove intTemp
                TabType.Remove intTemp
                Tabs.Tabs.Remove intTemp
                DCCTabRemove intTemp
                Set Tabs.SelectedItem = Tabs.Tabs.Item(GetTab(GetStatusID))
                i = i - 1
            End If
        End If
    Loop
    
    If Not Options.KeepChannelsOpen Then
        Set childNode = ServerNode.Child
        For i2 = 1 To ServerNode.children
            If Left$(childNode.Key, 2) = "c_" Then
                intTemp = childNode.index
                Set childNode = childNode.Next
                frmMain.tvConnections.Nodes.Remove intTemp
            End If
        Next i2
    End If
    
    If CurrentActiveServer Is Me Then
        frmMain.UpdateTabsBar
    End If
    frmOptions.txtNickname.Tag = vbNullString
    frmOptions.txtAlt.Tag = vbNullString
    frmOptions.txtAltTwo.Tag = vbNullString
End Sub
Public Function GetTab(ByVal ChannelId As Long) As Integer
    Dim i As Integer
    
    For i = 1 To Tabs.Tabs.Count
        If TabType(i) <> TabType_WebSite And TabType(i) <> TabType_DCCFile And TabInfo(i) = ChannelId Then
            GetTab = i
            Exit Function
        End If
    Next i
    GetTab = -1 'No such channel
End Function
Public Function GetDCCFromTab(ByVal TabIndex As Integer, ByVal RCV As Boolean) As Integer
    Dim i As Integer
    Dim intForTarget As Integer
    
    If RCV Then
        intForTarget = UBound(DCCFilesRCV)
    Else
        intForTarget = UBound(DCCFiles)
    End If
    
    For i = 0 To intForTarget
        If RCV Then
            If DCCFilesRCV(i).TabIndex = TabIndex Then
                GetDCCFromTab = i
                Exit Function
            End If
        Else
            If DCCFiles(i).TabIndex = TabIndex Then
                GetDCCFromTab = i
                Exit Function
            End If
        End If
    Next i
End Function
Public Function DCCTransfer_Speed(ByVal DCCIndex As Integer, ByVal RCV As Boolean) As Double  'KB/s
    Dim dblDCCSpeed As Double
    Dim lStartTime As Long
    Dim lBytesSent As Long
    
    If RCV Then
        lStartTime = DCCFilesRCV(DCCIndex).StartTime
        lBytesSent = DCCFilesRCV(DCCIndex).bytesSent
    Else
        lStartTime = DCCFiles(DCCIndex).StartTime
        lBytesSent = DCCFiles(DCCIndex).bytesSent
    End If
    
    'get DCC speed in bytes per millisecond:
    'we can get the speed by dividing the
    'bytessent by the time the operation took
    'To get the time the operation took until now
    'we get the difference between now and the
    'time the operation started.
    dblDCCSpeed = lBytesSent / (GetTickCount - lStartTime)
    
    'get DCC speed in kilobytes per second:
    'both now and starttime are in milliseconds
    'multiplication by 1000 makes them seconds.
    'BytesSend are in bytes(obviously)
    'divition with 1024 makes them KBytes
    'Finally, we'll have to round to 2 decimal
    'digits using Round()
    DCCTransfer_Speed = 1000 * dblDCCSpeed / 1024 ' (in KB/s)
End Function
Public Function DCCTransfer_ETL(ByVal DCCIndex As Integer, ByVal RCV As Boolean) As String
    'TO DO: Do we really need a Double Type for .FileSize ?
    '       BytesSent is a Long, so there's no need for that waste!
    '       (Anyway if the size needed more bytes than a Long Type can hold,
    '        we'd need a Double for bytesSent as well) *>>
    Dim dblFileSize As Double
    Dim lBytesSent As Long
    Dim intDCCETL As Double, strETL As String
    Dim intDCCETL_Minutes As Byte, intDCCETL_Hours As Byte, intDCCETL_Days As Long
    
    If RCV Then
        dblFileSize = DCCFilesRCV(DCCIndex).FileSize
        lBytesSent = DCCFilesRCV(DCCIndex).bytesSent
    Else
        dblFileSize = DCCFiles(DCCIndex).FileSize
        lBytesSent = DCCFiles(DCCIndex).bytesSent
    End If
    
    'get DCC Estimated Time Left
    'To get the Estimated Time Left we have
    'to get the time that the rest of the bytes
    'are going to take to be transfered. This
    'can be done by dividing the number of bytes
    'by the speed we evaluated before.
    'as speed is in bytes/ms and filesize is in bytes
    'we will have to multiply with 1000 to get the number
    'of seconds left...
    On Error GoTo Infinitesimal
    '*>> take a look here, it's more obvious
    intDCCETL = Int((dblFileSize - lBytesSent) / (DCCTransfer_Speed(DCCIndex, RCV) * 1024))
    '                ^----------------------^
    
    intDCCETL_Days = intDCCETL \ 86400 '= (60 * 60 * 24)
    intDCCETL = intDCCETL - intDCCETL_Days * 86400 '= 60 * 60 * 24
    intDCCETL_Hours = intDCCETL \ 3600 '= (60 * 60)
    intDCCETL = intDCCETL - intDCCETL_Hours * CLng(3600) ' = 60 * 60
    intDCCETL_Minutes = intDCCETL \ 60
    intDCCETL = intDCCETL - intDCCETL_Minutes * 60
    If intDCCETL_Days > 0 Then
        strETL = intDCCETL_Days & " "
        strETL = strETL & Language(IIf(intDCCETL_Days = 1, 258, 257))
    End If
    If intDCCETL_Hours > 0 Then
        If LenB(strETL) <> 0 Then
            strETL = strETL & " "
        End If
        strETL = strETL & intDCCETL_Hours & " "
        strETL = strETL & Language(IIf(intDCCETL_Hours = 1, 256, 255))
    End If
    If intDCCETL_Minutes > 0 Then
        If LenB(strETL) <> 0 Then
            strETL = strETL & " "
        End If
        strETL = strETL & intDCCETL_Minutes & " "
        strETL = strETL & Language(IIf(intDCCETL_Minutes = 1, 254, 253))
    End If
    If intDCCETL > 0 Then
        If LenB(strETL) <> 0 Then
            strETL = strETL & " "
        End If
        strETL = strETL & intDCCETL & " "
        strETL = strETL & Language(IIf(intDCCETL = 1, 252, 251))
    End If
    If False Then
Infinitesimal:
        'MsgBox Err.Description
        strETL = Language(261)
    End If
    If LenB(strETL) = 0 Then
        strETL = Language(268)
    End If

    DCCTransfer_ETL = strETL
End Function
Public Sub DCCTransfer_SendFile(ByVal strUserName As String, ByVal strFile As String, ByVal blnHidden As Boolean) 'RCV: No
    
    'MsgBox "DCC transfers are not yet supported in Node 0.33! Sorry...", vbInformation
    
    'Exit Sub
    
    Dim i As Integer
    Dim filenm As String
    Dim yourip As String
    
    frmMain.wsDCCSend.LoadNew
        
    i = UBound(DCCFiles) + 1
    'DCCFiles is an array that stores all the info for sending and receiving files
    ReDim Preserve DCCFiles(i) 'add on to the array DCCFiles
    
    With DCCFiles(i)
        .FileName = strFile 'save the file name w/path in DCCFiles
        'save just the file name in DCCFiles also
        .FileNameShort = GetFileName(.FileName)
        filenm = DCCFiles(i).FileNameShort
        'save the file's size to DCCFiles
        .FileSize = FS.GetFile(.FileName).Size
        'save the i to DCCFiles because its the same as the index of the wsDCCSend
        'that the file will be sent over
        .WinSockIndex = frmMain.wsDCCSend.Count - 1
        'set the hidden to whatever was sent over
        .Hidden = blnHidden
        'get your ip address
        
        'TO DO:
        'Don't get the IP of the local machine.
        'Use resolved IP!
        
        yourip = LocalIP 'frmMain.wsDCCSend.Item(.WinSockIndex).LocalIP
        If LenB(yourip) = 0 Then
            yourip = WinSockConnection.LocalIP
        End If
        frmMain.wsDCCSend.Item(.WinSockIndex).Close
        
        'TO DO:
        'Port range, specified by user.
        
        frmMain.wsDCCSend.Item(.WinSockIndex).LocalPort = ChooseOpenPort(frmOptions.TxtRangeL, frmOptions.TxtRangeH)
        frmMain.wsDCCSend.Item(.WinSockIndex).Listen 'wait for a reply from the other user
        
        'send the NOTICE message saying that we want to send the other user a file
        filenm = Replace(filenm, " ", "_")
        preExecute "/NOTICE " & strUserName & " :DCC SEND " & filenm & "(" & yourip & ")" & vbNewLine
        yourip = IrcGetLongIP(yourip) 'get your long ip
        'replace any spaces in the file name with underscores
        'send the PRIVMSG saying that we want to send the other user a file
        preExecute "/PRIVMSG " & strUserName & " :DCC SEND " & filenm & " " & yourip & " " & frmMain.wsDCCSend.Item(.WinSockIndex).LocalPort & " " & DCCFiles(i).FileSize & "" & vbNewLine
        If DCCFiles(i).Hidden = False Then
            'add a new tab called: Sending filename...
            Tabs.Tabs.Add , , Language(209) & " " & filenm & "..."
            'save the tab in DCCFiles
            .TabIndex = Tabs.Tabs.Count
            TabType.Add TabType_DCCFile 'set the tab for a DCC
            TabInfo.Add 1 'sending
            If CurrentActiveServer Is Me Then
                frmMain.UpdateTabsBar
            End If
        End If
        .Port = frmMain.wsDCCSend.Item(.WinSockIndex).LocalPort 'save the port to DCCFiles
    End With
Canceled:
End Sub

'TO DO: RemotePort should be a signed int!
Public Sub DCCTransfer_Incoming(ByVal Nickname As String, ByVal FileName As String, ByVal FileSize As Long, ByVal RemotePort As Long)
    
    'MsgBox "DCC transfers are not yet supported in Node 0.33! Sorry...", vbInformation
    
    'Exit Sub
    
    Dim bolTemp As Boolean
    Dim strTemp As String
    Dim strTemp2 As String
    Dim i As Integer
    Dim i2 As Integer
    
    DCCTotal = DCCTotal + 1
    ReDim Preserve DCCFilesRCV(DCCTotal)
    DCCFilesRCV(DCCTotal).FileSize = FileSize
    DCCFilesRCV(DCCTotal).FileName = FileName
    DCCFilesRCV(DCCTotal).FileNameShort = GetFileName(FileName)
    'accept?
    bolTemp = False
    For i = 0 To UBound(AllowedHidden)
        If Nickname = AllowedHidden(i).AllowedNickname Then
            If DCCFilesRCV(DCCTotal).FileName = IIf(AllowedHidden(i).EventID = 1, "myavatar_", vbNullString) & AllowedHidden(i).AllowedFileName & IIf(AllowedHidden(i).EventID = 1, ".jpg", ".wav") Then
                If IPSender = AllowedHidden(i).AllowedIP Then
                    bolTemp = True
                    'addStatus "user is trying to send avatar or audio data" & vbnewline
                    DCCFilesRCV(DCCTotal).Hidden = True
                    'write to another filename
                    DCCFilesRCV(DCCTotal).FileName = AllowedHidden(i).WriteFileName & IIf(AllowedHidden(i).EventID = 1, ".jpg", ".wav")
                End If
            End If
        End If
    Next i
    If bolTemp = False Then
        DCCFilesRCV(DCCTotal).Hidden = False
        If frmOptions.lstBdyNk.ListCount > 0 Then
            For i2 = 0 To frmOptions.lstBdyNk.ListCount - 1
                If Nickname = frmOptions.lstBdyNk.List(i2) Then
                    If frmOptions.optDCCB(0).value Then
                        bolTemp = True
                        Exit For
                    ElseIf frmOptions.optDCCB(1).value Then
                        Exit Sub
                    End If
                End If
            Next i2
        End If
    End If
    If bolTemp = False Then
        If frmOptions.lstIgnore(0).ListCount > 0 Then
            For i2 = 0 To frmOptions.lstIgnore(0).ListCount - 1
                If Nickname = frmOptions.lstIgnore(0).List(i2) Then
                    If frmOptions.optDCCI(0).value Then
                        bolTemp = True
                        Exit For
                    ElseIf frmOptions.optDCCI(1).value Then
                        Exit Sub
                    End If
                End If
            Next i2
        End If
    End If
    If bolTemp = False Then
        If frmOptions.optDCCE(0).value Then
            bolTemp = True
        ElseIf frmOptions.optDCCE(1).value Then
            Exit Sub
        End If
    End If
    If bolTemp = False Then
        If GetSetting(App.EXEName, "InfoTips", "DCCIncoming", "0") = "0" Then
            frmMain.ShowInfoTip DCCIncoming
            SaveSetting App.EXEName, "InfoTips", "DCCIncoming", "1"
        End If
        strTemp = MsgBox(Replace(Replace(Language(699), "%1", Nickname), "%2", DCCFilesRCV(DCCTotal).FileName) & ". " & Language(204), vbYesNo Or vbQuestion, Language(205))
    Else
        'automatically accept the transfer
        strTemp = vbYes
    End If
    If strTemp = vbNo Then
        frmMain.AddStatus IMPORTANT_PREFIX & Language(206) & IMPORTANT_SUFFIX & vbNewLine, Me
        Exit Sub
    Else
        frmMain.AddStatus IMPORTANT_PREFIX & Language(207) & IMPORTANT_SUFFIX & vbNewLine, Me
        On Error Resume Next 'already loaded
        If DCCFilesRCV(DCCTotal).Hidden Then
            If AllowedHidden(i).EventID = 1 Then
                If FS.FileExists(App.Path & "/temp/" & DCCFilesRCV(DCCTotal).FileName) Then
                    FS.DeleteFile (App.Path & "/temp/" & DCCFilesRCV(DCCTotal).FileName)
                End If
            ElseIf AllowedHidden(i).EventID = 2 Then
                If FS.FileExists(App.Path & "/temp/" & DCCFilesRCV(DCCTotal).FileName) Then
                    FS.DeleteFile (App.Path & "/temp/" & DCCFilesRCV(DCCTotal).FileName)
                End If
            End If
        ElseIf FS.FileExists(App.Path & "/downloads/" & DCCFilesRCV(DCCTotal).FileName) Then
            'resume?
            'TO DO: Display filename in case of auto-accept.
            strTemp2 = MsgBox(Language(250), vbYesNo Or vbQuestion, Language(633))
            If strTemp2 = vbYes Then
                Dim ffile As File
                Set FS = New FileSystemObject
                Set ffile = FS.GetFile(App.Path & "/downloads/" & DCCFilesRCV(DCCTotal).FileNameShort)
                frmMain.AddStatus IMPORTANT_PREFIX & Replace(Language(717), "%1", FileName) & IMPORTANT_SUFFIX & vbNewLine, Me
                SendData "PRIVMSG " & Nickname & " :" & Strings.ChrW$(1) & "DCC RESUME " & DCCFilesRCV(DCCTotal).FileName & " " & RemotePort & " " & ffile.Size & ChrW$(1) & vbNewLine
                DCCFilesRCV(DCCTotal).IPAddress = IPSender
                DCCFilesRCV(DCCTotal).Port = RemotePort
                'addStatus DCCFilesRCV(intTemp).Port & "/" & intTemp
                Exit Sub
            Else
                FS.DeleteFile (App.Path & "/downloads/" & DCCFilesRCV(DCCTotal).FileName)
            End If
        End If
        With frmMain.wsDCC.Item(frmMain.wsDCC.LoadNew)
            DCCFilesRCV(DCCTotal).WinSockIndex = frmMain.wsDCC.Count - 1
            DCCFilesRCV(DCCTotal).IPAddress = IPSender
            If DCCFilesRCV(DCCTotal).Hidden = False Then
                Tabs.Tabs.Add , , Language(208) & " " & DCCFilesRCV(DCCTotal).FileName
                DCCFilesRCV(DCCTotal).TabIndex = Tabs.Tabs.Count
                TabType.Add TabType_DCCFile
                TabInfo.Add 0 'receiving
                If CurrentActiveServer Is Me Then
                    frmMain.UpdateTabsBar
                End If
            End If
            .Close
            'TO DO: Don't use a random port. Add range selection support
            .LocalPort = ChooseOpenPort(frmOptions.TxtRangeL, frmOptions.TxtRangeH)
            .Connect IPSender, RemotePort
        End With
        Exit Sub
    End If
End Sub
'TO DO: Port as signed int
Public Sub DCCTransfer_Accepted(ByVal Nickname As String, ByVal FileName As String, ByVal Port As Long, ByVal FileSize As Long)
    Dim i As Integer
    
    For i = 1 To UBound(DCCFilesRCV)
        If DCCFilesRCV(i).Port = Port Then
            Exit For
        End If
    Next i
    
    If DCCFilesRCV(i).Hidden = False Then
        'AddStatus DCCFilesRCV(i).Port & "/" & i
        Tabs.Tabs.Add , , Language(208) & " " & DCCFilesRCV(i).FileName
        DCCFilesRCV(i).TabIndex = Tabs.Tabs.Count
        TabType.Add TabType_DCCFile
        TabInfo.Add 0 'recieving
        If CurrentActiveServer Is Me Then
            frmMain.UpdateTabsBar
        End If
        DCCFilesRCV(i).TabIndex = Tabs.Tabs.Count
    End If
    With frmMain.wsDCC.Item(frmMain.wsDCC.LoadNew)
        .Close
        .Connect DCCFilesRCV(i).IPAddress, Port
    End With
End Sub
Public Sub DCCTransfer_Resuming(ByVal Nickname As String, ByVal Port As Long, ByVal FilePos As Long) 'RCV: No
    Dim i As Integer
    
    For i = 1 To UBound(DCCFiles)
        If DCCFiles(i).Port = Port Then
            Exit For
        End If
    Next i
    
    'TO DO:
    'SECURITY ISSUE: If a remote user sends a `DCC RESUME' request
    '                when we haven't offered to send a file, Node
    '                will crash.
    
    frmMain.AddStatus IMPORTANT_PREFIX & Language(246) & " " & DCCFiles(i).FileName & IMPORTANT_SUFFIX & vbNewLine, Me
    DCCFiles(i).Resuming = 1
    DCCFiles(i).bytesSent = FilePos
    
    'TO DO:
    'Jnfoot, please check that. FilePos has already
    'been converted to Long, is it OK if we send
    'the converted value?? Or should we send the
    'original string?
    
    SendData "PRIVMSG " & Nickname & " :" & Strings.ChrW$(1) & "DCC ACCEPT " & Strings.ChrW$(34) & Strings.ChrW$(34) & " " & Port & " " & FilePos & Strings.ChrW$(1) & vbNewLine
    If DCCFiles(i).Hidden = False Then
        TabType.Add TabType_DCCFile
        TabInfo.Add 1 'sending
        If CurrentActiveServer Is Me Then
            frmMain.UpdateTabsBar
        End If
    End If
End Sub
Public Sub IncomingDCCChatRequest(ByVal Nickname As String, ByVal Port As Long)
    'somebody wants to DCC chat. Connect.
    Dim i As Integer
    Dim DCCIndex As Integer
    Dim IRCDataIndex As Integer
    
    'TO DO:
    '   Add auto-accept, auto-refuse DCC chat feature,
    '   similar to DCC send auto-accept, auto-refuse.
    
    i = frmMain.wsDCCChat.Count
    If MsgBox(Replace(Language(701), "%1", Nickname), vbYesNo Or vbQuestion, Language(205)) = vbNo Then
        'resused the incoming dcc chat request
        frmMain.AddStatus IMPORTANT_PREFIX & Language(241) & IMPORTANT_SUFFIX & vbNewLine, Me
        Exit Sub
    Else
        'accepted the incoming dcc chat request
        frmMain.AddStatus IMPORTANT_PREFIX & Language(242) & IMPORTANT_SUFFIX & vbNewLine, Me
        Load frmMain.wsDCCChat.Item(i)
        DCCIndex = UBound(DCCChats) + 1
        ReDim Preserve DCCChats(DCCIndex)
        DCCChats(DCCIndex).WinSockIndex = i
        frmMain.wsDCCChat.Item(i).Connect IPSender, Port
        DCCChats(DCCIndex).WinSockIndex = i
        DCCChats(DCCIndex).UserName = Nickname
        DCCChats(DCCIndex).Port = frmMain.wsDCCChat.Item(i).LocalPort
        Tabs.Tabs.Add , , Nickname, TabImage_Private
        TabInfo.Add xLet(IRCDataIndex, GetEmptyChannelID)
        TabType.Add TabType_Private
        If CurrentActiveServer Is Me Then
            frmMain.UpdateTabsBar
        End If
        DCCChats(DCCIndex).TabIndex = Tabs.Tabs.Count
        If IRCDataIndex > IRCData_Count - 1 Then
            IRCData_ReDim IRCDataIndex
        End If
        IRCData(IRCDataIndex) = vbNullString
        frmMain.AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Replace(Language(718), "%1", REASON_PREFIX & Nickname & REASON_SUFFIX) & EVENT_SUFFIX & vbNewLine, Me, GetChanID(Nickname), False
        Exit Sub
    End If
End Sub
Public Sub DCCTabRemove(ByVal intTabIndex As Integer)
    Dim i As Integer
    
    For i = 0 To UBound(DCCFiles)
        If DCCFiles(i).TabIndex > intTabIndex Then
            DCCFiles(i).TabIndex = DCCFiles(i).TabIndex - 1
        End If
    Next i
    For i = 0 To UBound(DCCFilesRCV)
        If DCCFilesRCV(i).TabIndex > intTabIndex Then
            DCCFilesRCV(i).TabIndex = DCCFilesRCV(i).TabIndex - 1
        End If
    Next i
    For i = 0 To UBound(DCCChats)
        If DCCChats(i).TabIndex > intTabIndex Then
            DCCChats(i).TabIndex = DCCChats(i).TabIndex - 1
        End If
    Next i
End Sub
Public Function GetStatusID() As Long
    Dim i As Integer
    Dim intStatusIndex As Integer
    
    For i = 1 To Tabs.Tabs.Count
        On Error GoTo TabType_NotSet
        If TabType(i) = TabType_Status Then
            GetStatusID = TabInfo(i)
            Exit Function
        End If
    Next i
    'No status tab; create one
    intStatusIndex = Tabs.Tabs.Add(, , vbNullString, TabImage_Status).index
TabType_NotSet_2:
    TabInfo.Add xLet(i, UBound(txtStatus))
    TabType.Add TabType_Status
    
    'do NOT move these properties to the tab creation
    'call, as it is not always made, as the tab may already have been
    'created if we got here from TabType_NotSet
    Tabs.Tabs.Item(intStatusIndex).Caption = Language(70)
    Tabs.Tabs.Item(intStatusIndex).Image = TabImage_Status
    
    If CurrentActiveServer Is Me Then
        frmMain.UpdateTabsBar
    End If
    
    If i > UBound(txtStatus) Then
        ReDim Preserve txtStatus(i)
    End If
    ReDim Preserve m_NickList(i)
    Set m_NickList(i).List = New Collection
    NickList_Size(i) = NickList_DefaultSize
    GetStatusID = i
    'frmMain.tsTabs_Click TabsIndex
    Exit Function
TabType_NotSet:
    'the tab type of this tab has not been set
    'this means that this is the first tab when
    'the program starts, so we'll have to set the
    'type of the tab right now to Status Tab
    'this is only done if the first tab is
    'not a web site tab, so only if the user
    'has selected not to load a web site on
    'startup
    intStatusIndex = 1
    GoTo TabType_NotSet_2
End Function
Public Sub RequestDCCChat(ByVal Nickname As String)
    'initiate the DCC Chat request
    Dim i As Integer
    Dim lIP As String
    
    i = UBound(DCCChats) + 1
    
    With frmMain.wsDCCChat.Item(frmMain.wsDCCChat.LoadNew)
        ReDim Preserve DCCChats(i)
        DCCChats(i).WinSockIndex = frmMain.wsDCCChat.Count - 1
        .Close
        .LocalPort = ChooseOpenPort(frmOptions.TxtRangeL, frmOptions.TxtRangeH)
        DCCChats(i).Port = .LocalPort
        DCCChats(i).UserName = Nickname
        .Listen 'wait for reply from the other user
        'send the NOTICE that we are requesting a DCC Chat
        preExecute "/NOTICE " & Nickname & " :DCC CHAT (" & LocalIP & ")"
        lIP = IrcGetLongIP(LocalIP) 'get our long ip address
        'send the PRIVMSG  that we are requesting a DCC Chat
        preExecute "/PRIVMSG " & Nickname & " :" & MIRC_CTCP & "DCC CHAT chat " & lIP & " " & .LocalPort & MIRC_CTCP
    End With
End Sub
Public Function GetDCCFileIndexFromWsIndex(ByVal intWsDccIndex As Integer, ByVal RCV As Boolean) As Integer
    Dim i As Integer
    Dim intForTarget As Integer
    Dim bolSuccess As Boolean
    
    If RCV Then
        intForTarget = UBound(DCCFilesRCV)
    Else
        intForTarget = UBound(DCCFiles)
    End If
    
    For i = 1 To intForTarget
        If RCV Then
            bolSuccess = DCCFilesRCV(i).WinSockIndex = intWsDccIndex
            If DCCFilesRCV(i).WinSockIndex = intWsDccIndex Then
                bolSuccess = True
            End If
        Else
            bolSuccess = DCCFiles(i).WinSockIndex = intWsDccIndex
            If DCCFiles(i).WinSockIndex = intWsDccIndex Then
                bolSuccess = True
            End If
        End If
        If bolSuccess Then
            GetDCCFileIndexFromWsIndex = i
            Exit Function
        End If
    Next i

    GetDCCFileIndexFromWsIndex = -1
End Function
Public Sub DccRcv_Error(ByVal WinSockIndex As Integer)
    'if there is an error, close the file and the winsock
    'so that it can be tried again
    Close DCCFilesRCV(GetDCCFileIndexFromWsIndex(WinSockIndex, True)).BLNVar
    frmMain.wsDCC.Item(WinSockIndex).Close
End Sub
Public Sub DccRcv_DataReceived(ByVal WinSockIndex As Integer)
    Dim xbyte() As Byte
    Dim rplyhex
    Dim intFL As Long
    Dim i As Integer
    Dim sendrply() As Byte
    Dim intNewProgress As Integer
    Dim strWritePath As String
    Dim DCCIndex As Integer
    
    DCCIndex = GetDCCFileIndexFromWsIndex(WinSockIndex, True)
    DoEvents
    
    DCCFilesRCV(DCCIndex).BLNVar = FreeFile 'get a free file number
    
    With DCCFilesRCV(DCCIndex)
        'open the file that you are receiving
        'On Error GoTo Create_Downloads
        If DCCFilesRCV(DCCIndex).Hidden Then
            strWritePath = App.Path & "/temp/" & DCCFilesRCV(DCCIndex).FileName
        Else
            strWritePath = App.Path & "/downloads/" & DCCFilesRCV(DCCIndex).FileName
        End If
        
        Open strWritePath For Binary As .BLNVar
        
        On Error GoTo Error_Getting_Data
        frmMain.wsDCC.Item(WinSockIndex).GetData xbyte, vbByte 'get the info received over the winsock
        'find the byte after the last one in the file
        Seek .BLNVar, LOF(.BLNVar) + 1
        'save the new bytes where we seeked in the last line
        Put .BLNVar, , xbyte
        'convert the file size of the currently receiving file to hex
        rplyhex = Conversion.Hex(LOF(.BLNVar))
        rplyhex = Strings.String$(8 - Len(rplyhex), "0") & rplyhex
        ReDim sendrply(3) As Byte
        'set up the reply to the sender of the file converting it to a byte
        For i = 1 To Len(rplyhex) Step 2
            sendrply((i - 1) \ 2) = Val("&H" & Strings.Mid$(rplyhex, i, 2))
        Next i
        'send the sender the reply, so that you will receive more of the file
        frmMain.wsDCC.Item(WinSockIndex).SendData sendrply
        'find out how much has been received
        intNewProgress = Int((LOF(.BLNVar) / .FileSize) * 100)
        .bytesSent = LOF(.BLNVar) 'actually bytes recieved(or bytes sent by the remote client)
        'check the size of the received file to the size that was supposed to be sent
        'if its the same then "Received filename" and close the winsock as its
        'not needed any more
        If LOF(.BLNVar) = .FileSize Then
            Close .BLNVar
            frmMain.AddStatus SPECIAL_PREFIX & Language(158) & " " & DCCFilesRCV(DCCIndex).FileName & " " & Language(159) & SPECIAL_SUFFIX & vbNewLine, Me
            frmMain.wsDCC.Item(WinSockIndex).Close
            'we recieved the file successfully
            'if we need to virus-check it...
            If Len(Options.DCCAntivirus) <> 0 Then
                Shell Options.DCCAntivirus & " """ & strWritePath & """"
            End If
            For i = 0 To UBound(AllowedHidden)
                If .FileNameShort = "myavatar_" & AllowedHidden(i).AllowedFileName & ".jpg" Then
                    If frmMain.wsDCC.Item(WinSockIndex).RemoteHostIP = AllowedHidden(i).AllowedIP Then
                    
                        '1 = avatars event ID
                        If AllowedHidden(i).EventID = 1 Then
                            If FS.FileExists(App.Path & "/conf/avatars/" & DCCFilesRCV(DCCIndex).FileName) Then
                                FS.DeleteFile App.Path & "/conf/avatars/" & DCCFilesRCV(DCCIndex).FileName
                            End If
                            FileCopy App.Path & "/temp/" & AllowedHidden(i).WriteFileName & ".jpg", App.Path & "/conf/avatars/" & DCCFilesRCV(DCCIndex).FileName
                            NDCConnections(AllowedHidden(i).NDCConnectNUM).UserAvatar = App.Path & "/conf/avatars/" & DCCFilesRCV(DCCIndex).FileName
                            If TabType(Tabs.SelectedItem.index) = TabType_Private Then
                                If NDCConnections(AllowedHidden(i).NDCConnectNUM).strNicknameA = Tabs.SelectedItem.Caption Then
                                    If frmMain.strCurrentPanel = "avatar" Then
                                        frmMain.tmrPanelRefreshSoon.Enabled = False
                                        frmMain.tmrPanelRefreshSoon.Enabled = True
                                    End If
                                    frmMain.sbar.Panels(2).Text = Language(158) & " " & Replace(Language(458), "%1", NDCConnections(AllowedHidden(i).NDCConnectNUM).strNicknameA) & " " & Language(159)
                                End If
                            End If
                             
                        '2 = audio data event ID
                        ElseIf AllowedHidden(i).EventID = 2 Then
                            'we recieved some audio data
                            NDCConnections(AllowedHidden(i).NDCConnectNUM).AudioFile = AllowedHidden(i).WriteFileName
                            'playback the audio file
                            'MP.FileName = AllowedHidden(i).WriteFileName
                            'MP.Play
                        End If
                    End If
                End If
            Next i
        Else
            Close .BLNVar
        End If
        If .Hidden = False Then
            If intNewProgress <> .Progress Then
                'On Error GoTo closedtab
                .Progress = intNewProgress
                Tabs.Tabs.Item(.TabIndex).Caption = Language(208) & " " & .FileNameShort & "(" & .Progress & "%)"
                If CurrentActiveServer Is Me Then
                    frmMain.UpdateTabsBar
                    If GetDCCFromTab(Tabs.SelectedItem.index, True) = DCCIndex Then
                        frmMain.buildStatus
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
ClosedTab:
    frmMain.wsDCC.Item(WinSockIndex).Close
    Exit Sub
Create_Downloads:
Invalid_Path:
    'The path `/downloads' does not exist
    'create it
    If DCCFilesRCV(DCCIndex).Hidden Then
        MkDir App.Path & "\temp"
    Else
        MkDir App.Path & "\downloads"
    End If
    Resume
Error_Getting_Data:
    'DCC Error  - TO DO in 0.32
End Sub
Public Sub DCCChat_Accepted(ByVal Nickname As String, ByVal Port As Long, ByVal IPaddy As String)
    Dim i As Integer, i2 As Integer, i3 As Integer
    Dim chanID As Long
    Dim nickname2 As String
    
    i = UBound(DCCChats) + 1

    With frmMain.wsDCCChat.Item(frmMain.wsDCCChat.LoadNew)
        ReDim Preserve DCCChats(i)
        DCCChats(i).WinSockIndex = frmMain.wsDCCChat.Count - 1
        .Close
        .LocalPort = ChooseOpenPort(frmOptions.TxtRangeL, frmOptions.TxtRangeH)
        DCCChats(i).Port = .LocalPort
        DCCChats(i).UserName = Nickname
        .Connect IPaddy, Port
    End With
    
    i2 = 1
    For i3 = 1 To Tabs.Tabs.Count
        If InStr(1, Tabs.Tabs.Item(i3).Caption, Nickname) > 0 Then i2 = i2 + 1
    Next i3
    If i2 > 1 Then Nickname = Nickname & "(" & i2 & ")"
    
    Tabs.Tabs.Add , , Nickname, TabImage_Private
    DCCChats(i).TabIndex = Tabs.Tabs.Count
    TabType.Add TabType_Private
    TabInfo.Add xLet(chanID, GetEmptyChannelID)
    frmMain.tvConnections.Nodes.Add Me.ServerNode, tvwChild, "d_" & GetServerIndexFromActiveServer(Me) & "_" & Nickname & "_chat" & Tabs.Tabs.Count, Language(803) & ": " & Nickname, TabImage_Private
    If CurrentActiveServer Is Me Then
        frmMain.UpdateTabsBar
    End If
    If chanID > UBound(txtStatus) Then
        ReDim Preserve txtStatus(chanID)
    End If
    txtStatus(chanID) = vbNullString
    frmMain.AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Replace(Language(718), "%1", DCCChats(i).UserName) & EVENT_SUFFIX & vbNewLine, Me, GetChanID(Nickname), False
End Sub

Public Sub DccChat_ConnectionRequest(ByVal WinSockIndex As Integer, ByVal requestID As Integer)
    'ConnectionRequest after *we* send a DCC Chat Request
    Dim i As Integer, i2 As Integer, i3 As Integer
    Dim chanID As Long
    Dim DCCIndex As Integer
    Dim Nickname As String
    
    'TO DO:
    'SECURITY ISSUE: If, after a user makes a DCC Chat request
    '                somebody else connects to us, Node will
    '                accept the connection and allow a DCC
    '                Chat, allowing a user to fake messages.
    '                We should add code to check if the IP
    '                of the remote computer is the one that
    '                did the initial request.
    'For i = 1 To UBound(DCCChats)
    '    If DCCChats(i).Port = frmMain.wsDCCChat.Item(WinSockIndex).LocalPort Then
    '        DCCIndex = i
    '        Exit For
    '    End If
    'Next i
    DCCIndex = Me.GetDCCChatsIndexFromWsIndex(WinSockIndex)
    
    i2 = 1
    For i3 = 1 To Tabs.Tabs.Count
        If InStr(1, Tabs.Tabs.Item(i3).Caption, DCCChats(DCCIndex).UserName) > 0 Then i2 = i2 + 1
    Next i3
    If i2 > 1 Then
        Nickname = DCCChats(DCCIndex).UserName & "(" & i2 & ")"
    Else
        Nickname = DCCChats(DCCIndex).UserName
    End If
    
    With frmMain.wsDCCChat.Item(WinSockIndex)
        .Close
        .accept requestID 'allow the connection
    
        'create a new tab named: Nickname of other person@port connected
        Tabs.Tabs.Add , , Nickname, TabImage_Private
        frmMain.tvConnections.Nodes.Add Me.ServerNode, tvwChild, "d_" & GetServerIndexFromActiveServer(Me) & "_" & Nickname & "_chat" & Tabs.Tabs.Count, Language(803) & ": " & Nickname, TabImage_Private
        TabInfo.Add xLet(chanID, GetEmptyChannelID)
        TabType.Add TabType_Private 'set the tab type to private
        DCCChats(DCCIndex).TabIndex = Tabs.Tabs.Count
        DCCChats(DCCIndex).Port = .LocalPort
    End With
    
    If CurrentActiveServer Is Me Then
        frmMain.UpdateTabsBar
    End If
    If chanID > UBound(txtStatus) Then
        ReDim Preserve txtStatus(chanID)
    End If
    txtStatus(chanID) = vbNullString
    'announce that the other person has entered the chat session
    frmMain.AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & REASON_PREFIX & DCCChats(DCCIndex).UserName & REASON_SUFFIX & " has joined DCC Chat." & EVENT_SUFFIX & vbNewLine, Me, GetChanID(Nickname), False
End Sub
Public Sub DccChat_DataReceived(ByVal WinSockIndex As Integer, ByVal Data As String)
    Dim i As Integer
    Dim Pos As Integer
    Dim GoodTab As Integer
    Dim DCCIndex As Integer
    
    'find which tab is supposed to show the info
    GoodTab = -1
    
        For i = 1 To UBound(DCCChats)
            If DCCChats(i).Port = frmMain.wsDCCChat.Item(WinSockIndex).LocalPort Then
                GoodTab = DCCChats(i).TabIndex
                Exit For
            End If
        Next i
        If GoodTab > -1 Then
            'if the message is an action, display it correctly
            If Strings.Left$(Data, 7) = "ACTION" Then
                'eliminate the ChrW$(1) from the beginning and end
                Data = Replace(Data, "", vbNullString)
                Pos = InStr(1, Data, " ")
                'show the message the way an action is supposed to be displayed
                frmMain.AddStatus HTML_OPEN & "font color=#FF00FF" & HTML_CLOSE & HTML_OPEN & "strong" & HTML_CLOSE & _
                "* " & DCCChats(i).UserName & Strings.Mid$(Data, Pos) & HTML_OPEN & "/strong" & HTML_CLOSE & HTML_OPEN & "/font" & HTML_CLOSE, Me, TabInfo(GoodTab)
            Else 'its not an action, just display it then
                frmMain.AddStatus "&lt;" & DCCChats(i).UserName & "&gt; " & Data, Me, TabInfo(GoodTab)
            End If
        End If
End Sub
Public Sub DccSend_Connected(ByVal WinSockIndex As Integer) 'RCV: NO
    Dim lSize As Long
    Dim fLength As Long
    Dim Buffer As String
    Dim DCCIndex As Integer
    
    DCCIndex = GetDCCFileIndexFromWsIndex(WinSockIndex, False)
    With DCCFiles(DCCIndex)
        .StartTime = GetTickCount
        'resuming file transfer, connected
        'Display: "DCC Connection with 127.0.0.1 was established"
        frmMain.AddStatus Replace(Language(715), "%1", .IPAddress) & vbNewLine, Me
        .BLNVar = FreeFile 'find a free file number
        'open the file that is to be sent
        Open .FileName For Binary Access Read As .BLNVar
        lSize = 1024 'the amount of data to be sent
        fLength = LOF(.BLNVar) 'size of the file to be sent
        'if there isnt enough bytes left to satisfy the normal amount sent
        'then set the size to be sent to whatever amount is left in the file
        'to be sent
        If fLength - Loc(.BLNVar) < lSize Then
            lSize = fLength - Loc(.BLNVar)
        End If
        'if the size is 0, then we have sent the message
        If lSize = 0 Then
            Exit Sub
        End If
        'save the amount sent to DCCFiles
        .bytesSent = .bytesSent + lSize
        Buffer = Space(lSize)
        'get the bytes to be sent from the file
        Get .BLNVar, , Buffer
        'send the bytes to the other user
        frmMain.wsDCCSend.Item(WinSockIndex).SendData Buffer
    End With
End Sub
Public Sub DccSend_ConnectionRequest(ByVal WinSockIndex As Integer, ByVal requestID As Long) 'RCV: NO
    Dim DCCIndex As Integer
    Dim Buffer As String
    Dim lSize As Long
    Dim i As Long
    
    DCCIndex = GetDCCFileIndexFromWsIndex(WinSockIndex, False)
       
    'Resuming file transfer
    With DCCFiles(DCCIndex)
        .IPAddress = frmMain.wsDCCSend.Item(WinSockIndex).RemoteHostIP
        'check and see if the other user is resuming the file being sent
        If .Resuming = 0 Then
            'nope;
            'start new transfer
            DccSend_Connected WinSockIndex
            Exit Sub
        End If
        
        .StartTime = GetTickCount
        
        'the user is resuming so we need to find where in the file to start
        .BLNVar = FreeFile 'get a free file number
        'open the file to be sent
        Open .FileName For Binary Access Read As .BLNVar
        i = 0
        'loop until i = the place in the file that the other user said that
        'they are resuming from
        Do Until i >= .bytesSent
            'count by the normal amount of bytes that would be sent
            lSize = 1024
            'if we are going to go over the amount we were supposed to,
            'then shrink the bytes to check to make us get to the
            'correct place in the file
            If lSize + i >= .bytesSent Then
                lSize = .bytesSent - i
                Exit Do 'we are in the correct place so stop moving through the file
            End If
            Buffer = Space(lSize)
            'get the bytes just so that we can move further in the file
            Get .BLNVar, , Buffer
            i = i + lSize
        Loop
        Buffer = Space(lSize)
        'get the bytes that will be sent over the winsock
        Get .BLNVar, , Buffer
        'send the bytes
        frmMain.wsDCCSend_DataArrival WinSockIndex, .bytesSent
    End With
End Sub
Public Sub DccSend_DataReceived(ByVal WinSockIndex As Integer, ByVal Data As String) 'RCV: NO
    Dim lSize As Long
    Dim fLength As Long
    Dim Buffer As String
    Dim intNewProgress As Integer
    Dim DCCIndex As Integer
    
    DCCIndex = GetDCCFileIndexFromWsIndex(WinSockIndex, False)
    
    With DCCFiles(DCCIndex)
        lSize = 1024
        fLength = LOF(.BLNVar)
        If .bytesSent > fLength Then
            frmMain.wsDCCSend.Item(WinSockIndex).SendData vbNullString
            Exit Sub
        End If
        If .bytesSent + lSize > fLength Then
            lSize = fLength - .bytesSent
        End If
        Buffer = Space(lSize)
        Get .BLNVar, , Buffer
        frmMain.wsDCCSend.Item(WinSockIndex).SendData Buffer
        .bytesSent = .bytesSent + lSize
        If Not .Hidden Then
            intNewProgress = Int((.bytesSent / Val(.FileSize)) * 100)
            If intNewProgress <> .Progress Then
                 'On Error GoTo closedtab
                .Progress = intNewProgress
                Tabs.Tabs.Item(.TabIndex).Caption = Language(209) & " " & .FileNameShort & "(" & .Progress & "%)"
                If CurrentActiveServer Is Me Then
                    frmMain.UpdateTabsBar
                    If GetDCCFromTab(Tabs.SelectedItem.index, False) = DCCIndex Then
                        frmMain.buildStatus
                    End If
                End If
            End If
        End If
        If .bytesSent > Val(.FileSize) Then
            Wait 0.1
            Close .BLNVar
            frmMain.wsDCCSend.Item(WinSockIndex).Close
            'File XX sent successfully 127.0.0.1
            frmMain.AddStatus Replace(Replace(Language(716), "%1", .FileNameShort), "%2", .IPAddress) & vbNewLine, Me
        End If
        Exit Sub
ClosedTab:
        Close .BLNVar
        frmMain.wsDCCSend.Item(WinSockIndex).Close
    End With
End Sub
Public Function GetDCCTabIndexFromTab(ByVal TabIndex As Integer, ByVal RCV As Boolean) As Integer
    Dim i As Integer
    
    On Error Resume Next
    If RCV Then
        For i = 0 To UBound(DCCFilesRCV)
            If DCCFilesRCV(i).TabIndex = TabIndex Then
                GetDCCTabIndexFromTab = i
                Exit Function
            End If
        Next i
    Else
        For i = 0 To UBound(DCCFiles)
            If DCCFiles(i).TabIndex = TabIndex Then
                GetDCCTabIndexFromTab = i
                Exit Function
            End If
        Next i
    End If
End Function
Public Function GetDCCChatsIndexFromWsIndex(ByVal WinSockIndex As Integer) As Integer
    Dim i As Integer
    
    For i = 1 To UBound(DCCChats)
        If DCCChats(i).WinSockIndex = WinSockIndex Then
            GetDCCChatsIndexFromWsIndex = i
            Exit Function
        End If
    Next i
    GetDCCChatsIndexFromWsIndex = -1
End Function
Public Function GetDCCChatsTabIndexFromWsIndex(ByVal WinSockIndex As Integer) As Integer
    Dim i As Integer
    
    For i = 1 To UBound(DCCChats)
        If DCCChats(i).WinSockIndex = WinSockIndex Then
            GetDCCChatsTabIndexFromWsIndex = DCCChats(i).TabIndex
            Exit Function
        End If
    Next i
    GetDCCChatsTabIndexFromWsIndex = -1
End Function
Public Function GetChanID(ByVal ChanName As String) As Long
    Dim i As Integer
    For i = 1 To Tabs.Tabs.Count
        If Strings.LCase$(Tabs.Tabs(i).Caption) = Strings.LCase$(ChanName) Then
            If Not TabType(i) = TabType_WebSite Then
                GetChanID = TabInfo(i)
                Exit Function
            End If
        End If
    Next i
    GetChanID = 0 'No such channel
End Function
Public Function GetTabFromNick(ByVal Nickname As String) As Integer
    Dim i As Integer
    For i = 1 To Tabs.Tabs.Count
        If LCase$(Tabs.Tabs.Item(i).Caption) = LCase$(Nickname) Then
            If TabType(i) = TabType_Private Then
                GetTabFromNick = i
                Exit Function
            End If
        End If
    Next i
    GetTabFromNick = -1
End Function
Public Function GetDCCFromSize(ByVal Size As Long) As Integer
    Dim i As Integer
    For i = 0 To UBound(DCCFiles)
        If DCCFiles(i).FileSize = Size Then
            GetDCCFromSize = i
            Exit Function
        End If
    Next i
End Function
Public Function GetEmptyChannelID() As Integer
    Dim i As Integer
    Dim intEmpty As Integer
    
    intEmpty = 1
    For i = 1 To TabInfo.Count
        If TabInfo(i) > intEmpty Then
            intEmpty = TabInfo(i)
        End If
    Next i
    intEmpty = intEmpty + 1
    GetEmptyChannelID = intEmpty
End Function
Public Sub RealDataArrival_Proxy(ByVal bytesTotal As Long)
    'this procedure handles proxy data
    'and establishes a SOCKS connection
    
    Dim strProxyData As String
    Dim bVersion As String
    Dim bMethod As String
    Dim bPortUpper As Byte
    Dim bPortLower As Byte
    Dim bReply As Byte
    Dim strBND_ADDRS As String
    Dim intBND_PORT As Integer
    Dim intActiveServerIndex As Integer
    Dim strIP As String
    Dim strAIP() As String
    Dim bIP(3) As Byte
    Static strStaticProxyData As String
       
    'Note that...
    'An octet is 8 bits, and is more or less synonymous with byte.
    
    If CurrentActiveServer.bProxyPhase = 2 Then
        'this must be a method selection reply
        'as we just sent a method selection request
        
        '( recieved method selection reply )
        'read the packet
        WinSockConnection.GetData strProxyData, , bytesTotal
        strStaticProxyData = strStaticProxyData & strProxyData
        
        'the reply should look like this:
        '                         +----+--------+
        '                         |VER | METHOD |
        '                         +----+--------+
        '                         | 1  |   1    |
        '                         +----+--------+
        
        'check to see if we recieved a fragmented pack
        'or if we have the full packet
        
        'we will need two bytes
        If Len(strStaticProxyData) = 2 Then
            'recieved correctly;
            'check to see if our method is
            'accepted by the SOCKS server
            'get the server's SOCKS version number
            bVersion = AscW(Strings.Left$(strStaticProxyData, 1))
            bMethod = AscW(Strings.Right$(strStaticProxyData, 1))
            If bVersion <> 5 Then
                'incompatible version number
                'we can only use the SOCKS5 protocol
                'display warning
                frmMain.AddStatus Language(590) & vbNewLine, Me
                'and close connection
                WinSockConnection.Close
                Exit Sub
            End If
            If bMethod <> 0 Then
                'the server uses an incompatible
                'SOCKS method
                'currently we can only handle
                'method 0 (no authentication required)
                'display warning
                frmMain.AddStatus Language(591) & vbNewLine, Me
                'close connection
                WinSockConnection.Close
                Exit Sub
            End If
            'great; if we reached this point
            'it means that our protocol is compatible
            'with the SOCKS server...
            'we're moving to phase 3
            
            'remove parsed data from buffer
            strStaticProxyData = vbNullString
        Else
            'we didn't recieve a full method reply
            'yet... exit sub and check back when we
            'get the rest of the bytes
            'TO DO: Timeout?
            Exit Sub
        End If
        'entering phase 3...
        m_bProxyPhase = 3
        
        'as the protocols are compatible
        'we can now request to start a connection
        'with the IRC server on the other side
        'This requires us to...
        
        'Send a SOCKS request
        
        '   The SOCKS request is formed as follows:
        '
        '        +----+-----+-------+------+----------+----------+
        '        |VER | CMD |  RSV  | ATYP | DST.ADDR | DST.PORT |
        '        +----+-----+-------+------+----------+----------+
        '        | 1  |  1  | X'00' |  1   | Variable |    2     |
        '        +----+-----+-------+------+----------+----------+

        'evaluate the octects to send defining the DST.PORT
        IntToUpperLower Port, bPortUpper, bPortLower
        
        strIP = mdlNetwork.GetIPAddress(HostName)
        If LenB(strIP) <> 0 Then
            strAIP = Split(strIP, ".")
            bIP(0) = strAIP(0)
            bIP(1) = strAIP(1)
            bIP(2) = strAIP(2)
            bIP(3) = strAIP(3)
            strIP = ChrW$(bIP(0)) & ChrW$(bIP(1)) & ChrW$(bIP(2)) & ChrW$(bIP(3))
        Else
            frmMain.AddStatus Replace(Language(926), "%1", HostName), Me
            WinSockConnection.Close
        End If
        'form the request and send it
                                   'VER=5   CMD=CNCT RSV      ATYP=IP  IP      PORT/u            PORT/l
        WinSockConnection.SendData ChrW$(5) & ChrW$(1) & ChrW$(0) & ChrW$(1) & strIP & ChrW$(bPortUpper) & ChrW$(bPortLower)
        
        'which basically means...
        '"We are using SOCKS version *5* and we want to *connect*
        ' to the server with the following *domain name* and port:
        ' *(domain)* *(port)*"
        
        'sent SOCKS request
        m_bProxyPhase = 4
    ElseIf m_bProxyPhase = 4 Then
        'this must be a SOCKS reply, as we just
        'sent a SOCKS request
        
        'get the packet
        WinSockConnection.GetData strProxyData, , bytesTotal
        
        'append it to our buffer
        strStaticProxyData = strStaticProxyData & strProxyData
        
        'the message must be at least 7 octets long
        If Len(strStaticProxyData) >= 7 Then
            'the message seems to be OK.
            'as we don't use an authentication
            'method, the reply won't be encapsulated
            '
            'TO DO: Check to see if we succeeded (DONE)
            '
            'The reply must be formed as follows:
            '        +----+-----+-------+------+----------+----------+
            '        |VER | REP |  RSV  | ATYP | BND.ADDR | BND.PORT |
            '        +----+-----+-------+------+----------+----------+
            '        | 1  |  1  | X'00' |  1   | Variable |    2     |
            '        +----+-----+-------+------+----------+----------+
            
            'analyse it
            bVersion = AscW(Strings.Left$(strStaticProxyData, 1))
            bReply = AscW(Strings.Mid$(strStaticProxyData, 2, 1))
            intBND_PORT = UpperLowerToInt( _
                          AscW(Strings.Mid$(strStaticProxyData, Len(strStaticProxyData) - 1, 1)), _
                          AscW(Strings.Mid$(strStaticProxyData, Len(strStaticProxyData), 1)))
            strBND_ADDRS = Strings.Mid$(strStaticProxyData, 5, Len(strStaticProxyData) - 6)
            
            'recieved SOCKS reply
            m_bProxyPhase = 5
            
            'we don't need to check if the version is compatible(we checked that before)
            
            'look at bReply to see if the operation succeeded
            If bReply = 0 Then
                'yep! success :)
                'let the user know that we have connected to the IRC server
                frmMain.AddStatus Replace(Replace(Language(593), "%1", HostName), "%2", Options.ProxyIP) & vbNewLine, Me
                'success, we have a connection over SOCKS
                m_bProxyPhase = 6
                'so we have an IRC connection
                'fire the IRC Connect "event"
                RealConnect
            Else
                'SOCKS failure
                'display the error message
                frmMain.AddStatus Language(577 + bReply) & vbNewLine, Me
                'close SOCKS connection
                WinSockConnection.Close
                Exit Sub
            End If
        Else
            'we have not recieved the hole message yet
            'TO DO: Timeout?
            'exit sub, and we'll get back later to
            'handle the hole message; when we recieve
            'the rest of it!
            Exit Sub
        End If
    End If
    
    'TO DO: Test!
End Sub

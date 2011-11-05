Attribute VB_Name = "mdlNode"
'Project NodeIRC(node) hosted by SourceForge (http://sourceforge.net)
'
'Please, visit our website
'http://node.sourceforge.net
'
'The project site is:
'http://sourceforge.net/projects/node
'
'The current developers are(alphabetically):
'ch-world
'dionyziz
'jnfoot
'mozillagodzilla
'nano
'
'Should you have any comments, questions, bug reports or feature requests, please contact us.
'
'Other people that have contributed to Node:
'See http://node.sourceforge.net/link.php?p=contributors
'
'
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'Allow only declared variables to be used.
Option Explicit

Public Const VERSION_CODENAME As String = "" ' <-- 0.36

Public Const NickList_DefaultSize = "*,100" 'the default size of the nicklist frame

'Node HTML recognition characters; these will later be replaced by < and >
'(before the message appears in a window)
Public Const HTML_OPEN As String = "" 'HtmlOpen = ChrW$(27)
Public Const HTML_CLOSE As String = "" 'HtmlOpen = ChrW$(29)

'this constant is added at the beginning of important data of events
'such as who took an action, what mode was set, what is the new topic, etc.
Public Const REASON_PREFIX As String = " " & HTML_OPEN & "span class=""node_reason""" & HTML_CLOSE
'this constant is added at the end of the reason data
Public Const REASON_SUFFIX As String = " " & HTML_OPEN & "/span" & HTML_CLOSE
'these constants is added at the beginning and end of each event text
Public Const EVENT_PREFIX As String = " " & HTML_OPEN & "span class=""node_event""" & HTML_CLOSE
Public Const EVENT_SUFFIX As String = " " & HTML_OPEN & "/span" & HTML_CLOSE
'these constants is added at the beginning and end of notices and error messages
Public Const IMPORTANT_PREFIX As String = HTML_OPEN & "span class=""node_important""" & HTML_CLOSE
Public Const IMPORTANT_SUFFIX As String = HTML_OPEN & "/span" & HTML_CLOSE
'these constants are added at the beginning and end of `DCC File Recieved' message and the buddy messages
Public Const SPECIAL_PREFIX As String = HTML_OPEN & "span class=""node_special""" & HTML_CLOSE
Public Const SPECIAL_SUFFIX As String = " " & HTML_OPEN & "/span" & HTML_CLOSE

'HTML Line break
Public Const HTML_BR As String = HTML_OPEN & "br" & HTML_CLOSE

'Editors
Public Const VBS_EDITOR = """$app\misc\ScriptEditor\sedit.exe"" $file"
Public Const DIALOG_EDITOR = """$app\misc\ScriptEditor\sedit.exe"" $file"
Public Const TXT_EDITOR = """notepad.exe"" $file"

'`mIRC format' Constants
'These characters are used by mIRC.
'CTCP: To indicate a CTCP request or reply the hole message is included in two of these characters
'Sample Request:
'VERSION
Public Const MIRC_CTCP As String = "" 'CTCP Request or Reply = ChrW$(1)
'For a message to appear in bold, it is included in the bold characters:
'This is some bold text
Public Const MIRC_BOLD As String = "" 'Bold = ChrW$(2)
'Some colored text is included in two color characters, the first followed by one or two numbers seperated by comma:
'1,2This is some colored text
'The first number is the forecolor whereas the second is the background color.
'The numbers are between 0 and 15.
Public Const MIRC_COLOR As String = "" 'Color = ChrW$(3)
'For a message to appear underlined, it is included in the underline characters:
'This is some underlined text
Public Const MIRC_UNDERLINE As String = "" 'Underline = ChrW$(31)
Public Const MIRC_ITALIC As String = ""

Public Const TabImage_Channel = 1 'tab-image index for channels is 1
Public Const TabImage_Private = 2 'for privates it's 2
Public Const TabImage_Status = 3 'where as the status tab 3
Public Const TabImage_WebSite = 4 'a website has image index 4
Public Const TabImage_DCC = 5 'Finally, a DCC has image index 5
Public Const TabImage_Look = 6 'To highlight a tab we'll use index 6

'NickList
Public Type NodeNickList
    List As Collection 'stores the nickList itself
    Size As String 'stores the width of the nicklist
    Topic As String 'stores the topic of the nicklist's channel
    Topic_Parsed As String 'same as Topic, but after CreateMainText() has been applied to it
    Modes As String 'stores the modes in the channels that are set ( + )
End Type

'IRC Actions
Public Enum NodeAction
    ndPrivMsg = 1 'someone is sending a message either to a channel we are in or privately to us
    ndJoin 'someone joins a channel we are in, or we join a new channel
    ndPart 'someone leaves a channel we are in, or we leave a channel
    ndQuit 'someone disconnects from the network we are connected to
    ndKick 'someone is kicked from one of the channels we are in, we kicked someone, or we were kicked by someone
    ndMode 'someone changes the modes of a channel we are in, or we changed the modes of a channel
    ndError 'an error caused us to disconnect from the server
    ndNames 'when joining a channel this event fills the nicklist of it
    ndNick 'someone that is in a channel we are in changes his/her nickname or we changed our nickname
    ndUseAltNick 'the nickname the user chose is in use; we are going to use an alternative one
    ndTopic 'the topic in a channel we are in changes(either by us or by another user) or we are informed about the topic when we join a channel
    ndTopicTime 'if we are being informed about the topic in a particular channel this action lets us know who set the topic and when he did it
    ndNotice 'we recieved a notice from the server
    ndInvite 'someone invited us to join a channel.
    ndChannelList 'this event occurs when using /list command to list the channels of the server we are connected to
    ndBanList 'this event occurs when using /mode #chan +b (without any arguments) to list the bans in a channel
    ndSpecialNotice 'some text describing a `special event', an event that is not in another category, has to appear in a channel
    ndBuddyName 'saving the real name of the buddy to his/her profile file
    ndNamesSpecial 'allows /names to be used
    ndIson 'reply to an /ison command
    ndInitModes 'initial modes when we enter a channel
    
    ndMemoUnread
End Enum

Public Type NodeOptions
    'loggins options
    LogChannels As Boolean 'are channel messages being logged?
    LogPrivates As Boolean 'are private message being logged?
    LogByNetwork As Boolean 'seperate logs by network?
    LogRAW As Boolean 'log incoming RAW data
       
    EnablePerform As Boolean 'are we going to execute the commands in perform uppon connect?
    PerformSingle As Boolean 'are we using the same perform for all servers?
    
    LanguageFile As String 'what is the language file used
    
    BuddyEnterMSG As Boolean 'shall we display a message when a buddy enters?
    BuddyLeaveMSG As Boolean 'shall we display a message when a buddy leaves?
    BuddyEnterWIN As Boolean 'shall we pop-up a window when a buddy enters?
    BuddyLeaveWIN As Boolean 'shall we pop-up a window when a buddy leaves?
    
    EnableScripting As Boolean 'is Node Scripting enabled?
    EnableCodeBehind As Boolean 'is CodeBehind feature for Skins enabled?
    
    HTMLLoading As Boolean 'is loading HTML displayed?
    HTMLError As Boolean 'is error HTML displayed?
    
    SessionN As Byte 'session handling on normal startup
    SessionC As Byte 'session handling after crash
    
    DCCOptionsB As Byte 'dcc for buddies
    DCCOptionsI As Byte 'dcc for ignored nicks
    DCCOptionsE As Byte 'dcc for everyone
    DCCAntivirus As String 'executable of the antivirus
    AutoNDC As Boolean 'try to automatically connect via NDC when you open up an NDC window?
    
    KeepTrayRunning As Boolean 'keep Node running when the user closes it using `X'.
    RestoreStatus As Boolean 'restore status after restarting the program?
    JoinPanel As Boolean 'show Join Panel when connected?
    CheckLatest As Boolean 'should the program check for latest version on startup?
    FadeTransaction As Boolean 'is fade transaction used on startup or when unloading the program?
    InfoTips As Boolean 'show infotips?
    TOD As Boolean 'show Tips of the Day?
    AutoComplete As Boolean 'does the user want to enable Autocomplete when he/she presses Tab?
    XPCommonControls As Boolean 'are we going to use XP common controls?(only available to XP users)
    StartPage As Boolean 'start with a web site tab
    StartPageURL As String 'the URL of the startup web site tab
    FocusJoined As Boolean 'focus the channel's tab when the user joins a channel
    
    QuitMultiple As Boolean
    QuitFile As String
    QuitMsg As String 'the quit message send to the server before disconnecting
    
    
    UseProxy As Boolean 'are we using a SOCKS proxy?
    ProxyIP As String 'hostname or ip of the proxy we are using
    ProxyPort As Long 'proxy port
    
    'ctcp options
    CTCPPing As Boolean 'enable CTCP ping
    CTCPPingToIgnored As Boolean 'if true, pings from ignored nicknames are ignored
    CTCPVersion As Boolean 'enable CTCP version
    CTCPVersionToIgnored As Boolean 'if true, versions from ignored nicknames are ignored
    CTCPVersionCustomize As Boolean 'has the user set a custom CTCP version message?
    CTCPVersionMessage As String 'if yes, that's the message
    CTCPTime As Boolean 'enable CTCP time
    CTCPTimeToIgnored As Boolean 'if true, times from ignored nicknames are ignored
    CTCPFloodProtect As Boolean 'enable CTCP flood protection
    CTCPFloodBounce As Boolean 'use CTCP-bounce to protect from CTCP flood?
    
    'timestamp options
    TimeStamp As Boolean 'show timestamps?
    TimeStampStatus As Boolean '...in statuses?
    TimeStampChannels As Boolean '...in channels?
    TimeStampPrivates As Boolean '...in privates?
    TimeStampLogs As Boolean 'log timestamps that are visible?
    
    'accessibility options
    Narration As Boolean 'enable narration? [needs restart]
    NarrationInterface As Boolean 'narrate interface events
    NarrationChannels As Boolean 'narrate channels' text
    NarrationPrivates As Boolean 'narrate privates' text
    NarrationStatus As Boolean 'narrate statuses' text
    
    SelectionCopy As Boolean 'copy when the user selects some text
    SelectionClear As Boolean 'un-select the selected text when the user selects it
    
    NickLinkChan As Boolean
    NickLinkPriv As Boolean
    NickLinkMineChan As Boolean
    NickLinkMinePriv As Boolean
    
    StartupConnect As Boolean 'connect to a server on startup
    StartupConnectHostname As String 'if yes, the hostname to connect to
    StartupConnectPort As Long 'if yes, the port to connect to
    
    ConnectRetry As Boolean 'enable connection retry?
    ConnectRetryDelay As Integer 'how many seconds do we have to wait before retrying a connection? default = 20
    
    JoinOnInvite As Boolean 'automatically join a channel when invited?
    JoinOnKick As Boolean 'automatically rejoin a channel when kicked?
    ModesOnJoin As Boolean 'get channel modes on join?
    
    KeepChannelsOpen As Boolean 'keep channels open after disconnecting?

    SoundScheme As String 'the .xml filename of the sound scheme that is currently selected (only contains filename, not path)
    SmileyPack As String 'the .xml filename of the smiley pack that is currently selected (only contains filename, not path)
    UseSmileys As Boolean 'are smileys enabled?
    
    'away system
    AwayEnabled As Boolean
    AwayMinutes As Integer
    AwayNick As Boolean
    AwayNickStr As String
    AwayStatus As Boolean
    AwayStatusID As Byte
    AwayPerform As Boolean
    AwayPerformStr As String
    AwayBackNick As Boolean
    AwayBackNickStr As String
    AwayBackStatus As Boolean
    AwayBackStatusID As Boolean
    AwayBackPerform As Boolean
    AwayBackPerformStr As String
    
    IdentEnable As Boolean
    IdentPort As Boolean
    IdentDefaultUserID As Boolean
    IdentDefaultOSID As Boolean
    IdentCustomUserID As String
    IdentCustomOSID As String
    
    ParseMemoServ As Boolean
    
    'browsing options
    BrowseParseLinks As Boolean
    BrowseInternalBrowser As Boolean
    
    'display options
    DisplayNormal As String
    DisplaySkinFontSize As String
End Type

Public Type DCCFile
    FileName As String
    FileNameShort As String
    FileSize As Double
    TabIndex As Integer
    WinSockIndex As Integer
    Progress As Byte
    bytesSent As Long
    BLNVar As Integer
    IPAddress As String
    Resuming As Integer '1-the file sent will be resumed 0-No
    Port As Integer
    StartTime As Long 'Secs
    ETL As Long 'Secs
    Speed As Double 'KB/Sec
    Hidden As Boolean
End Type

Public Type DCCChatINFO
    UserName As String
    WinSockIndex As Integer
    Port As Integer
    TabIndex As Integer
End Type

Public Type NDCConnection
    intVersion As Integer
    bTimeZone As Byte
    intTCP As Integer
    strNicknameA As String
    strNicknameB As String
    aData As String
    IntroPackSent As Boolean
    IntroPackRecieved As Boolean
    Typing As Boolean
    TypingTime As Long
    TypingSent As Long
    RemoteStatus As NodeStatus
    AudioRequested As Boolean
    AudioSentTime As Long
    AudioFile As String
    AudioConnected As Boolean
    AudioTCP As Integer
    AvatarToSend As String
    UserAvatar As String
    ActiveServer As clsActiveServer
    MMNetMeetingStatus As Byte
End Type

Public Type HiddenDCCConnection
    AllowedNickname As String
    AllowedFileName As String
    AllowedIP As String
    EventID As Integer
    NDCConnectNUM As Integer
    WriteFileName As String
End Type

Public Type NodeCTCP
    strType As String
    strNickname As String
    lngTime As Long
    TheServer As clsActiveServer
End Type

Public Type NodeProtection
    bReason As Byte
    strNickname As String
    lngTime As Long
    TheServer As clsActiveServer
End Type

Public Type ndChannelListEntry
    Channel As String
    Users As Integer
    Topic As String
End Type

Public Enum NodeInfoTips
    WelcomeToNode = 1
    Connected
    Joined
    PrivateMessage
    DCCIncoming
    Kicked
    CrashDis
    Invitation
    NickInUse
    LangChange
    SkinChange
    CrashAsk
    TrayExit
    BuddySignOn
    BuddySignOff
End Enum

Public Enum NodeStatus
    Status_Online
    Status_Away
    Status_BRB
    Status_Sleeping
    Status_Busy
    Status_Lunch
    Status_Phone
    Status_Vacation
    Status_Not_Home
    Status_Not_At_Office
    Status_Not_At_Desk
End Enum

Public Type PendingNDCConnectionRequest
    Nickname As String
    ActiveServer As clsActiveServer
End Type

Public Enum SecureLockIconConstants
    secureLockIconUnsecure = 0
    secureLockIconMixed = 1
    secureLockIconUnknownBits = 2
    secureLockIcon40Bit = 3
    secureLockIcon56Bit = 4
    secureLockIconFortezza = 5
    secureLockIcon128Bit = 6
End Enum

'Following is used for checking user's already open ports
Public Type MIB_TCPROW
    dwState As Long        'state of the connection
    dwLocalAddr As String * 4    'address on local computer
    dwLocalPort As Long    'port number on local computer
    dwRemoteAddr As String * 4   'address on remote computer
    dwRemotePort As String * 4   'port number on remote computer
End Type

Public Type MIB_TCPTABLE
    dwNumEntries As Long    'number of entries in the table
    mibtable(100) As MIB_TCPROW   'array of TCP connections
End Type

Public Declare Function GetTcpTable Lib "IPhlpAPI" _
  (pTcpTable As MIB_TCPTABLE, pdwSize As Long, Border As Long) As Long

Public PendingNDCConnectionRequests() As PendingNDCConnectionRequest
Public CurrentConnections() As Byte
Public NM As Object
Public LocalIP As String
Public LocalLookupNick As String

'Tab Types
Public Const TabType_Channel = 0 'Channel Tab
Public Const TabType_Private = 1 'Private Tab
Public Const TabType_WebSite = 2 'WebSite Tab
Public Const TabType_Status = 3 'Status Tab
Public Const TabType_DCCFile = 4 'DCC File Progress Tab

'Script Sphere Variables
Public ScriptNick As String 'nickname who took that last action
Public ScriptNick2 As String 'nickname who took part in the last action
Public ScriptChan As String 'the channel where the last action happened
Public ScriptThisChan As String 'the selected channel

'the Splash Screen
Public SplashScreen As Form

'FileSystemObject; we finally decided to make it public as it is used many times inside the program
Public FS As FileSystemObject 'object variable used to access the filesystem of local disks.

'Language Array
Public Language(1000) As String 'array that contains the captions of the selected language
Public LanguageMulti(5) As New CSparseMatrix
Public LangMultiNames() As String
Public LangCharSet As Integer

Public IRCMsg(50, 1) As String

'Debug Class
Public DB As clsDebug

'The currently applied skin
Public ThisSkin As NodeSkin
'The active smiley pack
Public ThisSmileyPack As nodeSmileyPack

Public Options As NodeOptions 'this variable stores almost all the options the user has set using the options dialog

Public Restarting As Boolean 'if some certain options change the program needs to restart; if we are restarting this variable is set to True and then back to False.
Public Ending As Boolean

Public bCrash As Boolean 'did Node crash the last time it runned?

Public ActiveServers() As clsActiveServer
Public CurrentActiveServer As clsActiveServer
'class to store buddy list info to
Public AllBuddies() As clsIdentity

Public MSSpeech As Object

Public PastCTCPs() As NodeCTCP
Public ProtectedFrom() As NodeProtection

Public NDCConnections() As NDCConnection
Public AllowedHidden() As HiddenDCCConnection

Public Quitting As Boolean

Public StartTime As Long

Public IsAway As Boolean

Public GlobalLinkID As Long
Private PortArray() As Long
'API32
'API Shell
Public Const SW_SHOW = 5

'API Constants
'this is the style we want
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
'this is the constant to determine the the value passed to SetLayered is an alpha value
Private Const LWA_ALPHA = &H2

'used to execute commands, just like VB's Shell command
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'API Timer
'used to get the internal system's timer value
Public Declare Function GetTickCount Lib "kernel32" () As Long

'Init XP Controls
'used to make the program display XP-style controls
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()

'WinSock API
Public Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

'Function to set the window style(to make it able to be transparent)
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

'function to get the window long using the handle of the window
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

'the actual function to set the transparency of the window
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long


Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, _
                        ByVal bInvert As Long) As Long

'Windows Blending, (not used - we use our own method)
Public Const AW_BLEND = &H80000
Public Declare Function AnimateWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

'Lock window for some time while we update so it doesn't look ugly
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'mci used for NDC audio recording
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Sub Main()
    Dim strSkinFile As String
    Dim i As Long
    Dim UsageTime As Long
    
    'Startup Sub
    'This sub is executed when the program starts
    
    'Load debug class
    Set DB = New clsDebug
    
    ReDim CurrentConnections(3, 0)
    
    DB.Enter "Main"
    
    DB.X "Loading FileSystemObject"
    'Load FS; this object variable is used to access the local filesystem later
    Set FS = New FileSystemObject
    
    ReDim LoadedPlugin(0)
    DB.X "Listing PlugIns"
    ListPlugIns
    
    'Load Language
    'read what language file we are using from the registry
    DB.X "Loading Language"
    Options.LanguageFile = GetSetting("Node", "Options", "LanguageFile", vbNullString)
    If LenB(Options.LanguageFile) = 0 Then
        Options.LanguageFile = App.Path & "\data\languages\english.lang"
        SaveSetting "Node", "Options", "LanguageFile", Options.LanguageFile
    End If
    
    'if the file exists...
    If FS.FileExists(Options.LanguageFile) Then
        '...load the language file
        LoadLanguage Options.LanguageFile
    
    'if it doesn't exist
    Else
        'load the english language file
        LoadLanguage App.Path & "\data\languages\english.lang"
    End If
    
    'load plugins' settings
    For i = 0 To NumToPlugIn.Count - 1
        Plugins(i).boolLoadOnStartup = GetSetting(App.EXEName, "Plugins", Right$(Plugins(i).strName, Len(Plugins(i).strName) - Len("prjPlugIn")), False)
    Next i
        
    'set the options for codeBehind and Scripting
    'we need to load these settings now, as they
    'are necessary to see if we have to execute
    'method "Begin" and the current skin's CodeBehind
    Options.EnableCodeBehind = GetSetting(App.EXEName, "Options", "CodeBehind", True)
    Options.EnableScripting = GetSetting(App.EXEName, "Options", "Scripting", True)
    
    UsageTime = GetSetting("Node", "Remember", "UsageTime", 0)
    If UsageTime > 86400000 Then
        If GetSetting("Node", "Remember", "FeedbackMsg", False) = False Then
            SaveSetting "Node", "Remember", "FeedbackMsg", True
            If MsgBox(Language(758), vbYesNo Or vbQuestion, Language(759)) = vbYes Then
                xShell "http://node.sourceforge.net/link.php?p=feedback """"", 0
            End If
        End If
    End If
    StartTime = GetTickCount
    
    DB.X "Checking for Crash"
    bCrash = FS.FileExists(App.Path & "\temp\session.xml")
    If bCrash Then
        DB.X "There was a crash"
        FS.CopyFile App.Path & "\temp\session.xml", App.Path & "\temp\crash.xml"
    End If
    
    Options.Narration = GetSetting(App.EXEName, "Options\Accessibility\Narration", "Enabled", False)
    If Options.Narration Then
        DB.X "Loading MS Speech"
        On Error GoTo SAPI_Runtimes_Not_Installed
        Set MSSpeech = CreateObject("SAPI.SpVoice.1")
        MSSpeech.Speak Language(15)
    End If
    
    'get the current skin
    DB.X "Loading Skin"
    strSkinFile = GetSetting("Node", "Options", "Skin", vbNullString)
    If LenB(strSkinFile) = 0 Then
        strSkinFile = App.Path & "\data\skins\default.skin"
        SaveSetting "Node", "Options", "Skin", strSkinFile
    End If
    ThisSkin = LoadSkin(strSkinFile)
    
    'Load the Splash Screen
    'we'll need to store it in an object variable(SplashScreen)
    'in order to be able to modify the "Loading" caption
    Set SplashScreen = New frmCustom
    
    'On Error GoTo Splash_Error
    DB.X "Loading Splash Screen"
    LoadDialog App.Path & "/data/skins/" & ThisSkin.SplashDialog, SplashScreen

    '...refresh view
    SplashScreen.Refresh
    On Error GoTo LoadDefaultSkin
    SplashScreen.lblCustom(1).Caption = Language(31) '"Loading XP Support..."
    SplashScreen.lblCustom(1).Refresh
    'load windows XP support; this will only work if the file node.exe.manifest exists in the application's folder
    DB.X "InitCommonControls"
    InitCommonControls
      
    'create "real" pseudo-random numbers
    Randomize
    
    ReDim ProtectedFrom(0)
        
    'change the text displayed on the Splash Screen
    SplashScreen.lblCustom(1).Caption = Language(217) '"Loading Skin..."
    'and refresh the view
    SplashScreen.lblCustom(1).Refresh
    'apply the current skin
    'Note: this will load frmOptions and frmMain,
    '      so how can they be loaded again later?
    DB.X "Applying Skin"
    ApplySkin ThisSkin
    
    SplashScreen.lblCustom(1).Caption = Language(163) 'Loading Options
    SplashScreen.lblCustom(1).Refresh
    
    'load the options dialog
    DB.X "Loading frmOptions"
    Load frmOptions
    'and load the options set by the user
    DB.X "Loading All Options"
    frmOptions.LoadAll
        
    SplashScreen.lblCustom(1).Caption = Language(18) 'Loading Interface
    SplashScreen.lblCustom(1).Refresh
    
    'Load Main Window
    'TO DO:
    ' frmMain has already been loaded by
    ' the Skin Loader. The following line
    ' is not really necessary.
    Load frmMain
        
    SplashScreen.SetFocus
    
    'starting plugins
    If Not Restarting Then
        DB.X "Loading Startup Plugins"
        For i = 0 To NumToPlugIn.Count - 1
            If Plugins(i).boolLoadOnStartup Then
                SplashScreen.lblCustom(1).Caption = Replace(Language(628), "%1", Right$(Plugins(i).strName, Len(Plugins(i).strName) - Len("prjPlugIn"))) 'Loading Plugin
                SplashScreen.lblCustom(1).Refresh
                LoadPlugIn Plugins(i).strName & ".dll"
            End If
        Next i
    End If
    
    'execute LoadingCompleted CodeBehind procedure
    DB.X "Executing CodeBehind Script `Skin_LoadingCompleted'"
    SkinExecuteCodeBehind "Skin_LoadingCompleted"
    
    If Not frmMain.boolRestoring Then
        'and show the options dialog
        'frmOptions.Show vbModal
        'show the connect panel
        DB.X "Loading Connect Panel"
        frmMain.LoadPanel "connect"
    End If
        
    'and hide the splash screen
    SplashScreen.Hide
    DB.Leave "Main"
        
    Exit Sub
Splash_Error:
    'there was a splash screen error
    DB.Leave "Main", "Invalid Skin/Splash Screen"
    MsgBox Language(244), vbCritical, Language(166)
    SaveSetting App.EXEName, "Options", "Skin", App.Path & "/data/skins/default.skin"
    End
SAPI_Runtimes_Not_Installed:
    If MsgBox(Language(731), vbYesNo Or vbQuestion, Language(721)) = vbYes Then
        xShell "http://activex.microsoft.com/activex/controls/sapi/spchapi.exe """"", 0
        End
    Else
        Resume Next
    End If
LoadDefaultSkin:
    MsgBox Language(247), vbCritical, Language(166)
    SaveSetting App.EXEName, "Options", "Skin", App.Path & "\data\skins\default.skin"
    MsgBox "Please restart Node", vbInformation
    End
End Sub
Public Function NewServer(Optional ByVal Activate As Boolean = True, Optional ByVal boolStartup As Boolean = False) As clsActiveServer
    'Sub NewServer()
    'Loads a new server
    'and lets the user
    'connect to it
    
    Dim TheServer As clsActiveServer
    Dim TheServerIndex As Integer
    
    DB.Enter "NewServer"
    
    DB.X "Creating new ActiveServer"
    'create the new server
    Set TheServer = New clsActiveServer
    
    DB.X "Calculating TheServerIndex"
    'create a new index on "ActiveServers"
    'so that the new server is accessible
    'from the rest of the code
    TheServerIndex = UBound(ActiveServers) + 1
    
    DB.X "TheServerIndex = " & TheServerIndex
    
    DB.X "ReDimming ActiveServers() to the new TheServerIndex"
    ReDim Preserve ActiveServers(TheServerIndex)
    
    DB.X "Storing new ActiveServer into ActiveServers"
    'refer from "ActiveServers" to our
    'newly created server
    Set ActiveServers(TheServerIndex) = TheServer
    
    DB.X "AddStatus, first call"
    'add welcome text to the new status window
    frmMain.AddStatus Language(656) & HTML_BR, TheServer
    
    'if we were asked to "activate" the new
    'server...
    If Activate Then
        DB.X "Activating ActiveServer"
        'make it current
        ActiveServerMakeCurrent TheServer, boolStartup
    End If
    
    'set default server properties here...
    'TheServer.Connected = False
    
    DB.X "Referencing NewServer to TheServer"
    'return the new server
    Set NewServer = TheServer
    
    DB.X "Adding Server Node"
    Set TheServer.ServerNode = frmMain.tvConnections.Nodes.Add(frmMain.tvConnections.Nodes.Item(1), tvwChild, "s" & TheServerIndex, Language(773) & " " & TheServerIndex & " (" & Language(772) & ")", TabImage_Status)
    
    DB.X "Expanding Server Node"
    TheServer.ServerNode.Expanded = True
    
    DB.Leave "NewServer"
End Function
Public Sub ActiveServerMakeCurrent(ByRef TheServer As clsActiveServer, Optional ByVal boolStartup As Boolean = False)
    Dim i As Integer
    Dim ActiveServer As clsActiveServer
    
    DB.Enter "ActiveServerMakeCurrent"
    
    '...the selected server
    '   should be the passed one
    
    DB.X "Comparing CAS with TheServer"
    If Not CurrentActiveServer Is TheServer Then
        DB.X "Setting CAS to TheServer"
        Set CurrentActiveServer = TheServer
    End If
    
    DB.X "Enumerating ActiveServers and updating as necessary"
    For i = 0 To UBound(ActiveServers)
        DB.X "Server " & i
        Set ActiveServer = ActiveServers(i)
        If Not ActiveServer Is Nothing Then
            DB.X "(valid)"
            If ActiveServer.Tabs.Visible Then
                ActiveServer.Tabs.Visible = False
            End If
        End If
    Next i
    
    DB.X "Showing Server Tabs"
    If Not TheServer.Tabs.Visible Then
        TheServer.Tabs.Visible = True
        DB.X "Resizing"
        frmMain.Form_Resize
    End If
    
    If Not boolStartup Then
        DB.X "Building Status"
        frmMain.tsTabs_Click TheServer.Tabs.index
        frmMain.buildStatus
    End If

    DB.Leave "ActiveServerMakeCurrent"
End Sub
Public Sub DeleteServer(ByVal TheServer As clsActiveServer)
    Dim intServersCount As Integer
    Dim intServerIndex As Integer
    Dim asAnotherServer As clsActiveServer
    Dim i As Integer
    
    If TheServer.WinSockConnection.State = sckConnected Then
        Quit TheServer
    End If
    
    intServerIndex = GetServerIndexFromActiveServer(TheServer)
    For i = 0 To UBound(ActiveServers)
        If Not ActiveServers(i) Is Nothing Then
            intServersCount = intServersCount + 1
            If i <> intServerIndex Then
                Set asAnotherServer = ActiveServers(i)
            End If
        End If
    Next i
    If intServersCount <= 1 Then
        'cannot remove last server
        Err.Raise vbObjectError + 1, "DeleteServer()", "Can not delete last ActiveServer!"
    End If
    If TheServer Is CurrentActiveServer Then
        ActiveServerMakeCurrent asAnotherServer
    End If
    
    With frmMain.tvConnections
        For i = 1 To .Nodes.Count
            If .Nodes.Item(i).Key = "s" & GetServerIndexFromActiveServer(TheServer) Then
                .Nodes.Remove i
                Exit For
            End If
        Next i
    End With
    
    Set ActiveServers(intServerIndex) = Nothing
End Sub
Public Function ConnectionIsPresent() As Boolean
    Dim ActiveServer As Variant
    
    For Each ActiveServer In ActiveServers
        If Not ActiveServer Is Nothing Then
            If ActiveServer.WinSockConnection.State = sckConnected Or ActiveServer.WinSockConnection.State = sckClosing Then
                ConnectionIsPresent = True
                Exit Function
            End If
        End If
    Next ActiveServer
    
    ConnectionIsPresent = False
End Function
Public Sub Node_Unload()
    Dim fl As File
    For Each fl In FS.GetFolder(App.Path & "\temp").Files
        If fl.Name <> "session.xml" Then
            On Error Resume Next 'permission denied(?)
            fl.Delete
        Else
            If Options.SessionN <> 1 Then
                If FS.FileExists(App.Path & "\temp\normal.xml") Then
                    FS.DeleteFile App.Path & "\temp\normal.xml"
                Else
                    fl.Move App.Path & "\temp\normal.xml"
                End If
            Else
                fl.Delete
            End If
        End If
    Next fl
    Set frmMain.xpBalloon = Nothing
    SaveSetting "Node", "Remember", "UsageTime", (GetTickCount - StartTime) + GetSetting("Node", "Remember", "UsageTime", 0)
    End
End Sub
Public Function GetNick(ByVal strFullNick As String) As String
    'Function used to remove the leading priviledges symbols,
    'for voice and oper
    '      +   and  @
    If Strings.Left$(strFullNick, 3) = "+%@" Then
        GetNick = Strings.Right$(strFullNick, Len(strFullNick) - 3)
    ElseIf Strings.Left$(strFullNick, 2) = "+@" Or Strings.Left$(strFullNick, 2) = "%@" Or Strings.Left$(strFullNick, 2) = "+%" Then
        GetNick = Strings.Right$(strFullNick, Len(strFullNick) - 2)
    ElseIf Strings.Left$(strFullNick, 1) = "@" Or Strings.Left$(strFullNick, 1) = "%" Or Strings.Left$(strFullNick, 1) = "+" Then
        GetNick = Strings.Right$(strFullNick, Len(strFullNick) - 1)
    Else
        GetNick = strFullNick
    End If
End Function
Public Function GetLogFile(ByVal strID As String, ByRef TheServer As clsActiveServer)
    Dim strReturn As String
    Dim strNet As String
    Dim strFile As String
    
    If Not FS.FolderExists(App.Path & "\logs") Then
        MkDir App.Path & "\logs"
    End If
    If Options.LogByNetwork Then
        If LenB(TheServer.WinSockConnection.RemoteHost) = 0 Then
            strNet = vbNullString
        Else
            strNet = "\" & TheServer.WinSockConnection.RemoteHost
        End If
        If Not FS.FolderExists(App.Path & "\logs" & strNet) Then
            MkDir App.Path & "\logs" & strNet
        End If
    Else
        strNet = vbNullString
    End If
    strFile = strID
    'some characters that channels contain
    'cannot be used as filenames
    'and they must be replaced before
    'we can save anything
    strFile = Replace(strFile, "/", "_")
    strFile = Replace(strFile, "\", "_")
    strFile = Replace(strFile, ":", "_")
    strFile = Replace(strFile, "*", "_")
    strFile = Replace(strFile, "|", "_")
    strFile = Replace(strFile, "<", "_")
    strFile = Replace(strFile, ">", "_")
    strFile = Replace(strFile, """", "_")
    
    GetLogFile = App.Path & "\logs" & strNet & "\" & strFile & ".log.html"
End Function
Public Function GetNicklistNick(ByVal strFullNick As String) As String
    Dim strReturn As String
    'get the nickname as it should be displayed in a nicklist,
    'i.e. without a +@ but with either no priviledges, or @ or +
    If Strings.Left$(strFullNick, 3) = "+%@" Then
        GetNicklistNick = "@" & Strings.Right$(strFullNick, Len(strFullNick) - 3)
    ElseIf Strings.Left$(strFullNick, 2) = "+@" Or Strings.Left$(strFullNick, 2) = "%@" Then
        GetNicklistNick = "@" & Strings.Right$(strFullNick, Len(strFullNick) - 2)
    ElseIf Strings.Left$(strFullNick, 2) = "+%" Then
        GetNicklistNick = "%" & Strings.Right$(strFullNick, Len(strFullNick) - 2)
    Else
        GetNicklistNick = strFullNick
    End If
End Function
Public Function CreateServersList(Optional ByVal boolOrganizeServerDialog = False) As String
    Dim intFL As Integer
    Dim strServer As String
    Dim strServerHostname As String
    Dim strServerPort As String
    Dim strServerDescription As String
    Dim strDescriptionQuotes As String
    Dim strServerList As String
    
    'build the server list
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As #intFL
        Do Until EOF(intFL)
            Line Input #intFL, strServer
            strServerHostname = GetStatement(strServer)
            strServerPort = GetParameter(strServer)
            strServerDescription = GetParameter(strServer, 2)
            strDescriptionQuotes = IIf(InStr(1, strServerDescription, " "), """", vbNullString)
            If Not boolOrganizeServerDialog Then
                strServerList = strServerList & "<acronym title='" & Replace(Language(523), "%1", strServerHostname & ":" & strServerPort) & "'>" & _
                                "<a href='NodeScript:/connect " & _
                                strServerHostname & " " & strServerPort & " " & strDescriptionQuotes & _
                                strServerDescription & strDescriptionQuotes & "'>" & _
                                strServerDescription & "</a></acronym><br>"
            Else
                strServerList = strServerList & "<option>" & strServerDescription & "</option>"
            End If
        Loop
        If boolOrganizeServerDialog Then
            If LenB(strServerList) > 0 Then
                strServerList = "<select multiple size=""10"" id=""server_list"" onChange=""selection_change();"">" & _
                                strServerList & "</select>"
            End If
        End If
        Close #intFL
    End If
    
    If LenB(strServerList) = 0 Then
        strServerList = Language(522)
    End If
    CreateServersList = strServerList
End Function
Public Function ServerListCount() As Integer
    Dim i As Integer
    Dim intFL As Integer
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As #intFL
        Do Until EOF(intFL)
            i = i + 1
            xLineInput intFL
        Loop
        Close #intFL
    End If
    ServerListCount = i
End Function
Public Sub ReadServer(ByVal ServerIndex As Integer, ByRef OutDescription As String, ByRef OutHostName As String, ByRef OutPort As Long)
    Dim i As Integer
    Dim intFL As Integer
    Dim strServer As String
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As #intFL
        Do Until EOF(intFL)
            i = i + 1
            Line Input #intFL, strServer
            If i = ServerIndex Then
                OutHostName = GetStatement(strServer)
                OutPort = GetParameter(strServer)
                OutDescription = GetParameter(strServer, 2)
                Close #intFL
                Exit Sub
            End If
        Loop
        Close #intFL
    End If
End Sub
Public Sub AlterServer(ByVal ServerIndex As Integer, Optional boolDoDelete As Boolean, Optional ByVal strDescription As String, Optional ByVal strHostname As String, Optional ByVal lngPort As Long)
    Dim i As Integer
    Dim intFL As Integer
    Dim intFL2 As Integer
    Dim strDescriptionQuotes As String
    Dim strServer As String
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Lock Write As #intFL
        intFL2 = FreeFile
        Open App.Path & "\conf\servers.tmp" For Output Access Write Lock Read Write As #intFL2
        Do Until EOF(intFL)
            i = i + 1
            Line Input #intFL, strServer
            If i = ServerIndex Then
                If Not boolDoDelete Then
                    strDescriptionQuotes = IIf(InStr(1, strDescription, " ") > 0, """", vbNullString)
                    Print #intFL2, strHostname & " " & lngPort & " " & strDescriptionQuotes & strDescription & strDescriptionQuotes
                End If
            Else
                Print #intFL2, strServer
            End If
        Loop
        Close #intFL2
        Close #intFL
        FS.DeleteFile App.Path & "\conf\servers.lst", True
        FS.MoveFile App.Path & "\conf\servers.tmp", App.Path & "\conf\servers.lst"
    End If
End Sub
Public Sub SortServers()
    Dim i As Integer
    Dim intFL As Integer
    Dim strServer As String
    Dim cServers As Collection
    Dim cDescriptions As Collection
    
    Set cServers = New Collection
    Set cDescriptions = New Collection
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Lock Write As #intFL
        Do Until EOF(intFL)
            Line Input #intFL, strServer
            cServers.Add strServer
            cDescriptions.Add GetParameter(strServer, 2)
        Loop
        Close #intFL
        SortCollection cServers, cDescriptions
    
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Output Access Write Lock Write As #intFL
        i = 1
        Do Until i > cServers.Count
            Print #intFL, cServers.Item(i)
            i = i + 1
        Loop
        Close #intFL
    End If
End Sub
Public Sub MoveServer(ByVal ServerIndex As Integer, Optional ByVal boolDirectionUp As Boolean = True)
    Dim i As Integer
    Dim intFL As Integer
    Dim intFL2 As Integer
    Dim strServer As String
    Dim strServer2 As String
    
    If FS.FileExists(App.Path & "\conf\servers.lst") And (ServerIndex > 1 Or Not boolDirectionUp) Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Lock Write As #intFL
        intFL2 = FreeFile
        Open App.Path & "\conf\servers.tmp" For Output Access Write Lock Read Write As #intFL2
        Do Until EOF(intFL)
            i = i + 1
            Line Input #intFL, strServer
            If (i = ServerIndex - 1 And boolDirectionUp) Or (i = ServerIndex And Not boolDirectionUp) Then
                Line Input #intFL, strServer2
                Print #intFL2, strServer2
                Print #intFL2, strServer
                i = i + 1
            Else
                Print #intFL2, strServer
            End If
        Loop
        Close #intFL2
        Close #intFL
        FS.DeleteFile App.Path & "\conf\servers.lst", True
        FS.MoveFile App.Path & "\conf\servers.tmp", App.Path & "\conf\servers.lst"
    End If
End Sub
Public Sub ClearCTCPMemory()
    Dim i As Integer
    Dim g As Integer
    Dim NewCTCPs() As NodeCTCP
    
    For i = 0 To UBound(PastCTCPs)
        If PastCTCPs(i).lngTime > GetTickCount - 60000 Then
            'not too old;
            'it shouldn't expire: keep
            ReDim Preserve NewCTCPs(g)
            NewCTCPs(g).strType = PastCTCPs(i).strType
            NewCTCPs(g).strNickname = PastCTCPs(i).strNickname
            NewCTCPs(g).lngTime = PastCTCPs(i).lngTime
            g = g + 1
        End If
    Next i
    
    ReDim PastCTCPs(g - 1)
    For i = 0 To UBound(PastCTCPs)
        PastCTCPs(i).lngTime = NewCTCPs(i).lngTime
        PastCTCPs(i).strNickname = NewCTCPs(i).strNickname
        PastCTCPs(i).strType = NewCTCPs(i).strType
    Next i
End Sub
Public Sub ClearProtectionMemory()
    Dim i As Integer
    Dim g As Integer
    Dim NewProtection() As NodeProtection
    
    For i = 0 To UBound(ProtectedFrom)
        If ProtectedFrom(i).lngTime <= GetTickCount - 60000 Then
            If LenB(ProtectedFrom(i).strNickname) > 0 Then
                'protection expired
                frmMain.AddStatus Replace(Language(645), "%1", ProtectedFrom(i).strNickname) & vbNewLine, ProtectedFrom(i).TheServer
            End If
        Else
            'protection not expired
            ReDim Preserve NewProtection(g)
            NewProtection(g).bReason = ProtectedFrom(i).bReason
            NewProtection(g).strNickname = ProtectedFrom(i).strNickname
            NewProtection(g).lngTime = ProtectedFrom(i).lngTime
            g = g + 1
        End If
    Next i
    
    If g = 0 Then
        ReDim ProtectedFrom(0)
    Else
        ReDim ProtectedFrom(g - 1)
        For i = 0 To UBound(ProtectedFrom)
            ProtectedFrom(i).lngTime = NewProtection(i).lngTime
            ProtectedFrom(i).strNickname = NewProtection(i).strNickname
            ProtectedFrom(i).bReason = NewProtection(i).bReason
        Next i
    End If
End Sub
Public Sub CheckCTCPFlood()
    Dim lngRequestIndex As Long
    Dim lngFlooderIndex As Long
    Dim lngFloodersCount As Long
    Dim lngProtectionIndex As Long
    Dim strNicknames() As String
    Dim TheServers() As clsActiveServer
    Dim lngFloodCount() As Long
    
    ReDim strNicknames(0)
    ReDim lngFloodCount(0)
    For lngRequestIndex = 0 To UBound(PastCTCPs)
        'check to see if there
        'is another entry for the current
        'nickname
        For lngFlooderIndex = 0 To UBound(strNicknames)
            'there is
            If LCase$(strNicknames(lngFlooderIndex)) = LCase$(PastCTCPs(lngRequestIndex).strNickname) Then
                'add one to the total number
                'of CTCP requests
                lngFloodCount(lngFlooderIndex) = lngFloodCount(lngFlooderIndex) + 1
                GoTo Next_I
            End If
        Next lngFlooderIndex
        'there isn't
        'create an entry and
        'set the total CTCP requests count
        'to one
        ReDim Preserve strNicknames(lngFloodersCount)
        ReDim Preserve TheServers(lngFloodersCount)
        strNicknames(lngFloodersCount) = PastCTCPs(lngRequestIndex).strNickname
        lngFloodCount(lngFloodersCount) = 1
        lngFloodersCount = lngFloodersCount + 1
        Set TheServers(lngFloodersCount) = PastCTCPs(lngRequestIndex).TheServer
Next_I:
    Next lngRequestIndex
    
    For lngFlooderIndex = 0 To UBound(strNicknames)
        If lngFloodCount(lngFlooderIndex) > 8 Then
            'this user has sent more than
            'eight CTCP requests within
            'one minute, which seems
            'to be a flood.
            'check to see if we are already protected from this user
            For lngProtectionIndex = 0 To UBound(ProtectedFrom)
                If LCase$(ProtectedFrom(lngProtectionIndex).strNickname) = LCase$(strNicknames(lngFlooderIndex)) And _
                    ProtectedFrom(lngProtectionIndex).bReason = 0 Then
                    'protection is already enabled for this user
                    'update time of protection
                    '(protect for one more minute)
                    ProtectedFrom(lngProtectionIndex).lngTime = GetTickCount 'now
                    Exit Sub
                End If
            Next lngProtectionIndex
            
            'we aren't protected from this user
            'display message
            frmMain.AddStatus Replace(Language(643), "%1", strNicknames(lngFlooderIndex)) & " " & Language(644) & vbNewLine, TheServers(lngFlooderIndex)
            
            'create a new protection
            lngProtectionIndex = UBound(ProtectedFrom) + 1
            ReDim Preserve ProtectedFrom(lngProtectionIndex)
            ProtectedFrom(lngProtectionIndex).bReason = 0 'ICMP flood
            ProtectedFrom(lngProtectionIndex).lngTime = GetTickCount 'now
            ProtectedFrom(lngProtectionIndex).strNickname = strNicknames(lngFlooderIndex)
        End If
    Next lngFlooderIndex
End Sub
Public Function CompleteWord(ByVal strWord As String, Optional ByVal boolFillSuggestions As Boolean = False) As String
    'Complete Nick or Word when the user presses Tab
    'return the completed word
    Dim lstCurrentNickList As Collection 'the nicklist collection for the active channel
    Dim i As Integer 'a counter variable for the loops
    Dim intFL As Integer 'file index variable
    Dim strCurrentWord As String 'a possible word we read from complete.dat or from another resource
        
    If Left$(strWord, 2) = vbNewLine Then
        strWord = Right$(strWord, Len(strWord) - 2)
    End If
    
    If LenB(strWord) = 0 Then
        Exit Function
    End If
    
    'if we're in a channel,
    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Channel Then
        'It's a channel; see if the nickname is somewhere inside the current nicklist
        Set lstCurrentNickList = CurrentActiveServer.NickList_List(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index))
        For i = 1 To lstCurrentNickList.Count
            strCurrentWord = GetNick(lstCurrentNickList.Item(i))
            If strCurrentWord <> strWord Then
                If WordMatch(strWord, strCurrentWord) Then
                    If LenB(CompleteWord) = 0 Then
                        CompleteWord = strCurrentWord
                    End If
                    If Not boolFillSuggestions Then
                        Exit Function
                    Else
                        frmMain.lstSuggestions.AddItem strCurrentWord
                    End If
                End If
            End If
        Next i
    ElseIf CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
        'It's a private message tab; see if the user is typing his buddy's nickname
        If WordMatch(strWord, CurrentActiveServer.Tabs.SelectedItem.Caption) Then
            strCurrentWord = CurrentActiveServer.Tabs.SelectedItem.Caption
            If strCurrentWord <> strWord Then
                If LenB(CompleteWord) = 0 Then
                    CompleteWord = strCurrentWord
                End If
                If Not boolFillSuggestions Then
                    Exit Function
                Else
                    frmMain.lstSuggestions.AddItem strCurrentWord
                End If
            End If
        End If
    End If
    
    'Word not found in the nicklist or wasn't the name of the private contact;
    'maybe it's a channel we're in
    For i = 1 To CurrentActiveServer.TabType.Count
        If CurrentActiveServer.TabType(i) = TabType_Channel Then
            strCurrentWord = CurrentActiveServer.Tabs.Tabs.Item(i).Caption
            If strCurrentWord <> strWord Then
                If WordMatch(strWord, strCurrentWord) Then
                    If LenB(CompleteWord) = 0 Then
                        CompleteWord = strCurrentWord
                    End If
                    If Not boolFillSuggestions Then
                        Exit Function
                    Else
                        frmMain.lstSuggestions.AddItem strCurrentWord
                    End If
                End If
            End If
        End If
    Next i
    
    'it's not a channel
    'is it a common word, then?
    'get a free file index
    intFL = FreeFile
    'open the data file which contains the words
    Open App.Path & "\data\complete.dat" For Input As #intFL
    'go through the file
    Do Until EOF(intFL)
        'read a word
        Line Input #intFL, strCurrentWord
        If Strings.Left$(strCurrentWord, 1) = ">" Then
            'remove escape character
            strCurrentWord = Strings.Right$(strCurrentWord, Len(strCurrentWord) - 1)
        End If
        
        'if it is something and not a comment
        If Not (LenB(strCurrentWord) = 0 Or Strings.Left$(strCurrentWord, 1) = "#") Then
            If strWord <> strCurrentWord Then
                'if it matches the word the user started typing
                If WordMatch(strWord, strCurrentWord) Then
                    'return the match
                    If LenB(CompleteWord) = 0 Then
                        CompleteWord = strCurrentWord
                    End If
                    If Not boolFillSuggestions Then
                        Close #intFL
                        'we don't need to return something else
                        Exit Function
                    Else
                        frmMain.lstSuggestions.AddItem strCurrentWord
                    End If
                End If
            End If
        End If
    Loop
    Close #intFL
        
    'no suggestions
    'CompleteWord = ""
End Function
Public Sub EditScript(FileName As String, Optional useDialog As Boolean = False)
    'Sub used to load the script editor to edit a script
    
    'execute the editor
    Shell App.Path & "\misc\ScriptEditor\sedit.exe """ & FileName & """", vbNormalFocus
End Sub
Public Sub LoadLanguage(ByVal FileName As String)
    'This sub is used to load the FileName language-file into the array Language
    Dim intFL As Integer 'file index
    Dim strCurrentLine As String 'the current line of the file
    Dim strStatement As String
    Dim i As Integer 'the key index for each language caption
    Dim objFile As File
    Dim objTextStream As TextStream
    Dim strLanguage As String
    Dim strTemp As String

    'get a free file index
    intFL = FreeFile
    'open the language file
    'Set objFile = fs.GetFile(FileName)
    'Set objTextStream = objFile.OpenAsTextStream(ForReading)
    strLanguage = DecodeFile(FileName, False)
    'Open FileName For Binary Access Read Shared As #intFl
    'skip the first two lines
    'strCurrentLine = objTextStream.ReadLine & objTextStream.ReadLine
    strCurrentLine = DecodedLineInput(strLanguage) & DecodedLineInput(strLanguage)
    i = 2
    
    'go through the file
    Do Until Len(strLanguage) = 0
        i = i + 1
        
        'get a line from the file(either a key or a comment)
        strCurrentLine = DecodedLineInput(strLanguage)
        'strCurrentLine = objTextStream.ReadLine
        'if it has a leading escape, remove it
        If Strings.Left$(strCurrentLine, 1) = ">" Then
            strCurrentLine = Strings.Right$(strCurrentLine, Len(strCurrentLine) - 1)
        End If
        'if it's not a comment add the current key to the array
        '                                        & ChrW$(0)
        If Strings.Left$(strCurrentLine, 1) <> "#" And Len(strCurrentLine) <> 0 Then
            'this line is not a comment
            'error trap for language file bugs
            On Error GoTo Language_File_Error
            'key indexes can't be unicode
            strStatement = GetStatement(strCurrentLine) ', vbFromUnicode)
            If IsNumeric(strStatement) Then
                'TO DO
                'Language(strStatement) = GetParameter(StrConv(strCurrentLine, vbFromUnicode))
                'Language(strStatement) = GetParameter(strCurrentLine, , , True)
                Language(strStatement) = Replace(GetParameter(strCurrentLine, , , False), ChrW$(0), vbNullString)
            Else
                If strStatement = "@" Then
                    'setting char-set
                    'the charset index can't be unicode: it's a simple number
                    LangCharSet = GetParameter(strCurrentLine) 'vbFromUnicode))
                End If
            End If
        End If
    Loop
EOF_Return:
    'objTextStream.Close
    Close #intFL
    Exit Sub
Language_File_Error:
    If Err.Number = 62 Then
        'EOF
        Resume EOF_Return
    End If
    MsgBox "The language file " & FileName & " has an error in line " & i, vbCritical, "Language File Error"
End Sub
Public Function CreateMainText(ByVal strText As String, ByRef TheServer As clsActiveServer, Optional ByVal UseSmileys As Boolean = True, Optional ByVal HighlightMe As Boolean = False) As String
    'Everything is passed in this function before it is displayed
    'replaces smileys and
    'mIRC bold/underline symbols with the HTML ones
    'In addition, it converts < and > to the HTML equilants
    'and the two ASCII characters to < and >.
    'Contains code for HTML Security Issues fixes.
    'Code for highlighting lines
    'Everything is passed only once here.
    '
    '              Hopefully changing i from integer to long will speed this HORRIBLY SLOW function up...
    Dim i As Long 'a counter variable for the loops and a temporary integer variable
    Dim strReturn As String 'the string we are going to return
    Dim strHREF As String 'the target URL for links
    Dim strAddition As String 'the string we need to add, e.g we will need to add <b> if there's an mIRC bold character there.
    Dim CurFile As Integer 'file index variable for the smileys file
    Dim strTemp As String, strTemp2 As String, strTemp3 As String 'temporary string variables
    Dim strATemp() As String
    Dim bTemp As Byte 'a temporary byte variable
    Dim inStrong As Boolean 'are we current in a bold text?
    Dim inUnderline As Boolean 'in an underlined?
    Dim inItalic As Boolean 'in an italic?
    Dim inColor As Boolean 'in a colored?
    Dim inTag As Boolean 'in an HTML tag?
    Dim intTagNest As Integer
    Dim intFL As Integer
    Dim intTemp As Integer
    Dim intTemp2 As Integer
    
    'DB.Enter "CreateMainText"
    
    'add a space to the end of the line, it won't be displayed
    strReturn = strText & " "
    
    'if this line has the ability to be highlighted
    
    'Seems to be fixed now
    'DB.Enter "Nick highlighting"
    
    If HighlightMe Then
        'if we aren't the one speaking
        If Strings.Left$(Strings.LCase$(strReturn), Len("&lt; " & Strings.LCase$(TheServer.myNick) & " &gt;")) <> Strings.LCase$("&lt; " & TheServer.myNick & " &gt;") Then
            i = InStr(1, strReturn, " &gt;")
            'only execute the next things if strReturn is a message
            If i > 0 Then
            
                'get the position of our nickname in the line(if any)
                'and ignore everything in front of the real text (e.g. the nick)
                i = InStr(InStr(1, strReturn, " &gt;"), strReturn, TheServer.myNick)          'if it exists in the line
                If i > 0 Then
                    'if our nickname exists in the line(as a whole word)
                    If i = 1 Or InStr(i - 1, strReturn, " " & TheServer.myNick) > 0 Or InStr(i - 1, Strings.LCase$(strReturn), ">" & Strings.LCase$(TheServer.myNick)) Then
                        'highlight the line
                        strReturn = SPECIAL_PREFIX & strReturn & SPECIAL_SUFFIX
                    End If
                End If
            End If
        End If
    End If
    
    'DB.Leave "Nick highlighting"
    'DB.Enter "Smileys replacing"
    
    If Options.UseSmileys Then
        'Replace smileys
        ':) -> <img src="......\graphics\smileys\icon_smile.gif">
        'if smileys are on...
        If UseSmileys Then
            For i = 1 To UBound(ThisSmileyPack.AllSmileys)
                'create a smiley using img()
                'change the < and > symbols for the HTML tags
                'to HTML_OPEN and HTML_CLOSE which will
                'be replaced later
                strReturn = Replace(strReturn, ThisSmileyPack.AllSmileys(i).ShortcutText, _
                            Replace(Replace(img(App.Path & "/data/smileys/" & Options.SmileyPack & "/" & ThisSmileyPack.AllSmileys(i).FileName, ThisSmileyPack.AllSmileys(i).ShortcutText), "<", HTML_OPEN), ">", HTML_CLOSE))
            Next i
        End If
    End If
    
    'DB.Leave "Smileys replacing"
    'DB.Enter "lt/gt/br"
    
    '< -> &lt;
    '> -> &gt;
    'this displays < and > correctly
    strReturn = Replace(strReturn, "<", " &lt; ")
    strReturn = Replace(strReturn, ">", " &gt; ")

    'Return -> <BR>
    strReturn = Replace(strReturn, vbLf, vbLf & " <BR>")
    
    'Replace ChrW$(2) with <b> and </b>
    'Replace ChrW$(31) with <u> and </u>
    'Replace ChrW$(22) with <i> and </i>
    
    'On Error GoTo DEBUG_CMT_NOW
    'DB.Leave "lt/gt/br"
    'DB.Enter "bold/underline/colors"
    
    i = 0
    Do: i = i + 1
        If i > Len(strReturn) Then
            Exit Do
        End If
        'if we are "inside" an HTML tag
        If intTagNest > 0 Then
            'if the current text starts with on
            If Strings.Mid$(LCase$(strReturn), i, 2) = "on" Then
                'it could be an HTML event
                'get the twenty next characters
                'which should contain the hole
                'event text
                strTemp = Mid$(LCase$(strReturn), i, 20)
                'now we need to check to see
                'if it really is an HTML event
                'get a free file index
                intFL = FreeFile
                'open the data file which contains a list with all the possible HTML events
                Open App.Path & "\data\htmlevents.dat" For Input Access Read Shared As #intFL
                'go through it
                Do Until EOF(intFL)
                    'read a line
                    Line Input #intFL, strTemp2
                    'if it's not a comment
                    If Left$(strTemp2, 1) <> "#" Then
                        'does the event in the data file match the text one?
                        If Left$(LCase$(strTemp), Len(strTemp2)) = LCase$(strTemp2) Then
                            'yes. We need to disable this event
                            'change On to 0n
                            'This makes the event inactive
                            Mid$(strReturn, i, 2) = "0n" 'zero - n
                            'don't read any more events from the data file
                            Exit Do
                        End If
                    End If
                'next event/line in the data file
                Loop
                'close the data file
                Close #intFL
            End If
        End If
        'check to see if the current character is a
        'BOLD / UNDERLINE / ITALIC / COLOR
        'character
        Select Case Strings.Mid$(strReturn, i, 1)
            'bold character
            Case MIRC_BOLD
                'if the current text is already bold add </b>, else add <b>
                strAddition = HTML_OPEN & IIf(inStrong, "/", vbNullString) & "b" & HTML_CLOSE
                'add the HTML tag
                strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
                'if the text was bold, it's not now... If it wasn't it is now
                inStrong = Not inStrong
                'we move to the next character, just after the tag we added
                i = i + Len(strAddition) - 1
            'underline character
            Case MIRC_UNDERLINE
                strAddition = HTML_OPEN & IIf(inUnderline, "/", vbNullString) & "u" & HTML_CLOSE
                strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
                inUnderline = Not inUnderline
                i = i + Len(strAddition) - 1
            'italic character
            Case MIRC_ITALIC
                strAddition = HTML_OPEN & IIf(inItalic, "/", vbNullString) & "i" & HTML_CLOSE
                strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
                inItalic = Not inItalic
                i = i + Len(strAddition) - 1
            'a color character
            Case MIRC_COLOR
                'get five characters from the text after the color character
                strTemp = Strings.Mid$(strReturn, i + 1, 5) ' ##,##
                'get the first two characters from the text we just got
                strTemp2 = Strings.Left$(strTemp, 2)
                strTemp3 = vbNullString
                'assume that the color has two digits
                bTemp = 2
                'Isn't it a valid color number?
                If (Not (Strings.Trim$(Conversion.Str$(Val(strTemp2))) = strTemp2)) Or Val(strTemp2) > 15 Or Val(strTemp2) < 0 Then
                    'it is not
                    'it may be a zero-prefixed number
                    'example: 08
                    'check if it is
                    If Replace(Conversion.Str$(Val(Right$(strTemp2, 1))), " ", "0") = CStr(strTemp2) Then
                        'Stop
                    End If
                    If Left$(strTemp2, 1) = "0" Then
                        'Stop
                    End If
                    If Not (Val(strTemp2) > "9" Or Val(strTemp2) < "0") Then
                        'Stop
                    End If
                    If Left$(strTemp2, 1) = "0" And Replace(Conversion.Str$(Val(Right$(strTemp2, 1))), " ", "0") = strTemp2 And Not (Val(strTemp2) > "9" Or Val(strTemp2) < "0") Then
                        'yes, it is zero-prefixed
                        'remove the zero
                        strReturn = Strings.Left$(strReturn, i) & Strings.Right$(strTemp2, 1) & Strings.Right$(strTemp, 3) & Strings.Right$(strReturn, Len(strReturn) - i - Len("00,00"))
                        'get strTemp and strTemp2 again
                        'so that they are valid after the
                        'zero removal
                        strTemp = Strings.Mid$(strReturn, i + 1, 5) ' ##,##
                        strTemp2 = Strings.Left$(strTemp, 1)
                        'the color has only one digit
                        bTemp = 1
                    Else
                        'it's not zero-prefixed
                        'it may be a single number
                        'Example: [K]59
                        '( [K] is Ctrl + K )
                        'this means we start coloring
                        'with color 5 and there is
                        'some text after, starting with 9
                        'get the actual color code
                        strTemp2 = Strings.Left$(strTemp, 1)
                        'the color has only one digit
                        bTemp = 1
                        'if the first character is not numeric
                        If Not IsNumeric(strTemp2) Then
                            'it is a [K] character alone...
                            'no valid characters
                            bTemp = 0
                            strTemp2 = vbNullString
                        End If
                    End If
                End If
                'if we have a valid color code
                If LenB(strTemp2) > 0 Then
                    'if just after the color code there's a comma
                    If Strings.Mid$(strTemp, Len(strTemp2) + 1, 1) = "," Then
                        'assume we have a background color code
                        'with two digits
                        bTemp = bTemp + 3
                        'get the two characters that construct the possible background color
                        strTemp3 = Strings.Mid$(strTemp, Len(strTemp2) + 2, 2)
                        'if they are invalid...
                        If (Not (Strings.Trim$(Conversion.Str$(Val(strTemp3))) = strTemp3)) Or Val(strTemp3) > 15 Or Val(strTemp3) < 0 Then
                            'it may be a zero-prefixed background code
                            'is it?
                            If Conversion.Val(Left$(strTemp3, 1)) = 0 And (Replace(Conversion.Str$(Val(Right$(strTemp3, 1))), " ", 0) = strTemp3) And Not (Conversion.Val(strTemp3) > 9 Or Conversion.Val(strTemp3) < 0) Then
                                'yes. Remove the zero
                                strReturn = Strings.Left$(strReturn, i + Len(strTemp2) + 1) & Strings.Right$(strTemp3, 1) & Strings.Right$(strTemp, 5 - Len(strTemp2) - 3) & Strings.Right$(strReturn, Len(strReturn) - i - Len("00,00") - Len(strTemp2) + 1)
                                strTemp3 = Strings.Mid$(strTemp, Len(strTemp2) + 3, 1)
                                bTemp = bTemp - 1
                            Else
                                'no. Check to see if it's a single number
                                bTemp = bTemp - 1
                                strTemp3 = Strings.Mid$(strTemp, Len(strTemp2) + 2, 1)
                                'isn't it?
                                If Not IsNumeric(strTemp3) Then
                                    bTemp = bTemp - 2
                                    strTemp3 = vbNullString
                                End If
                            End If
                        End If
                    End If
                End If
                'if there are some valid color code characters
                If bTemp > 0 Then
                    'if we already are inside a color tag
                    If inColor Then
                        'close it first
                        strAddition = HTML_OPEN & "/font" & HTML_CLOSE
                        strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - Len(vbNullString) + 1)
                        i = i + Len(strAddition)
                    End If
                    'we are going to create a new color tag
                    inColor = True
                    'create the tag. Get the HTML color text from the color code number
                    strAddition = HTML_OPEN & "font color=""" & MIRCColorToHTML(strTemp2) & """"
                    'if there's a background color
                    If Len(strTemp3) > 0 Then
                        'add a background style attribute
                        strAddition = strAddition & " style=""background-color:" & MIRCColorToHTML(strTemp3) & """"
                    End If
                    'and close the tag
                    strAddition = strAddition & HTML_CLOSE
                    'add the tag
                    '(ommiting - Len("K") + 1    'K = Ctrl + K)
                    strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - bTemp)
                'there are no valid color codes after the K
                Else
                    'if we are inside a colored text
                    If inColor Then
                        'close it
                        strAddition = HTML_OPEN & "/font" & HTML_CLOSE
                        strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i + 1)
                        i = i + Len(strAddition)
                    End If
                    'we are no longer inside a colored text
                    inColor = False
                    'don't add anything
                    strAddition = vbNullString
                    'just ommit the K
                    '(ommiting - Len("K") + 1    'K = Ctrl + K)
                    strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i)
                End If
                'we have ommited the K character(Len = 1)
                'but added an HTML color tag(Len = Len(strAddition))
                'move to the next character, just after the tag we've added
                i = i + Len(strAddition) - 1
            'HTML <
            Case HTML_OPEN
                'Somebody may wrongly nest two html tags:
                '<A HREF="<b>nested tag</b>">Test</a>
                'we will not count it as a tag closing
                'but we will count the tag nesting
                intTagNest = intTagNest + 1
            'HTML >
            Case HTML_CLOSE
                intTagNest = intTagNest - 1
            Case vbLf
                'line break
                'if the current text is bold
                If inStrong Then
                    'make it non-bold
                    strReturn = strReturn & HTML_OPEN & "/b" & HTML_CLOSE
                End If
                If inUnderline Then
                    strReturn = strReturn & HTML_OPEN & "/u" & HTML_CLOSE
                End If
                If inItalic Then
                    strReturn = strReturn & HTML_OPEN & "/i" & HTML_CLOSE
                End If
                If inColor Then
                    strReturn = strReturn & HTML_OPEN & "/font" & HTML_CLOSE
                End If
        End Select
    Loop
    
    'DB.Leave "bold/underline/colors"

'DEBUG_CMT_NOW:
'    Stop
'    Resume
    'DB.Enter "links"
    
    If Options.BrowseParseLinks Then
        'http://... -> <A HREF="...">http://...</a>
        i = 0
        'go through the text again
        Do: i = i + 1
            'if we've reached the end of the text, stop seeking
            If i > Len(strReturn) Then Exit Do
            'if the current position in text is before a valid internet protocol tag:
            'http:// or ftp:// or https://
            If Strings.Mid$(strReturn, i, Len("http://")) = "http://" Or _
                Strings.Mid$(strReturn, i, Len("ftp://")) = "ftp://" Or _
                Strings.Mid$(strReturn, i, Len("https://")) = "https://" Then
                
                'if there is a character before the current one
                If i > 1 Then
                    'check to see if it is a quotation mark
                    If Strings.Mid$(strReturn, i - 1, 1) = """" Then
                        'it is; do not make the text a link
                        GoTo Continue
                    End If
                End If
                'get the next space, < and > characters
                intTemp = InStr(i, strReturn, " ")
                intTemp2 = InStr(i, strReturn, HTML_OPEN)
                If intTemp2 < intTemp And intTemp2 > 0 Then
                    intTemp = intTemp2
                End If
                intTemp2 = InStr(i, strReturn, HTML_CLOSE)
                If intTemp2 < intTemp And intTemp2 > 0 Then
                    intTemp = intTemp2
                End If
                'intTemp contains the position of the next space, < or > character
                'get the text that is going to be a link
                strHREF = Strings.Mid$(strReturn, i, intTemp - i)
                'create the HTML tag
                'add tooltip "Click here to visit this web site" using <acronym>
                strAddition = HTML_OPEN & "acronym title=""" & Language(521) & """" & HTML_CLOSE & HTML_OPEN & "a href=""" & strHREF & """" & HTML_CLOSE & _
                    strHREF & HTML_OPEN & "/a" & HTML_CLOSE & HTML_OPEN & "/acronym" & HTML_CLOSE
                'add the tag into the text
                strReturn = Strings.Left$(strReturn, i - 1) & strAddition & Strings.Right$(strReturn, Len(strReturn) - i - Len(strHREF) + 1)
                'move to the next character
                '( -1 because we are going to add one at the beginning of the loop)
                i = i + Len(strAddition) - 1
            End If
Continue:
        Loop
    End If
    
    'DB.Leave "links"
    
    'DB.Enter "S.I."
    strReturn = Replace(strReturn, HTML_OPEN & "script", "&lt;script")
    strReturn = Replace(strReturn, HTML_OPEN & "iframe", "&lt;iframe")
    strReturn = Replace(strReturn, HTML_OPEN & "ol", "&lt;ol")
    strReturn = Replace(strReturn, HTML_OPEN & "ul", "&lt;ul")
    strReturn = Replace(strReturn, HTML_OPEN & "input", "&lt;input")
    strReturn = Replace(strReturn, HTML_OPEN & "embed", "&lt;embed")
    strReturn = Replace(strReturn, HTML_OPEN & "object", "&lt;object")
    
    'ChrW$(27) = <           ASCII = ^ + [
    'ChrW$(29) = >           ASCII = ^ + ]
    strReturn = Replace(strReturn, HTML_OPEN, "<")
    strReturn = Replace(strReturn, HTML_CLOSE, ">")
    
    'DB.Leave "S.I."
    'return the parsed text
    CreateMainText = strReturn
    
    'DB.Leave "CreateMainText"
End Function
Public Function CreateNarrationText(ByVal Text As String) As String
    'Remove all HTML tags
    
    Dim strReturn As String
    Dim intTextLen As Integer
    Dim bThisCharacter As Byte
    Dim boolInHTML As Boolean
    Dim i As Integer
    
    intTextLen = Len(Text)
    Do
        i = i + 1
        If i > intTextLen Then
            Exit Do
        End If
        bThisCharacter = AscW(Mid$(Text, i, 1))
        If bThisCharacter = AscW(HTML_OPEN) Then
            boolInHTML = True
        ElseIf bThisCharacter = AscW(HTML_CLOSE) Then
            boolInHTML = False
        Else
            If Not boolInHTML Then
                strReturn = strReturn & ChrW$(bThisCharacter)
            End If
        End If
    Loop
    
    CreateNarrationText = strReturn
End Function
'Public Function GetServerIndexFromWSIndex(ByVal Index As Integer) As Integer
'    Dim i As Integer
'
'    For i = 0 To UBound(ActiveServers)
'        If Not ActiveServers(i).WinSockConnection Is Nothing Then
'            If ActiveServers(i).WinSockConnection.Index = Index Then
'                GetServerIndexFromWSIndex = i
'                Exit Function
'            End If
'        End If
'    Next i
'
'    GetServerIndexFromWSIndex = -1
'End Function
Public Function NickLink(ByVal strNick As String, Optional ByVal SimpleHTML As Boolean = False) As String
    Dim strReturn As String
    
    strReturn = "<a class=""nick"" href=""NodeScript:/nickmenu " & GetNick(strNick) & """>" & _
                strNick & "</a>"
    If Not SimpleHTML Then
        NickLink = Replace(Replace(strReturn, "<", HTML_OPEN), ">", HTML_CLOSE)
    End If
End Function
Public Function GetServerIndexFromActiveServer(ByRef TheServer As clsActiveServer) As Integer
    Dim i As Integer
    
    For i = 0 To UBound(ActiveServers)
        'compare ConnectTime with ConnectTime
        'This comparison doesn't need a lot of
        'resources
        If Not ActiveServers(i) Is Nothing Then
            If ActiveServers(i).ConnectTime = TheServer.ConnectTime Then
                'if the comparison above is true, it
                'is very likely that we are talking
                'about the same server. (unless both aren't connected)
                'We are going to make it absolutely sure
                'by comparing the WinSock objects;
                'this comparison needs more time
                'to complete and that's why we
                'didn't directly check that
                If ActiveServers(i).WinSockConnection Is TheServer.WinSockConnection Then
                    GetServerIndexFromActiveServer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    GetServerIndexFromActiveServer = -1
End Function
Public Function GetServerFromDCCChatWsIndex(ByVal WinSockIndex As Integer) As Integer
    Dim i As Integer
    
    For i = 0 To UBound(ActiveServers)
        If Not ActiveServers(i) Is Nothing Then
            If ActiveServers(i).GetDCCChatsIndexFromWsIndex(WinSockIndex) > -1 Then
                GetServerFromDCCChatWsIndex = i
                Exit Function
            End If
        End If
    Next i
End Function
Public Sub RememberChannel(ByVal strChannel As String)
    Dim intFL As Integer
    Dim intFL2 As Integer
    Dim strTemp As String
    intFL = FreeFile
    Open App.Path & "\conf\channels.lst" For Input Access Read Shared As #intFL
    intFL2 = FreeFile
    Open App.Path & "\conf\channels.tmp" For Output Access Write Lock Write As #intFL2
    Print #intFL2, strChannel
    Do Until EOF(intFL)
        Line Input #intFL, strTemp
        If LCase$(strTemp) <> strChannel Then
            Print #intFL2, strTemp
        End If
    Loop
    Close #intFL
    Close #intFL2
    Kill App.Path & "\conf\channels.lst"
    Name App.Path & "\conf\channels.tmp" As App.Path & "\conf\channels.lst"
End Sub
Public Function GetPastChannels(Optional ByVal Limit As Integer = 7) As String
    Dim intFL As Integer
    Dim strResult As String
    Dim strTemp As String
    Dim i As Integer
    intFL = FreeFile
    Open App.Path & "\conf\channels.lst" For Input Access Read Shared As #intFL
    Do Until EOF(intFL) Or i > Limit
        i = i + 1
        Line Input #intFL, strTemp
        strResult = strResult & "<a href=""NodeScript:/join " & strTemp & """>" & strTemp & "</a><br>" & vbNewLine
    Loop
    Close #intFL
    If LenB(strResult) = 0 Then
        '"No Channels"
        strResult = "(" & Language(451) & ")<br>"
    End If
    GetPastChannels = strResult
End Function
Public Function GetFavWebs(Optional ByVal Limit As Integer = 15) As String
    Dim intFL As Integer
    Dim strResult As String
    Dim strTemp As String, strTemp2 As String
    Dim i As Integer
    intFL = FreeFile
    Open App.Path & "\conf\favwebs.lst" For Input Access Read Shared As #intFL
    Do Until EOF(intFL) Or i > Limit
        i = i + 1
        Line Input #intFL, strTemp
        Line Input #intFL, strTemp2
        strResult = strResult & "<acronym title=""" & strTemp2 & """>" & "<a href=""NodeScript:/browse " & strTemp2 & """>" & strTemp & "</a></acronym>&nbsp;&nbsp;&nbsp;" & "<acronym title=""" & Language(83) & """>" & "<a href=""NodeScript:/remove-fav " & strTemp & """><img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Delete & """ border=""0""></a></acronym><br>" & vbNewLine
    Loop
    Close #intFL
    If LenB(strResult) = 0 Then
        '"No Favorites"
        strResult = "(" & Language(484) & ")<br>"
    End If
    GetFavWebs = strResult
End Function
Public Function GetBuddies(Optional ByVal Limit As Integer = 100) As String
    Dim strResult As String
    Dim strTemp As String, strTemp2 As String
    Dim i As Integer, i2 As Integer
    
    For i = 0 To UBound(AllBuddies)
        If AllBuddies(i).isOnline = True Then
            For i2 = 0 To AllBuddies(i).ServerCount
                If i2 = AllBuddies(i).ServerCount Then
                    If LenB(AllBuddies(i).Servers(i2)) > 0 Then strTemp2 = strTemp2 & "<img src=""" & App.Path & "/data/skins/" & ThisSkin.TreeSubNodeImage & """ width=""16px"" height=""16px"" align=""left""> " & AllBuddies(i).Servers(i2) & "<br>"
                Else
                    If LenB(AllBuddies(i).Servers(i2)) > 0 Then strTemp2 = strTemp2 & "<img src=""" & App.Path & "/data/skins/" & ThisSkin.TreeSubNodeImage2 & """ width=""16px"" height=""16px"" align=""left""> " & AllBuddies(i).Servers(i2) & "<br><img src=""" & App.Path & "/data/skins/" & ThisSkin.TreeSeperator & """ height=""2px"" width=""100%""><br>"
                End If
            Next i2
            strTemp2 = strTemp2 & "<img src=""" & App.Path & "/data/skins/" & ThisSkin.TreeSeperator & """ height=""2px"" width=""100%""><br>"
            strResult = strResult & "<tr><td><a href=""nodescript:/buddychangeview " & AllBuddies(i).Name & """><img name='tree' src='../../skins/" & ThisSkin.TreeExpanedImage & "' border=""0"" height=""15px"" width=""15px"" align=""left""></a> <acronym title=""" & Language(498) & """>" & "<a class=""buddyon"" href=""NodeScript:/buddyprofile " & AllBuddies(i).Name & """>" & AllBuddies(i).Name & "</a></acronym><font class=""buddyon"">" & " " & Language(397) & " at:<br><div id=""" & AllBuddies(i).Name & """ style='display:inline'> " & strTemp2 & "</div></font></td></tr>&nbsp;&nbsp;&nbsp;" & vbNewLine
        Else
            strResult = strResult & "<tr><td>" & "<acronym title=""" & Language(498) & """>" & "<a class=""buddyoff"" href=""NodeScript:/buddyprofile " & AllBuddies(i).Name & """>" & AllBuddies(i).Name & "</a></acronym><font class=""buddyoff"">" & " " & Language(566) & "</font></td></tr>&nbsp;&nbsp;&nbsp;<br>" & vbNewLine
        End If
    Next i
    
    GetBuddies = strResult
Not_Connected:
End Function
Public Sub HotKeysAdd(ByVal strHotKey As String, ByVal strAction As String)
    Dim intFL As Integer, intFL2 As Integer
    Dim strHotKeyLine As String
    Dim boolReplaced As Boolean
    
    'if the hotkeys configuration file does not exist...
    If Not FS.FileExists(App.Path & "/conf/hotkeys.dat") Then
        '...create it
        intFL = FreeFile
        Open App.Path & "/conf/hotkeys.dat" For Output As #intFL
        Close #intFL
    End If
    intFL = FreeFile
    Open App.Path & "/conf/hotkeys.dat" For Input As #intFL
    intFL2 = FreeFile
    Open App.Path & "/conf/hotkeys.tmp" For Output As #intFL2
    Do Until EOF(intFL)
        Line Input #intFL, strHotKeyLine
        'syntax:
        'hotkey-string action-string
        
        'if hotkey already exists
        If GetStatement(strHotKeyLine) = strHotKey Then
            boolReplaced = True
            'ask if the user wants to replace it with the new one
            If MsgBox(Language(616), vbYesNo Or vbQuestion, Language(617)) = vbYes Then
                'replace
                Print #intFL2, strHotKey & " " & strAction
            Else
                'don't replace; the file stays the same
                Print #intFL2, strHotKeyLine
            End If
        Else
            Print #intFL2, strHotKeyLine
        End If
    Loop
    If Not boolReplaced Then
        Print #intFL2, strHotKey & " " & strAction
    End If
    Close #intFL
    Close #intFL2
    FS.DeleteFile App.Path & "/conf/hotkeys.dat"
    FS.MoveFile App.Path & "/conf/hotkeys.tmp", App.Path & "/conf/hotkeys.dat"
End Sub
Public Sub HotKeysRemove(ByVal strHotKey As String)
    Dim intFL As Integer, intFL2 As Integer
    Dim strHotKeyLine As String
    Dim boolReplaced As Boolean
    
    'if the hotkeys configuration file does not exist...
    If Not FS.FileExists(App.Path & "/conf/hotkeys.dat") Then
        'don't delete anything
        Exit Sub
    End If
    
    intFL = FreeFile
    Open App.Path & "/conf/hotkeys.dat" For Input As #intFL
    intFL2 = FreeFile
    Open App.Path & "/conf/hotkeys.tmp" For Output As #intFL2
    Do Until EOF(intFL)
        Line Input #intFL, strHotKeyLine
        'only copy the other hot keys
        'skin the passed argument
        If GetStatement(strHotKeyLine) <> strHotKey Then
            Print #intFL2, strHotKeyLine
        End If
    Loop
    Close #intFL
    Close #intFL2
    FS.DeleteFile App.Path & "/conf/hotkeys.dat"
    FS.MoveFile App.Path & "/conf/hotkeys.tmp", App.Path & "/conf/hotkeys.dat"
End Sub
Public Function KeysMatch(ByVal strHotKey, ByVal KeyCode As Integer, ByVal Shift As Integer) As Boolean
    Dim boolReturn As Boolean
    
    boolReturn = False
    
    If KeyCode >= 112 And KeyCode <= 123 Then
        'hit an f-key
        If Left$(strHotKey, Len("function.")) = "function." Then
            If KeyCode - 111 = Right$(strHotKey, Len(strHotKey) - Len("function.")) Then
                boolReturn = True
            End If
        End If
    ElseIf Shift = 2 Then
        If Left$(strHotKey, Len("ctrl.")) = "ctrl." Then
            If KeyCode = AscW(UCase$(Right$(strHotKey, Len(strHotKey) - Len("ctrl.")))) Then
                boolReturn = True
            End If
        End If
    End If
    
    KeysMatch = boolReturn
End Function
Public Sub ExecuteAction(ByVal strAction As String)
    Dim i As Integer
    Dim boolReseek As Boolean
    
    Select Case strAction
        Case "show.toolbar"
            frmMain.imgMore_Click
        Case "toolbar.bold"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(3)
        Case "toolbar.italic"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(4)
        Case "toolbar.underline"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(5)
        Case "toolbar.color"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(7)
        Case "toolbar.smiley"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(1)
        Case "toolbar.image"
            frmMain.tbText_ButtonClick frmMain.tbText.Buttons.Item(9)
        Case "go.nexttab"
            If CurrentActiveServer.Tabs.SelectedItem.index = CurrentActiveServer.Tabs.Tabs.Count Then
                Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(1)
            Else
                Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(CurrentActiveServer.Tabs.SelectedItem.index + 1)
            End If
        Case "go.nextfocus"
            i = CurrentActiveServer.Tabs.SelectedItem.index
            Do
                i = i + 1
                If i > CurrentActiveServer.Tabs.Tabs.Count Then
                    i = 1
                    If Not boolReseek Then
                        boolReseek = True
                    Else
                        Exit Do
                    End If
                End If
                If CurrentActiveServer.Tabs.Tabs.Item(i).Image = TabImage_Look Then
                    Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(i)
                    Exit Do
                End If
            Loop
        Case "go.help"
            frmMain.nmnuHelp_MenuClick 0
        Case "panel.join"
            frmMain.LoadPanel "join"
        Case "panel.connect"
            frmMain.LoadPanel "connect"
    End Select
End Sub
Public Sub Quit(ByRef TheServer As clsActiveServer)
    Dim strQuit As String
    Dim intFL As Integer
    Dim intLnCount As Integer
    Dim a As Integer
    
    If Options.QuitMultiple Then
        If FS.FileExists(Options.QuitFile) Then
            intFL = FreeFile
            Open Options.QuitFile For Input As #intFL
            Do Until EOF(intFL)
                xLineInput intFL
                intLnCount = intLnCount + 1
            Loop
            Close #intFL
            
            a = Rnd() * intLnCount
            
            intLnCount = 0
            intFL = FreeFile
            Open Options.QuitFile For Input As #intFL
            Do Until EOF(intFL)
                If intLnCount = a Then
                    Line Input #intFL, strQuit
                    Exit Do
                Else
                    xLineInput intFL
                End If
                intLnCount = intLnCount + 1
            Loop
            Close #intFL
        Else
            strQuit = Options.QuitMsg
            DB.XWarning "Quit List File does not exist!"
        End If
    Else
        strQuit = Options.QuitMsg
    End If
    On Error Resume Next
    TheServer.SendData "QUIT :" & strQuit & vbNewLine
End Sub
Public Sub LoadIRCMsg()
    Dim intFL As Integer
    Dim strIDs() As String
    Dim strTemp As String
    Dim intLineIndex As Integer
    Dim i As Integer
    
    intFL = FreeFile
    Open App.Path & "/data/ircmsg.dat" For Input Access Read Shared As #intFL
    Do Until EOF(intFL)
        Line Input #intFL, strTemp
        intLineIndex = intLineIndex + 1
        If Left$(strTemp, 1) <> "#" And LenB(strTemp) > 0 Then
            i = i + 1
            'it's not a comment
            strIDs = Split(strTemp, " ", 2)
            If Not IsNumeric(strIDs(0)) Or Not IsNumeric(strIDs(1)) Then
                MsgBox Replace(Language(382), "%1", CStr(intLineIndex)), vbExclamation
            Else
                IRCMsg(i, 0) = Val(strIDs(0))
                IRCMsg(i, 1) = Val(strIDs(1))
            End If
        End If
    Loop
    Close #intFL
End Sub
Public Sub CriticalError()
    ReportBug Language(621) & vbNewLine
    End
End Sub
Public Sub ReportBug(Optional ByVal strPretext As String = vbNullString)
    If MsgBox(strPretext & Language(533), vbYesNo Or vbQuestion, Language(534)) = vbYes Then
        xShell "http://sourceforge.net/tracker/?func=add&group_id=94591&atid=608388 """"", 0
    End If
End Sub
Public Function MIRCColorToHTML(ByVal mIRCColor As Integer) As String
    'Convert an mIRC color index to the HTML equivilant
    Select Case mIRCColor
        Case 0
            MIRCColorToHTML = "#FFFFFF"
        Case 1
            MIRCColorToHTML = "#000000"
        Case 2
            MIRCColorToHTML = "#000099"
        Case 3
            MIRCColorToHTML = "#009900"
        Case 4
            MIRCColorToHTML = "#CC0000"
        Case 5
            MIRCColorToHTML = "#660000"
        Case 6
            MIRCColorToHTML = "#9900FF"
        Case 7
            MIRCColorToHTML = "#FF6600"
                
        Case 8
            MIRCColorToHTML = "#FFFF00"
        Case 9
            MIRCColorToHTML = "#00FF33"
        Case 10
            MIRCColorToHTML = "#0099CC"
        Case 11
            MIRCColorToHTML = "#00FFFF"
        Case 12
            MIRCColorToHTML = "#0033FF"
        Case 13
            MIRCColorToHTML = "#CC0099"
        Case 14
            MIRCColorToHTML = "#666666"
        Case 15
            MIRCColorToHTML = "#CCCCCC"
        Case 16
            MIRCColorToHTML = "#FFFF00"
    End Select
End Function
Public Function GetPriviledges(ByVal strFullNick As String) As String
    GetPriviledges = Strings.Left$(strFullNick, Len(strFullNick) - Len(GetNick(strFullNick)))
End Function
Public Function GetNickID(ByVal Nickname As String, ByVal ChannelId As Long, ByRef TheServer As clsActiveServer) As Long
    Dim i As Integer
    
    For i = 1 To TheServer.NickList_List(ChannelId).Count
        'MsgBox Strings.LCase$(GetNick(TheServer.NickList_List(ChannelId).Item(i)))
        If Strings.LCase$(GetNick(TheServer.NickList_List(ChannelId).Item(i))) = Strings.LCase$(Nickname) Then
            GetNickID = i
            Exit Function
        End If
    Next i
    GetNickID = -1 'No such nick
End Function
Public Function GetTabFromBrowser(ByVal BrowserIndex As Integer) As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim ThisTag As String
    
    For i2 = 0 To UBound(ActiveServers)
        If Not ActiveServers(i2) Is Nothing Then
            For i = 1 To ActiveServers(i2).TabInfo.Count
                ThisTag = ActiveServers(i2).TabInfo(i)
                If ActiveServers(i2).TabType(i) = TabType_WebSite Then
                    If ThisTag = BrowserIndex Then
                        GetTabFromBrowser = i
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next i2
    GetTabFromBrowser = -1
End Function
Public Function GetServerFromBrowser(ByVal BrowserIndex As Integer) As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim ThisTag As String
    
    For i2 = 0 To UBound(ActiveServers)
        If Not ActiveServers(i2) Is Nothing Then
            For i = 1 To ActiveServers(i2).TabInfo.Count
                ThisTag = ActiveServers(i2).TabInfo(i)
                If ActiveServers(i2).TabType(i) = TabType_WebSite Then
                    If ThisTag = BrowserIndex Then
                        GetServerFromBrowser = i2
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next i2
    GetServerFromBrowser = -1
End Function
'Public Function GetServerFromWsIrcIndex(ByVal WinSockIndex As Integer) As Integer
'    Dim i As Integer
'
'    For i = 1 To UBound(ActiveServers)
'        If frmMain.wsIRC(WinSockIndex) Is ActiveServers(i).WinSockConnection Then
'            GetServerFromWsIrcIndex = i
'            Exit Function
'        End If
'    Next i
'
'    GetServerFromWsIrcIndex = -1
'End Function
Public Function GetServerFromWsDCCIndex(ByVal intWsDccIndex As Integer, ByVal RCV As Boolean) As Integer
    Dim i As Integer
    
    For i = 0 To UBound(ActiveServers)
        If Not ActiveServers(i) Is Nothing Then
            If ActiveServers(i).GetDCCFileIndexFromWsIndex(intWsDccIndex, RCV) > -1 Then
                GetServerFromWsDCCIndex = i
                Exit Function
            End If
        End If
    Next i
    
    GetServerFromWsDCCIndex = -1
End Function
Public Sub XMLPerformSaveServer(ByVal ServerHostName As String, ByVal Perform As String)
    Dim XMLDoc As DOMDocument
    Dim XMLNode As IXMLDOMElement
    Dim XMLAttribute As IXMLDOMAttribute
    Dim XMLVersion As IXMLDOMProcessingInstruction
    Dim XMLComment As IXMLDOMComment
    Dim strNodeValue As String
    Dim i As Integer
    Dim i2 As Integer
    Dim PerformFile As String
    
    PerformFile = App.Path & "/conf/perform.xml"
    
    Set XMLDoc = New DOMDocument
    
    If Not XMLDoc.Load(PerformFile) Then
        'failed to load perform XML file
        'create it
                
        'Init the DOM Document
        XMLDoc.async = False
        XMLDoc.validateOnParse = False
        XMLDoc.resolveExternals = False
        XMLDoc.preserveWhiteSpace = True
        
        'Create the ProcessingInstruction, this is <?xml version='1.0'>
        'That appears at the first line of the XML file
        Set XMLVersion = XMLDoc.createProcessingInstruction("xml", "version='1.0'")
        
        'and add it to the Document
        XMLDoc.appendChild XMLVersion
        
        Set XMLVersion = Nothing
        Set XMLComment = XMLDoc.createComment("This file was automatically created by Node")
        
        XMLDoc.appendChild XMLComment
        
        Set XMLComment = Nothing
        
        'build the Document Element
        Set XMLNode = XMLDoc.createElement("Perform")
        
        'save the version of Node using an attribute of the Document Element
        XMLNode.setAttribute "Vesion", App.Major & "." & App.Minor
        
        'set the root element
        Set XMLDoc.documentElement = XMLNode
            
        'save the result
        XMLDoc.save PerformFile
   
        'unload the XML objects
        Set XMLNode = Nothing
    End If
    
    For i = 0 To XMLDoc.documentElement.childNodes.length - 1
        Set XMLNode = XMLDoc.documentElement.childNodes.Item(i)
        If XMLNode.nodeName = "Server" Then
            For i2 = 0 To XMLNode.Attributes.length - 1
                Set XMLAttribute = XMLNode.Attributes.Item(i2)
                If XMLAttribute.nodeName = "HostName" Then
                    strNodeValue = XMLAttribute.nodeValue
                    If LCase$(strNodeValue) = LCase$(ServerHostName) Then
                        'this is the node we should be saving in
                        XMLNode.Text = Perform
                        GoTo SaveFile
                    End If
                End If
            Next i2
        End If
    Next i
    
    'the server node we're searching for seems not to exist
    'create it
    Set XMLNode = XMLDoc.createElement("Server")
    XMLNode.setAttribute "HostName", ServerHostName
    'add perform to the node
    XMLNode.Text = Perform
    
SaveFile:
    'append it to the XML document
    XMLDoc.documentElement.appendChild XMLNode
    
    XMLDoc.save PerformFile
End Sub
Public Function XMLPerformReadServer(ByVal ServerHostName As String) As String
    Dim XMLDoc As DOMDocument
    Dim XMLNode As IXMLDOMElement
    Dim XMLAttribute As IXMLDOMAttribute
    Dim strNodeValue As String
    Dim i As Integer
    Dim i2 As Integer
    Dim PerformFile As String
    
    PerformFile = App.Path & "/conf/perform.xml"
    
    Set XMLDoc = New DOMDocument
    
    If Not XMLDoc.Load(PerformFile) Then
        'failed to load perform XML file
        Exit Function
    End If
    
    For i = 0 To XMLDoc.documentElement.childNodes.length - 1
        Set XMLNode = XMLDoc.documentElement.childNodes.Item(i)
        If XMLNode.nodeName = "Server" Then
            For i2 = 0 To XMLNode.Attributes.length - 1
                Set XMLAttribute = XMLNode.Attributes.Item(i2)
                If XMLAttribute.nodeName = "HostName" Then
                    strNodeValue = XMLAttribute.nodeValue
                    If LCase$(strNodeValue) = LCase$(ServerHostName) Then
                        'this is the node we should be reading from
                        XMLPerformReadServer = XMLNode.Text
                        Exit Function
                    End If
                End If
            Next i2
        End If
    Next i
    
    'oops, Node not found
    'we won't return anything, then
End Function
Public Sub StartNetMeetingSession(ByVal LocalNetMeetingServer As Boolean, Optional ByVal RemoteIP As String)
    'Start a NetMeeting Session
    'both sides have to call this
    'sub in order to connect
    'The one with LocalNetMeetingServer set to True
    'and the second with LocalNetMeetingServer set to False
    'but with RemoteIP argument passed
    
    If LocalNetMeetingServer Eqv Not LenB(RemoteIP) = 0 Then
        'you cannot set LocalNetMeetingServer to true
        'and pass a remote IP address
        'or set LocaLNetMeetingServer to false
        'and not pass a remote IP address
        'wrong syntax
        Exit Sub
    End If
    
    'load NetMeeting
    Set NM = CreateObject("NetMeeting.App.1")
    If Not LocalNetMeetingServer Then
        NM.CallTo RemoteIP
    End If
End Sub
Public Function GetNDCFromNickname(ByVal Nickname As String) As Integer
    Dim i As Integer
    For i = 0 To UBound(NDCConnections)
        If LCase$(NDCConnections(i).strNicknameA) = LCase$(Nickname) Then
            GetNDCFromNickname = i
            Exit Function
        End If
    Next i
    GetNDCFromNickname = -1
End Function
Public Function IrcGetLongIP(ByVal AscIp As String) As String
    'this function converts an ascii ip string into a long ip in network byte order
    'and stick it in a string suitable for use in a DCC command.
    
    On Error GoTo IrcGetLongIpError:
    Dim inn As Long
    
    inn = htonl(inet_addr(AscIp))
    If inn < 0 Then
        IrcGetLongIP = CVar(inn + 4294967296#)
    Else
        IrcGetLongIP = CVar(inn)
    End If
    
    Exit Function
IrcGetLongIpError:
    IrcGetLongIP = "0"
End Function
Public Function IrcGetNormalIP(ByVal LongIp As String) As String
    'this function converts a long ip in network byte order
    'into a string ip in the form of "127.0.0.1"
    
    On Error GoTo IrcGetNormalIPError
    
    'TO DO!!
    MsgBox "Warning: Call to IrcGetNormalIP(). This function is still under construction!", vbCritical
    
    Exit Function
IrcGetNormalIPError:
    IrcGetNormalIP = "0.0.0.0"
End Function
Public Sub InitWindow(ByVal hwnd As Long)
    'This sub makes the window able to be transparent
    'it is only called once at the start of the program
    'for better performarce
    SetWindowLong hwnd, GWL_EXSTYLE, GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
End Sub
Public Sub SetLayered(ByVal hwnd As Long, ByVal bAlpha As Byte)
    'this sub changes the transparency value
    SetLayeredWindowAttributes hwnd, 0, bAlpha, LWA_ALPHA
End Sub
Public Sub PartCopyR(ByVal intDestinationFile As Integer, ByVal intSourceFile As Integer)
    Do Until EOF(intSourceFile)
        Print #intDestinationFile, DoReplace(xLineInput(intSourceFile), False)
    Loop
End Sub
Public Sub Reload()
    Dim Window As Form
    Dim i As Integer
    Dim ActiveServer As clsActiveServer
    
    DB.Enter "Reload"
    
    Restarting = True
    
    DB.X "Kill Session XML"
    
    If FS.FileExists(App.Path & "\temp\session.xml") Then
        Kill App.Path & "\temp\session.xml"
    End If

    Unload frmMain
    Unload frmOptions
    
    On Error GoTo No_Servers_Loaded
    For i = 0 To UBound(ActiveServers)
        If Not ActiveServers(i) Is Nothing Then
            Set ActiveServers(i) = Nothing
        End If
    Next i
    
No_Servers_Loaded:
    ReDim ActiveServers(0)
    
    Set frmMain.xpBalloon = Nothing
    
    DB.X "Restarting..."
    Set DB = Nothing
    
    Unload frmDebug
        
    SaveSetting "Node", "Remember", "Reload", True
    Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
    End
    'Restarting = False
    'Main
End Sub
Public Sub LoadMultiLanguage(ByVal FileName As String)
    'This sub is used to load the FileName language-file into the array LanguageMulti
    Dim intFL As Integer 'file index
    Dim strCurrentLine As String 'the current line of the file
    Dim strStatement As String, intstatement As Integer
    Dim i As Integer, i2 As Integer, linenum As Integer   'the key index for each language caption
    Dim objFile As File
    Dim objTextStream As TextStream
    Dim strLanguage As String
    Dim strTemp As String
        
        For i = 0 To UBound(LangMultiNames)
            If LCase$(LangMultiNames(i)) = LCase$(FileName) Then
                Exit Sub
            End If
        Next i
        ReDim Preserve LangMultiNames(UBound(LangMultiNames) + 1)
        'get a free file index
        intFL = FreeFile
        'open the language file
        'Set objFile = fs.GetFile(FileName)
        'Set objTextStream = objFile.OpenAsTextStream(ForReading)
        strLanguage = DecodeFile(App.Path & "\data\languages\" & FileName & ".lang", False)
        'Open FileName For Binary Access Read Shared As #intFl
        'skip the first two lines
        'strCurrentLine = objTextStream.ReadLine & objTextStream.ReadLine
        strCurrentLine = DecodedLineInput(strLanguage) & DecodedLineInput(strLanguage)
        i = 2
        linenum = -1
        'go through the file
        i2 = UBound(LangMultiNames)
        LangMultiNames(i2) = FileName
        Do Until Len(strLanguage) = 0
            i = i + 1
        
            'get a line from the file(either a key or a comment)
            strCurrentLine = DecodedLineInput(strLanguage)
            'strCurrentLine = objTextStream.ReadLine
            'if it has a leading escape, remove it
            If Strings.Left$(strCurrentLine, 1) = ">" Then
                strCurrentLine = Strings.Right$(strCurrentLine, Len(strCurrentLine) - 1)
            End If
            'if it's not a comment add the current key to the array
            '                                        & ChrW$(0)
            If Strings.Left$(strCurrentLine, 1) <> "#" And Len(strCurrentLine) <> 0 Then
                'this line is not a comment
                'error trap for language file bugs
                On Error GoTo Language_File_Error
                'key indexes can't be unicode
                strStatement = GetStatement(strCurrentLine) ', vbFromUnicode)
                If IsNumeric(strStatement) Then
                    intstatement = Val(strStatement)
                    'TO DO
                    'Language(strStatement) = GetParameter(StrConv(strCurrentLine, vbFromUnicode))
                    'Language(strStatement) = GetParameter(strCurrentLine, , , True)
                    LanguageMulti(i2).Cell(1, intstatement) = Replace(GetParameter(strCurrentLine, , , False), ChrW$(0), vbNullString)
                End If
             End If
         Loop
EOF_Return:
    'objTextStream.Close
    Close #intFL
    Exit Sub
Language_File_Error:
    If Err.Number = 62 Then
        'EOF
        Resume EOF_Return
    End If
    MsgBox "The language file " & FileName & ".lang has an error in line " & i, vbCritical, "Language File Error"
End Sub
Public Sub CheckGoBack()
    frmMain.tmrAway.Enabled = False
    frmMain.tmrAway.Enabled = True
    frmMain.AwayMins = 0
    
    If IsAway Then
        GoAway False
    End If
End Sub
Public Sub GoAway(ByVal Away As Boolean)
    Dim TheServer As clsActiveServer
    Dim intFL As Integer
    Dim i As Integer
    
    If Away Then
        frmMain.AwayMins = 0
        'going away
        IsAway = True
        For i = 0 To UBound(ActiveServers)
            Set TheServer = ActiveServers(i)
            If Not TheServer Is Nothing Then
                If TheServer.WinSockConnection.State = sckConnected Then
                    If Options.AwayNick Then
                        TheServer.preExecute "/nick " & Options.AwayNickStr
                    End If
                    If Options.AwayPerform Then
                        intFL = FreeFile
                        Open App.Path & "/temp/away_perform.dat" For Output As #intFL
                        Print #intFL, Options.AwayPerformStr
                        Close #intFL
                        TheServer.Perform App.Path & "/temp/away_perform.dat"
                    End If
                    If Options.AwayStatus Then
                        TheServer.MyStatus = Options.AwayStatusID
                    End If
                End If
            End If
        Next i
    Else
        'backing from away
        IsAway = False
        For i = 0 To UBound(ActiveServers)
            Set TheServer = ActiveServers(i)
            If Not TheServer Is Nothing Then
                If TheServer.WinSockConnection.State = sckConnected Then
                    If Options.AwayBackNick Then
                        TheServer.preExecute "/nick " & Options.AwayBackNickStr
                    End If
                    If Options.AwayBackPerform Then
                        intFL = FreeFile
                        Open App.Path & "/temp/away_perform.dat" For Output As #intFL
                        Print #intFL, Options.AwayBackPerformStr
                        Close #intFL
                        TheServer.Perform App.Path & "/temp/away_perform.dat"
                    End If
                    If Options.AwayBackStatus Then
                        TheServer.MyStatus = Options.AwayBackStatusID
                    End If
                End If
            End If
        Next i
    End If
End Sub
Public Function WinSockErrorIDToLangString(ByVal ErrorID As Long) As String
    Dim LangID As Integer
    Dim strReturn As String
    
    Select Case ErrorID
        Case sckBadState
            LangID = 889
        Case sckConnectionRefused
            LangID = 890
        Case sckConnectAborted
            LangID = 891
        Case sckHostNotFound
            LangID = 893
        Case sckAddressInUse
            LangID = 894
        Case Else
            LangID = -1
            'unknown error
            WinSockErrorIDToLangString = Replace(Language(895), "%1", CStr(ErrorID))
            Exit Function
    End Select
    
    strReturn = Language(LangID)
    
    WinSockErrorIDToLangString = strReturn
End Function
Public Function PortIsInUse(ByVal PortNumber As Long) As Boolean
    Dim mibTcpTable As MIB_TCPTABLE
    Dim pdwSize As Long
    Dim Border As Long
    Dim i As Integer
    Dim localportnum As Long
    
    DoEvents
    
    'get netstat
    GetTcpTable mibTcpTable, pdwSize, Border
    'twice (that's the only way it works!)
    GetTcpTable mibTcpTable, pdwSize, Border
    
    For i = 0 To mibTcpTable.dwNumEntries - 1
        localportnum = mibTcpTable.mibtable(i).dwLocalPort
        '
        '(mozillagodzilla) WARNING: forward slashes (/) only run at about 9.3 million operations
        'per second, while using backslashes (\) for division runs at about 21.2 million operations
        'per second. However, forward slashes also will keep the remainder, so it's recommended to
        'use forward slashes for non-integer division, while using back slashes for integer division.
        '
        'convert from API long to VB long
        localportnum = localportnum \ 256 + (localportnum Mod 256) * 256
        If localportnum = PortNumber Then
            PortIsInUse = True
            Exit Function
        End If
    Next i
    
    PortIsInUse = False
End Function
Public Function ChooseOpenPort(ByVal LowerPort As Long, ByVal HigherPort As Long) As Long
    Dim mibTcpTable As MIB_TCPTABLE
    Dim pdwSize As Long
    Dim Border As Long
    Dim i As Integer
    Dim i2 As Integer
    Dim i3 As Integer
    Dim localportnum As Long
    Dim portok As Boolean
    Dim cAllowedPorts As Collection
    ReDim PortArray(0)
    DoEvents
    
    'get netstat
    GetTcpTable mibTcpTable, pdwSize, Border
    'twice (that's the only way it works!)
    GetTcpTable mibTcpTable, pdwSize, Border
    
    For i = 0 To mibTcpTable.dwNumEntries - 1
        On Error GoTo Connections_Changed
        localportnum = mibTcpTable.mibtable(i).dwLocalPort
        localportnum = localportnum \ 256 + (localportnum Mod 256) * 256
        ReDim Preserve PortArray(i)
        PortArray(i) = localportnum
    Next i
    'subscript out of range
Connections_Changed:
    
    Set cAllowedPorts = New Collection
    
    For i = LowerPort To HigherPort
        cAllowedPorts.Add i
    Next i
    
    Do
        'i3 = Int(Rnd * (LowerPort - HigherPort) + LowerPort)
        i3 = Int(Rnd * cAllowedPorts.Count) + 1
        portok = True
        For i2 = 0 To UBound(PortArray)
            If i3 = PortArray(i2) Then
                portok = False
                Exit For
            End If
        Next i2
        If portok Then
            Exit Do
        End If
    Loop
    
    ChooseOpenPort = cAllowedPorts.Item(i3)
End Function
Public Function ParseMemo(ByVal strNotice As String, ByRef TheServer As clsActiveServer) As Boolean
    Dim strMemoText As String
    Dim strNick As String
    Dim intMemoIndex As Integer
    Dim strMemoIndex As String
    Dim intMemoCount As Integer
    Dim strMemoCount As String
    Dim strMemoTime As String
    Dim intPos As Integer
    Dim intPos2 As Integer
    
    ParseMemo = True
        
    strMemoText = LCase$(Trim$(strNotice))
    
    If TheServer.npmpData <> nodeMemoParseNone Then
        Select Case TheServer.npmpData
            Case nodeMemoParseList
                If strMemoText <> "-- end of list --" Then '<-- warning: aspect do not send '-- end of list --'!
                                                            '<-- warning: grnet does not send '-- end of list --'!
                    TheServer.strTempMemoParse = TheServer.strTempMemoParse & vbNewLine & strNotice
                Else
                    MsgBox TheServer.strTempMemoParse
                    TheServer.npmpData = nodeMemoParseNone
                    TheServer.strTempMemoParse = vbNullString
                End If
            Case nodeMemoParseRead
                'we're reading a memo
                MsgBox TheServer.strTempMemoParse & vbNewLine & vbNewLine & strNotice
                TheServer.strTempMemoParse = vbNullString
                TheServer.npmpData = nodeMemoParseNone
            Case nodeMemoParseNoIdentify
                'do nothing, just skin the command
                TheServer.npmpData = nodeMemoParseNone
        End Select
        Exit Function
    End If
    
    '"You have a new memo from xxx (#0) <- dancer
    If Left$(strMemoText, Len("you have a new memo from ")) = "you have a new memo from " Then
        intPos = InStr(Len("you have a new memo from "), strMemoText, MIRC_BOLD)
        If intPos <> 0 Then
            intPos2 = InStr(intPos + 1, strMemoText, MIRC_BOLD)
            If intPos2 <> 0 Then
                strNick = Mid$(strMemoText, intPos + 1, intPos2 - intPos - 1)
                intPos = InStr(intPos2 + 1, strMemoText, "(")
                If intPos <> 0 Then
                    intPos2 = InStr(intPos + 1, strMemoText, ")")
                    If intPos2 <> 0 Then
                        strMemoIndex = Mid$(strMemoText, intPos + 1, intPos2 - intPos - 1)
                        If Left$(strMemoIndex, 1) = "#" Then
                            strMemoIndex = Right$(strMemoIndex, Len(strMemoIndex) - 1)
                        End If
                        If IsNumeric(strMemoIndex) Then
                            intMemoIndex = Val(strMemoIndex)
                        End If
                    End If
                Else
                    'some services do not send the memo index
                    intMemoIndex = -1
                End If
                MsgBox "Hey there! There's a new memo for you! It was just sent by " & strNick & " and its index is " & intMemoIndex & ". Use the IRC menu to read it!", vbInformation
            End If
        End If
        Exit Function
    End If
    
    'You have 0 new memo(s) <- dancer
    '         ^bold
    If Left$(strMemoText, Len("you have ")) = "you have " Then
        If Right$(strMemoText, Len(" new memo")) = " new memo" Then
            intPos = InStr(Len("you have "), strMemoText, MIRC_BOLD)
            If intPos <> 0 Then
                intPos2 = InStr(intPos + 1, strMemoText, MIRC_BOLD)
                If intPos2 <> 0 Then
                    strMemoCount = Mid$(strMemoText, intPos + 1, intPos2 - intPos - 1)
                    If IsNumeric(strMemoCount) Then
                        intMemoCount = Val(strMemoCount)
                        MsgBox "Hey there! There " & IIf(intMemoCount = 1, "is", "are") & " " & intMemoCount & " new memo" & IIf(intMemoCount = 1, vbNullString, "s") & " for you! Use the IRC menu to read it!", vbInformation
                        Exit Function
                     End If
                End If
            End If
        End If
    End If
    
    'You have 0 new memo(s). <-- aspect services
    '        ^no bold      ^dot
    
    If Left$(strMemoText, Len("you have ")) = "you have " Then
        If Right$(strMemoText, Len(" new memo.")) = " new memo." Then
            intPos = InStr(Len("you have "), strMemoText, " ")
            If intPos <> 0 Then
                intPos2 = InStr(intPos + 1, strMemoText, " ")
                If intPos2 <> 0 Then
                    strMemoCount = Mid$(strMemoText, intPos + 1, intPos2 - intPos - 1)
                    If IsNumeric(strMemoCount) Then
                        intMemoCount = Val(strMemoCount)
                        MsgBox "Hey there! There " & IIf(intMemoCount = 1, "is", "are") & " " & intMemoCount & " new memo" & IIf(intMemoCount = 1, vbNullString, "s") & " for you! Use the IRC menu to read it!", vbInformation
                        Exit Function
                     End If
                End If
            End If
        End If
    End If
    
    'Type /msg MemoServ LIST to view them
    'Type /msg MemoServ READ 0 to read it
    'Type /msg MemoServ READ LAST to read it.
    'Type /msg NickServ IDENTIFY <password> and retry
    If Left$(strMemoText, Len("type ")) = "type " Then
        Exit Function
    End If
    
    'To delete, type /msg MemoServ DEL 1 <- aspect
    If Left$(strMemoText, Len("to delete, type ~/msg memoserv del ")) = "to delete, type " & MIRC_BOLD & "/msg memoserv del " Then
        Exit Function
    End If
   
    '-- Listing memos for [dionyziz] -- <- dancer
    If Left$(strMemoText, Len("-- Listing memos for [")) = "-- listing memos for [" Then
        If Right$(strMemoText, Len("] --")) = "] --" Then
            TheServer.npmpData = nodeMemoParseList
            TheServer.strTempMemoParse = "Memo List"
            Exit Function
        End If
    End If
    
    'Memos for dionyziz. To read, type /msg MemoServ READ num <-- aspect services
    '                                  ^bold               ^unrline
    '                                                          ^underline
    '                                                           ^bold
    'Memos for dionyziz. To read, type: /msg MemoServ READ num <-- grnet services
    If Left$(strMemoText, Len("memos for ")) = "memos for " Then
        If Right$(strMemoText, Len(".  To read, type ~/msg MemoServ READ ~num~~")) = ".  to read, type " & MIRC_BOLD & "/msg memoserv read " & MIRC_UNDERLINE & "num" & MIRC_UNDERLINE & MIRC_BOLD Then
            TheServer.npmpData = nodeMemoParseList
            TheServer.strTempMemoParse = "Memo List"
            Exit Function
            
        ElseIf Right$(strMemoText, Len(".  To read, type: ~/MemoServ READ ~num~~")) = ".  to read, type: " & MIRC_BOLD & "/memoserv read " & MIRC_UNDERLINE & "num" & MIRC_UNDERLINE & MIRC_BOLD Then
            TheServer.npmpData = nodeMemoParseList
            TheServer.strTempMemoParse = "Memo List"
            Exit Function
        End If
    End If
    
    'Memos for Pizzicato.  To read, type: /MemoServ READ num
    
    'Idx Sender             Time Sent <- dancer
    If Left$(strMemoText, Len("idx sender ")) = "idx sender " Then
        If Right$(strMemoText, Len(" time sent")) = " time sent" Then
            Exit Function
        End If
    End If
    
    'Memo #1 from ch-world (sent 22 minutes 22 seconds ago): <- dancer
    If Left$(strMemoText, Len("memo ")) = "memo " Then
        intPos = InStr(Len("memo "), strMemoText, "#")
        If intPos <> 0 Then
            intPos2 = InStr(intPos + 1, strMemoText, " ")
            If intPos2 <> 0 Then
                strMemoIndex = Mid$(strMemoText, intPos + 1, intPos2 - intPos - 1)
                If IsNumeric(strMemoIndex) Then
                    intMemoIndex = Val(strMemoIndex)
                    If Mid$(strMemoText, intPos2, Len(" from ")) = " from " Then
                        intPos = InStr(intPos2 + Len(" from "), strMemoText, " ")
                        If intPos <> 0 Then
                            strNick = Mid$(strMemoText, intPos2 + Len(" from "), intPos - intPos2 - Len(" from "))
                            intPos2 = InStr(intPos2 + 1, strMemoText, ")")
                            If intPos2 <> 0 Then
                                strMemoTime = Mid$(strMemoText, intPos + 2, intPos2 - intPos - 2)
                                TheServer.strTempMemoParse = "Reading memo " & intMemoIndex & " sent by " & strNick & " " & strMemoTime
                                TheServer.npmpData = nodeMemoParseRead
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'Please identify with NickServ first, using the command: <-- aspect services
    If strMemoText = "please identify with nickserv first, using the command:" Then
        'we'll need to skip the next incoming MemoServ text; it's the command
        TheServer.npmpData = nodeMemoParseNoIdentify
        'display warning
        '"You have not identified"
        If MsgBox(Language(914) & vbNewLine & Language(915), vbExclamation Or vbYesNo, Language(913)) = vbYes Then
            frmMain.nmnuIRC_MenuClick 2, 0 'identify
        End If
        Exit Function
    End If
    
    'Memo #1 has been marked for deletion <-- dancer
    If Left$(strMemoText, Len("Memo #")) = "memo #" Then
        intPos = InStr(Len("Memo #") + 1, strMemoText, " ")
        If intPos <> 0 Then
            strMemoIndex = Mid$(strMemoText, Len("Memo #") + 1, intPos - Len("Memo #") - 1)
            If Right$(strMemoText, Len(strMemoText) - intPos + 1) = " has been marked for deletion" Then
                MsgBox Replace(Language(916), "%1", strMemoIndex), vbInformation, Language(917)
                Exit Function
            End If
        End If
    End If
    
    'Password identification is required for [LIST] <-- dancer
    If Left$(strMemoText, Len("Password identification is required for [")) = "password identification is required for [" Then
        'display warning
        '"You have not identified"
        If MsgBox(Language(914) & vbNewLine & Language(915), vbExclamation Or vbYesNo, Language(913)) = vbYes Then
            frmMain.nmnuIRC_MenuClick 2, 0 'identify
        End If
        Exit Function
    End If
    
    'You have no recorded memos <- dancer
    If strMemoText = "you have no recorded memos" Then
        'if this message is displayed the user
        'must just have done a /memoserv list
        'display a notice
        MsgBox Language(918), vbInformation, Language(919)
        Exit Function
    End If
    
    'You have no new memos <- dancer
    If strMemoText = "you have no new memos" Then
        'this message is displayed on logon
        'do not show a msg box
        Exit Function
    End If
    
    'Memo has been recorded for [cafeina]
    If Left$(strMemoText, Len("memo has been recorded for [~")) = "memo has been recorded for [" & MIRC_BOLD Then
        If Right$(strMemoText, Len("~]")) = MIRC_BOLD & "]" Then
            strNick = Mid$(strNotice, Len("Memo has been recorded for [~") + 1, Len(strNotice) - Len("Memo has been recorded for [~") - Len("~]"))
            MsgBox Replace(Language(920), "%1", strNick), vbInformation, Language(921)
            Exit Function
        End If
    End If
    
    ParseMemo = False
    
    MsgBox "ParseMemo() developement:" & vbNewLine & "There'a new MemoServ notice, which we could not handle." & vbNewLine & vbNewLine & strNotice, vbInformation
    If Not IsCompiled Then
        If MsgBox("Break execution?", vbQuestion Or vbYesNo, "Stop?") = vbYes Then
            Stop
        End If
    End If
End Function

'Part of older CreateMainText
'                If inUnderline Then
'                    strAddition = "</u>" & vbLf
'                    strReturn = Left$(strReturn, i - 1) & strAddition & Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
'                    inUnderline = Not inUnderline
'                    i = i + Len(strAddition) - 1
'                End If
'                If inStrong Then
'                    strAddition = "</strong>" & vbLf
'                    strReturn = Left$(strReturn, i - 1) & strAddition & Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
'                    inUnderline = Not inUnderline
'                    i = i + Len(strAddition) - 1
'                End If
'                If inColor Then
'                    strAddition = "</font>" & vbLf
'                    strReturn = Left$(strReturn, i - 1) & strAddition & Right$(strReturn, Len(strReturn) - i - Len("|") + 1)
'                    inUnderline = Not inUnderline
'                    i = i + Len(strAddition) - 1
'                End If

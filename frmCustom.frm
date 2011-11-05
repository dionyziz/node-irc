VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Custom Dialog"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10155
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRefreshSoon 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   600
   End
   Begin VB.Timer tmrLoadingComplete 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   120
   End
   Begin SHDocVwCtl.WebBrowser wbCustom 
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   2355
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtCustom 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Text            =   "Custom TextBox"
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdCustom 
      Caption         =   "Custom Button"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSComctlLib.ImageCombo iccustom 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Custom ImageCombo"
   End
   Begin VB.Image imgCustom 
      Height          =   2055
      Index           =   0
      Left            =   120
      Top             =   1680
      Width           =   4455
   End
   Begin VB.Label lblCustom 
      Caption         =   "Custom Label"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'Allow only declared variables to be used
Option Explicit
Public Special As Boolean
Public UseAdvancedIdentifiers As Boolean
Public webdocWeb As HTMLDocument
Public ActiveDialog As String
Public DialogData As String
Public DialogData2 As Integer
Public DialogData3 As Boolean
Public DialogData4 As String
Public boolDialogData As Boolean
Public boolLoadingComplete As Boolean
Public NC2Call As Boolean 'is DocumentComplete called by NavigateComplete event?
Private Sub Form_Initialize()
    'Initialize XP Common Controls
    DB.Enter "frmCustom.Form_Initialize"
    DB.X "InitCommonControls"
    InitCommonControls
    DB.Leave "frmCustom.Form_Initialize"
End Sub

'these four event subs call the sub ExecuteEvent
'in order to act properly to the dialog events
'the .Tag property of each object stores the
'code to be executed when the object is clicked or changed.

Private Sub cmdCustom_Click(index As Integer)
    ExecuteEvent cmdCustom(index).Tag
End Sub
Private Sub Form_Click()
    ExecuteEvent Me.Tag
End Sub
Private Sub Form_Resize()
    Dim strResizableDialogs As String
    Dim aResizableDialogs() As String
    Dim i As Integer
    
    strResizableDialogs = "channels.xml/servers.xml/edit_server.xml/hotkey.xml"
    aResizableDialogs = Split(strResizableDialogs, "/")
    For i = 0 To UBound(aResizableDialogs)
        If LCase$(ActiveDialog) = LCase$(aResizableDialogs(i)) Then
            wbCustom(1).Width = Me.ScaleWidth
            wbCustom(1).Height = Me.ScaleHeight
        End If
    Next i
End Sub
Private Sub lblCustom_Click(index As Integer)
    ExecuteEvent lblCustom(index).Tag
End Sub
Private Sub tmrLoadingComplete_Timer()

    If Not ObjectCollectionItemExists(wbCustom, 1) Then
        tmrLoadingComplete.Enabled = False
        Exit Sub
    End If
    
    If Not wbCustom(1).Busy Then
        boolLoadingComplete = True
        tmrLoadingComplete.Enabled = False
        NC2Call = True
        wbCustom_DocumentComplete 1, Nothing, vbNullString
        NC2Call = False
    End If
End Sub
Private Sub tmrRefreshSoon_Timer()
    If LCase$(ActiveDialog) = "servers.xml" Then
        'we need to get the new servers data
        DialogData = CreateServersList(True)
        If LCase$(frmMain.strCurrentPanel) = "connect" Then
            'frmMain.wbPanel.Refresh2
            frmMain.tmrPanelRefreshSoon.Enabled = True
        End If
    End If
    wbCustom_DocumentComplete Val(tmrRefreshSoon.Tag), Nothing, vbNullString
    tmrRefreshSoon.Enabled = False
End Sub
Private Sub txtCustom_Change(index As Integer)
    ExecuteEvent txtCustom(index).Tag
End Sub
Public Sub wbCustom_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)
    Dim webdocButtons As HTMLDocument
    Dim strServerDescription As String
    Dim strServerHostname As String
    Dim lServerPort As Long
    Dim t As Long
    
    DB.Enter "frmCustom.wbCustom_DocumentComplete"
    DB.X "(Index := " & index & ", URL := " & URL & ")"
    
    Set webdocWeb = wbCustom(index).Document
    On Error Resume Next 'invalid DOM Document
    DB.X "XNode-ing DOM Document"
    xNode webdocWeb
    ExecuteEvent wbCustom(index).Tag

    DB.X "XNodeTag-ging Dom Document"
    Select Case Left$(LCase$(ActiveDialog), Len(ActiveDialog) - Len(".xml"))
        Case "full"
            If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Channel Then
                Set frmMain.webdocFullScreenChanMain = frmMain.webdocFullScreen.parentWindow.frames(0).Document
                Set frmMain.webdocFullScreenChanNicks = frmMain.webdocFullScreen.parentWindow.frames(1).Document
            Else
                Set frmMain.webdocFullScreenChanMain = Nothing
                Set frmMain.webdocFullScreenChanNicks = Nothing
            End If
        Case "wizard"
            Set webdocButtons = webdocWeb.parentWindow.frames(1).Document
            xNodeTag webdocButtons, "lang_back", " < " & Language(276) & " ", , "value"
            xNodeTag webdocButtons, "lang_next", " " & Language(771) & " > ", , "value"
            xNodeTag webdocButtons, "lang_back_dis", " < " & Language(276) & " ", , "value"
            xNodeTag webdocButtons, "lang_next_dis", " " & Language(771) & " > ", , "value"
            Wizard_CallBack Me
        Case "smiley"
            xNodeTag webdocWeb, "lang_insert_smiley", Language(514)
            xNodeTag webdocWeb, "smileys_area", DialogData
            xNodeTag webdocWeb, "lang_cancel", Language(121)
        Case "color"
            xNodeTag webdocWeb, "lang_insert_color", Language(567)
            xNodeTag webdocWeb, "lang_foreground", Language(568)
            xNodeTag webdocWeb, "lang_background", Language(569)
            xNodeTag webdocWeb, "lang_sample", Language(570)
            xNodeTag webdocWeb, "lang_ok", Language(120)
            xNodeTag webdocWeb, "lang_cancel", Language(121)
            xNodeTag webdocWeb, "lang_specify_foreground", Language(707)
        Case "buddysignoff"
            webdocWeb.All.Item("skin_buddyoutimage").src = App.Path & "\data\skins\" & ThisSkin.BuddyOutImage
            xNodeTag webdocWeb, "lang_buddyout", DialogData & " " & Language(516)
            xNodeTag webdocWeb, "closedialog", Language(515)
        Case "buddysignon"
            webdocWeb.All.Item("skin_buddyinimage").src = App.Path & "\data\skins\" & ThisSkin.BuddyInImage
            xNodeTag webdocWeb, "lang_buddyin", DialogData & " " & Language(517)
            xNodeTag webdocWeb, "closedialog", Language(515)
        Case "info"
            xNodeTag webdocWeb, "global_version", Language(0) & " " & VERSION_CODENAME & " " & App.Major & "." & App.Minor
            xNodeTag webdocWeb, "lang_internet", Language(537)
            xNodeTag webdocWeb, "lang_team", Language(538)
            xNodeTag webdocWeb, "lang_thanks", Language(712)
            xNodeTag webdocWeb, "lang_nodehomepage", Language(539)
            xNodeTag webdocWeb, "lang_getsupport", Language(540)
            xNodeTag webdocWeb, "lang_reportbug", Language(541)
            xNodeTag webdocWeb, "lang_makedonation", Language(543)
            xNodeTag webdocWeb, "license_2", Language(863)
            xNodeTag webdocWeb, "license_3", Language(859) & "<br>" & Language(860) & "<br>" & Language(861) & "<br>" & Language(862)
            xNodeTag webdocWeb, "license", Replace(Replace(Language(536), _
                                 "%1", "JavaScript:nodeopen(""http://opensource.org/docs/definition.php"");"), _
                                 "%2", "JavaScript:poplicense();")
            xNodeTag webdocWeb, "lang_contributors", Language(858)
        Case "servers"
            'Organize My Servers
            xNodeTag webdocWeb, "lang_my_servers", Language(595)
            xNodeTag webdocWeb, "lang_edit", Language(89), , "value"
            xNodeTag webdocWeb, "lang_delete", Language(197), , "value"
            xNodeTag webdocWeb, "lang_sort", Language(596), , "value"
            xNodeTag webdocWeb, "lang_close", Language(515)
            xNodeTag webdocWeb, "server_list", DialogData
        Case "edit_server"
            xNodeTag webdocWeb, "lang_ok", Language(120)
            xNodeTag webdocWeb, "lang_cancel", Language(121)
            xNodeTag webdocWeb, "server_id", DialogData

            xNodeTag webdocWeb, "lang_displayname", Language(471)
            xNodeTag webdocWeb, "lang_hostname", Language(472)
            xNodeTag webdocWeb, "lang_port", Language(473)
            xNodeTag webdocWeb, "lang_hostname_spaces", Language(602)
            xNodeTag webdocWeb, "lang_invalid_port", Language(114)
            xNodeTag webdocWeb, "lang_no_hostname", Language(603)
            xNodeTag webdocWeb, "lang_no_port", Language(604)
            xNodeTag webdocWeb, "lang_display_quotes", Language(711)

            ReadServer DialogData, strServerDescription, strServerHostname, lServerPort
            xNodeTag webdocWeb, "serv_name", strServerDescription, vbNullString, "value"
            xNodeTag webdocWeb, "serv_hostname", strServerHostname, vbNullString, "value"
            xNodeTag webdocWeb, "serv_port", lServerPort, vbNullString, "value"
        Case "hotkey"
            xNodeTag webdocWeb, "lang_showtoolbar", Language(618)
            xNodeTag webdocWeb, "lang_textbold", Language(510)
            xNodeTag webdocWeb, "lang_textitalic", Language(511)
            xNodeTag webdocWeb, "lang_textunderlined", Language(512)
            xNodeTag webdocWeb, "lang_insertcolor", Language(563)
            xNodeTag webdocWeb, "lang_insertsmiley", Language(514)
            xNodeTag webdocWeb, "lang_insertimage", Language(562)
            xNodeTag webdocWeb, "lang_nexttab", Language(546)
            xNodeTag webdocWeb, "lang_nexthighlighted", Language(547)
            xNodeTag webdocWeb, "lang_help", Language(136)
            xNodeTag webdocWeb, "lang_join", Language(453)
            xNodeTag webdocWeb, "lang_connect", Language(2)
            
            xNodeTag webdocWeb, "lang_ok", Language(120)
            xNodeTag webdocWeb, "lang_cancel", Language(121)
        Case "channels"
            t = GetTickCount
'            Do While (webdocWeb Is Nothing Or wbCustom(1).Busy) 'And t + 5000 > GetTickCount
'                Set webdocWeb = wbCustom(1).Document
'                Wait 0.1
'            Loop
            If wbCustom(1).Busy Then
                tmrLoadingComplete.Enabled = True
                DB.Leave "frmCustom.wbCustom_DocumentComplete", "We need to wait for WB to complete loading first!!"
                Exit Sub
            End If
            
            xNodeTag webdocWeb, "channel_list", DialogData
            xNodeTag webdocWeb, "lang_chnlst", Language(147)
            xNodeTag webdocWeb, "lang_chan", Language(703)
            xNodeTag webdocWeb, "lang_ppl", Language(704)
            xNodeTag webdocWeb, "lang_topic", Language(705)
            xNodeTag webdocWeb, "active_server_id", DialogData2
        Case "hyperlink"
            xNodeTag webdocWeb, "lang_text", Language(733)
            xNodeTag webdocWeb, "lang_linkto", Language(734)
            xNodeTag webdocWeb, "lang_ok", Language(120)
            xNodeTag webdocWeb, "lang_cancel", Language(121)
            xNodeTag webdocWeb, "data_text", frmMain.txtSend.SelText, , "value"
        Case "kick"
            xNodeTag webdocWeb, "lang_kick", Language(336), , "value"
            xNodeTag webdocWeb, "lang_cancel", Language(121), , "value"
            xNodeTag webdocWeb, "lang_kicktype", Language(807)
            xNodeTag webdocWeb, "lang_kickreason", Language(808)
            xNodeTag webdocWeb, "lang_kick2", Language(336)
            xNodeTag webdocWeb, "lang_ban", Language(812)
            xNodeTag webdocWeb, "lang_kban", Language(813)
            xNodeTag webdocWeb, "lang_profanity", Language(809)
            xNodeTag webdocWeb, "lang_flood", Language(810)
            xNodeTag webdocWeb, "lang_other", Language(811)
            xNodeTag webdocWeb, "lang_other2", Language(811)
        Case "chanmodes"
            'TO DO: Key, Limit, Hidden, No Whisper
            NC2Call = True
            If webdocWeb Is Nothing Then
                tmrLoadingComplete.Enabled = False
                tmrLoadingComplete.Enabled = True
                Exit Sub
            End If
            xNodeTag webdocWeb, "lang_t", Language(822), , , NC2Call
            xNodeTag webdocWeb, "lang_n", Language(823), , , NC2Call
            xNodeTag webdocWeb, "lang_i", Language(824), , , NC2Call
            xNodeTag webdocWeb, "lang_s", Language(825), , , NC2Call
            xNodeTag webdocWeb, "lang_p", Language(826), , , NC2Call
            xNodeTag webdocWeb, "lang_m", Language(827), , , NC2Call
            xNodeTag webdocWeb, "lang_c", Language(828), , , NC2Call
            xNodeTag webdocWeb, "lang_apply", Language(122), , , NC2Call
            xNodeTag webdocWeb, "lang_close", Language(515), , , NC2Call
            xNodeTag webdocWeb, "channame", DialogData4
            xNodeTag webdocWeb, "chan_modes", Replace(Language(876), "%1", DialogData4), , , NC2Call
            'xNodeTag webdocWeb, "lang_h", Language(829)
            xNodeTag webdocWeb, "lang_u", Language(830), , , NC2Call
            xNodeTag webdocWeb, "chanmodes_t", InStr(1, DialogData, "t") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_n", InStr(1, DialogData, "n") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_i", InStr(1, DialogData, "i") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_s", InStr(1, DialogData, "s") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_p", InStr(1, DialogData, "p") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_m", InStr(1, DialogData, "m") > 0, , "checked", NC2Call
            xNodeTag webdocWeb, "chanmodes_c", InStr(1, DialogData, "c") > 0, , "checked", NC2Call
    End Select
    DB.Leave "frmCustom.wbCustom_DocumentComplete"
End Sub
Private Sub wbCustom_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
    DB.Enter "frmCustom.wbCustom_NavigateComplete2"
    DB.X "Calling wbCustom_DocumentComplete"
    NC2Call = True
    wbCustom_DocumentComplete index, pDisp, URL
    NC2Call = False
    DB.Leave "frmCustom.wbCustom_NavigateComplete2"
End Sub

'The WebBrowser object is going to navigate
Private Sub wbCustom_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    Dim sURL As String 'variable used to store the unescaped URL
    Dim strNodeScript As String 'the NodeScript code to be executed
    
    DB.Enter "frmCustom.wbCustom_BeforeNavigate2"
    DB.X "(Index := " & index & ", URL := " & URL & ")"
    
    'if it is a NodeScript(i.e. the URL starts with "NodeScript:")...
    If Strings.Left$(Strings.LCase$(URL), Len("NodeScript:")) = "nodescript:" Then
        DB.X "Executing NodeScript"
        'do not navigate to the URL
        Cancel = True
        '...get the unescaped URL
        DB.X "UnEscape-ing from URL"
        sURL = UnEscape(URL)
        'get the NodeScript: execution code
        DB.X "Parsing NodeScript"
        strNodeScript = Strings.Right$(sURL, Len(sURL) - Len("NodeScript:"))
        'If it is a DialogScript(i.e. the URL starts with "NodeScript:!")...
        If Strings.Left$(strNodeScript, 1) = "!" Then
            '...execute the dialog line
            DB.X "NodeScript is a Dialog NodeScript. Executing"
            DialogScript Strings.Right$(strNodeScript, Len(strNodeScript) - 1)
        
        'if it's not a DialogScript
        Else
            
            'execute the NodeScript
            DB.X "NodeScript is a Normal NodeScript. Executing"
            CurrentActiveServer.preExecute strNodeScript, False
        End If
    End If
    If ActiveDialog = "channels.xml" Then
        Select Case Left$(URL, Len("http:/"))
            Case "http:/", "ftp://"
                If URL <> "http:///" Then
                    Cancel = True
                    xShell """" & URL & """ """"", 0
                End If
        End Select
    End If
    
    DB.Leave "frmCustom.wbCustom_BeforeNavigate2"
End Sub

'An object was clicked/changed. Execute its event code
Private Sub ExecuteEvent(ByVal strStatement As String)
    'if there is no event code to execute...
    If Len(strStatement) = 0 Then
        '...don't do anything
        Exit Sub
    End If
    
    'if the event is to unload the window
    If Strings.Left$(Strings.LCase$(strStatement), Len("unload")) = "unload" Then
        'unload it.
        Unload Me
    
    'else execute the IRC command passed as event execution code
    Else
    
        CurrentActiveServer.preExecute strStatement, False
    End If
End Sub

'There was a DialogScript navigation, to a URL starting with NodeScript:!
Public Sub DialogScript(ByVal FullStatement As String)
    Dim strStatement As String 'the statement to be executed
    Dim strParameter As String
    Dim wbControl As Object
    
    DB.Enter "frmCustom.DialogScript"
    
    strStatement = Strings.LCase$(GetStatement(FullStatement))
    
    'depending on the statement...
    Select Case strStatement
        
        'unload the dialog
        Case "closedialog"
            Unload Me
        
        Case "hidedialog"
            Me.Hide
            
        Case "wizard:goback"
            Wizard_PreviousPage Me
            
        Case "wizard:gonext"
            Wizard_NextPage Me
            
        Case "wizard:gofinish"
            Wizard_Finished Me
            
        Case "webrefresh"
            strParameter = GetParameter(FullStatement)
            If IsNumeric(strParameter) And Int(Val(strParameter)) = Val(strParameter) Then
                'wbCustom(strParameter).Refresh2
                tmrRefreshSoon.Tag = strParameter
                tmrRefreshSoon.Enabled = True
            End If
            
        'or execute a dialog line
        Case Else
            DB.X "Invalid DialogScript"
            'CallByName Me, GetStatement(FullStatement), VbLet, GetParameter(FullStatement)
            'executeDialog FullStatement, Me, 0
    End Select
    
    DB.Leave "frmCustom.DialogScript"
End Sub

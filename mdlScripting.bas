Attribute VB_Name = "mdlScripting"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'allow only declared variables to be used
Option Explicit

Private BolNextModal As Boolean 'Determines whether the next custom dialog will be modal or not

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageM Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'RGN
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type NodePlugIn
    strName As String
    boolLoaded As Boolean
    objPlugIn As prjNPIN.clsLittleFinger
    boolLoadOnStartup As Boolean
End Type

Public Type NodeSkin
    FileName As String 'the filename of the .skin file
    
    CodeBehind As String 'the codebehind script, .vbs
    
    ForeGround As Long 'forecolor
    BackGround As Long 'backcolor
    
    PanelTitleColor As Long
    
    HotColor As Long 'mouseover/active color
    
    FontName As String 'Font for standart windows captions
    FontBold As Boolean 'bold?
    FontItalic As Boolean 'italic?
    FontSize As Byte 'size?
    
    Font2Name As String 'Font for buttons
    Font2Bold As Boolean
    Font2Italic As Boolean
    Font2Size As Byte

    Font3Name As String 'Font for textboxes
    Font3Bold As Boolean
    Font3Italic As Boolean
    Font3Size As Byte
    
    TemplateFile As String 'Style Sheet for channels, privates, and more, .css
    TemplateFileNicks As String
    TemplateFilePrivate As String
    TemplateFileChannel As String
    TemplateFileMisc As String
    
    BackgroundColor As String
    SkinPic As String
    
    Icon_Help As String
    Icon_Options As String
    Icon_Join As String
    Icon_Add As String
    Icon_Web As String
    Icon_Channel As String
    Icon_Server As String
    Icon_Close As String
    Icon_Delete As String
    
    PanelTitleBackgroundStart As String
    PanelTitleBackgroundMid As String
    PanelTitleBackgroundEnd As String
    
    BuddyInImage As String
    BuddyOutImage As String
    
    'BackgroundFile As String 'The background for frmMain, .html
    SplashDialog As String 'The Splash dialog for the skin, .xml
    OptionsBackground As String 'The pic for options categories background, .jpg
    
    CheckBox3D As Boolean 'are checkboxes/radios 3d?
    TextBox3D As Boolean 'are textboxes 3d?
    Frame3D As Boolean 'are frames 3d?
    
    OptionCategoryUser As String 'jpg
    OptionCategoryServers As String 'jpg
    OptionCategoryScripts As String 'jpg
    OptionCategoryLogs As String 'jpg
    OptionCategoryPerform As String 'jpg
    OptionCategoryIgnore As String 'jpg
    OptionCategoryLanguage As String 'jpg
    OptionCategoryBuddys As String 'jpg
    OptionCategoryMisc As String 'jpg
    OptionCategorySkin As String 'jpg
    OptionCategorySmileys As String 'jpg
    OptionCategorySessions As String 'jpg
    
    ButtonUp As String 'jpg
    ButtonDown As String 'jpg
    
    private_ButtonUp As IPictureDisp
    private_ButtonDown As IPictureDisp
    
    MenuBackgroundNormal As Long
    MenuBackgroundHOver As Long
    MenuBackgroundExpanded As Long
    MenuForegroundNormal As Long
    MenuForegroundNormalVertical As Long
    MenuForegroundHOver As Long
    MenuSeperatorColor As Long
    MenuShadow As String
    
    Resize_BottomOffset As Integer 'The bottom margin, in twips, of the main window
    Resize_TopOffset As Integer
    Resize_RightOffset As Integer 'The right margin, in twips, of the main window
    Resize_LeftOffset As Integer 'The left margin, in twips, of the main window
    
    Resize_BottomHiddenOffset As Integer 'How much to increase or descrease the bottom margin, in twips, when the send textbox in the main window is hidden
    
    Resize_TabFocusOffset As Integer 'The right margin of a selected tab.
    Resize_TabNonFocusOffset As Integer 'The right margin of a non-selected tab.
    
    Resize_TextLeftOffset As Integer 'The left margin of the send textbox in the main window
    Resize_TextRightOffset As Integer 'The right margin of the send textbox in the main window
    
    Resize_TextBottomOffset As Integer 'The bottom margin offset of the send textbox in the main window

    TreeExpanedImage As String 'gif
    TreeCollapsedImage As String 'gif
    TreeSubNodeImage As String 'gif
    TreeSeperator As String 'gif
    TreeSubNodeImage2 As String 'gif
    
    MorePic As String
End Type

Public Type nodeSoundScheme
    SndStart As String
    SndEnd As String
    SndConnect As String
    SndDisconnect As String
    SndJoin As String
    SndPart As String
    SndKick As String
    SndQuit As String
    SndError As String
    SndTopicChange As String
    SndModeChange As String
    SndOp As String
    SndDeop As String
    SndVoice As String
    SndDevoice As String
    SndHalfop As String
    SndDehalfop As String
    SndWhisper As String
    SndInvitation As String
    FileName As String
End Type

Public Type nodeSmiley
    FileName As String
    ShortcutText As String
End Type

Public Type nodeSmileyPack
    AllSmileys() As nodeSmiley
    SpecialSmileys() As nodeSmiley
End Type

Public Enum WindowsMediaPlayerConstants
    mpStopped = 0 'Playback is stopped.
    mpPaused = 1 'Playback is paused.
    mpPlaying = 2 'Stream is playing.
    mpWaiting = 3 'Waiting for stream to begin.
    mpScanForward = 4 'Stream is scanning forward.
    mpScanReverse = 5 'Stream is scanning in reverse.
    mpSkipForward = 6 'Skipping to next.
    mpSkipReverse = 7 'Skipping to previous
    mpClosed = 8 'Stream is not open.
End Enum

Public xNodeTempValue As String

Public ThisSoundScheme As nodeSoundScheme
'Public LoadedPlugin() As Object
Public AtLeastOnePluginLoaded As Boolean
Public Plugins() As NodePlugIn
Public NumToPlugIn As New Collection

Public scCodeBehind As ScriptControl 'for skins
Private intScriptingFileIndex As Integer

Private WW_ImportSettingsFile As String
Private WW_Nickname As String
Private WW_Alt As String
Private WW_Alt2 As String
Private WW_RealName As String
Private WW_Email As String
Public NPINHost As Node.clsNPINInterface

'#<section id="html">
Public Function DoReplace(ByVal strText As String, Optional ByVal AdvancedIdentifiers As Boolean = False) As String
    Dim intEndOfDiesh As Integer
    Dim intVarIndex As Integer
    Dim i As Integer
    'replace identifiers with their values(does not apply in scripting, only in .xml dialogs)
    'e.g. $app with (App.Path)
    Dim strReturn As String
    strReturn = strText
    '$app
    strReturn = Replace(strReturn, "$app", App.Path)
    '$space
    strReturn = Replace(strReturn, "$space", " ")
    
    'If AdvancedIdentifiers Then
    '    '$currenttopic
    '    If TabType(frmMain.tsTabs.SelectedItem.Index) = TabType_Channel Then
    '        strReturn = Replace(strReturn, "$currenttopic", NickList(frmMain.tsTabs.SelectedItem.Index).Topic)
    '    End If
    'End If
    
    '$css
    strReturn = Replace(strReturn, "<$css>", "<link href=""" & App.Path & "/Data/Skins/" & ThisSkin.TemplateFile & """ rel=""stylesheet"" type=""text/css"">")
    strReturn = Replace(strReturn, "<$css-chan>", "<link href=""" & App.Path & "/Data/Skins/" & ThisSkin.TemplateFileChannel & """ rel=""stylesheet"" type=""text/css"">")
    strReturn = Replace(strReturn, "<$css-nicks>", "<link href=""" & App.Path & "/Data/Skins/" & ThisSkin.TemplateFileNicks & """ rel=""stylesheet"" type=""text/css"">")
    strReturn = Replace(strReturn, "<$css-priv>", "<link href=""" & App.Path & "/Data/Skins/" & ThisSkin.TemplateFilePrivate & """ rel=""stylesheet"" type=""text/css"">")
    strReturn = Replace(strReturn, "<$css-misc>", "<link href=""" & App.Path & "/Data/Skins/" & ThisSkin.TemplateFileMisc & """ rel=""stylesheet"" type=""text/css"">")
    
    strReturn = Replace(strReturn, "<$lang-err-description>", Language(218))
    strReturn = Replace(strReturn, "<$lang-be-online>", Language(219))
    strReturn = Replace(strReturn, "<$lang-error>", Language(42))
    strReturn = Replace(strReturn, "<$lang-loading>", Language(220))
    strReturn = Replace(strReturn, "<$lang-back>", Language(276))
   
    '$white
    strReturn = Replace(strReturn, "$white", "16777215")
    DoReplace = strReturn
End Function
Public Sub xNode(ByRef HTMLDomDocument As HTMLDocument)
    With HTMLDomDocument.All
        xNodeTag HTMLDomDocument, "version", App.Major & "." & App.Minor
        xNodeTag HTMLDomDocument, "major", App.Major
        xNodeTag HTMLDomDocument, "minor", App.Minor
        xNodeTag HTMLDomDocument, "home", GetSetting(App.EXEName, "Options", "NodeHome", "http://node.sourceforge.net")
        xNodeTag HTMLDomDocument, "lang_yes", Language(280)
        xNodeTag HTMLDomDocument, "lang_no", Language(281)
        xNodeTag HTMLDomDocument, "temp", xNodeTempValue
    End With
End Sub
Public Sub xNodeTag(ByRef HTMLDomDocument As HTMLDocument, ByVal TagName As String, ByVal TagValue As String, Optional ByVal PreFix = "xnode_", Optional ByVal AttributeName As String = "innerHTML", Optional ByVal PerformOneByOne As Boolean = False)
    Dim i As Integer
    
    If Not HTMLDomDocument.All.Item(PreFix & TagName) Is Nothing Then
        CallByName HTMLDomDocument.All.Item(PreFix & TagName), AttributeName, VbLet, TagValue
    ElseIf PerformOneByOne Then
        For i = 0 To HTMLDomDocument.All.length
            On Error GoTo Name_Not_Supported
            If HTMLDomDocument.All.Item(i).Name = PreFix & TagName Then
                CallByName HTMLDomDocument.All.Item(i), AttributeName, VbLet, TagValue
            End If
Name_Not_Supported:
            Resume Next_I
Next_I:
        Next i
    End If
End Sub
Public Sub xNodeTagShow(ByRef HTMLDomDocument As HTMLDocument, ByVal TagName As String, Optional ByVal Visible As Boolean = True, Optional ByVal PreFix = "xnode_")
    If Not HTMLDomDocument.All.Item(PreFix & TagName) Is Nothing Then
        HTMLDomDocument.All.Item(PreFix & TagName).Style.display = IIf(Visible, vbNullString, "none")
    End If
End Sub
Public Function img(ByVal src As String, Optional ByVal alt As String = vbNullString, Optional UseNodeHTML As Boolean = False, Optional Border As Boolean = True)
    Dim strReturn As String
    
    If UseNodeHTML Then
        strReturn = HTML_OPEN
    Else
        strReturn = "<"
    End If
    
    strReturn = strReturn & "img src=""" & src & """"
    If LenB(alt) > 0 Then
        strReturn = strReturn & " alt=""" & alt & """"
    End If
    
    If Not Border Then
        strReturn = strReturn & " border=""0"""
    End If
    
    If UseNodeHTML Then
        strReturn = strReturn & HTML_CLOSE
    Else
        strReturn = strReturn & ">"
    End If
    
    img = strReturn
End Function
'#</section>

'#<section id="scripts">
Public Sub SetSphereVariables()
    'Set sphere variables again; the may have changed
    frmMain.ndScript.ExecuteStatement "My_Nick = """ & CurrentActiveServer.myNick & """" 'the nickname of the user
    frmMain.ndScript.ExecuteStatement "Nick = """ & ScriptNick & """" 'the nickname of the user who did the last irc-action
    frmMain.ndScript.ExecuteStatement "Nick2 = """ & ScriptNick2 & """" 'the nickname of the user who took part in the last irc-action
    frmMain.ndScript.ExecuteStatement "Chan = """ & ScriptChan & """" 'the channel in which the last irc-action happened
    If CurrentActiveServer.Tabs.Tabs.Count <> 0 Then
        frmMain.ndScript.ExecuteStatement "ThisChan = """ & Replace(CurrentActiveServer.Tabs.SelectedItem.Caption, """", """""") & """" 'the current channel
    End If
    frmMain.ndScript.ExecuteStatement "CurrentNick = """ & frmMain.CurrentNick & """" 'the current selected nickname(the last clicked) in the nicklist
    'frmMain.ndScript.ExecuteStatement "CurrentTopic = """ & NickList(frmMain.tsTabs.SelectedItem.Index).Topic & """" 'the topic of the selected channel
End Sub
Public Sub InitSphereVariables(Optional ByRef scTarget As ScriptControl, Optional ByVal LoadMain As Boolean = True)
    Dim CurrentFile As Integer
    Dim strTempVar As String
    Dim strScript As String
    
    If scTarget Is Nothing Then
        Set scTarget = frmMain.ndScript
    End If
    
    'clear any previous code
    scTarget.Reset
    'Sphere variables that won't change while the application is running
    'The following Sphere Variables are deprecated. Use their object equals instead.
    'scTarget.ExecuteStatement "App_Path = """ & App.Path & """" 'Deprecated; Use App.Path instead.
    'scTarget.ExecuteStatement "Node_Version = """ & App.Major & "." & App.Minor & """" 'Deprecated; Use App.Major & "." & App.Minor instead
    'scTarget.ExecuteStatement "Node_Version_Maj = " & App.Major 'Deprecated; Use App.Major instead
    'scTarget.ExecuteStatement "Node_Version_Min = " & App.Minor 'Deprecated; Use App.Minor instead
    'scTarget.ExecuteStatement "QuestionMark = ""?""" 'Deprecated; use "?" instead
    'Object References! :)
    On Error Resume Next 'Already added
    'scTarget.AddObject "objSound", frmMain.MP, True 'Deprecated; Use Node.mp instead.
    scTarget.AddObject "Node", frmMain, True
    scTarget.AddObject "frmOptions", frmOptions, True
    scTarget.AddObject "App", App, True
    scTarget.AddObject "FS", FS, True
    
    'scTarget.AddObject "CAS", CurrentActiveServer, True 'User Node.CAS instead
    'Deprecated; use Node.wbBack.Document instead
    'scTarget.AddObject "mywebdoc", frmMain.wbBack.Document, True
    
    'get a free file index
    intScriptingFileIndex = FreeFile
    If LoadMain Then
        'open the script file
        Open App.Path & "\scripts\main.vbs" For Input Access Read Shared As #intScriptingFileIndex
        'go throught the file
        While Not EOF(intScriptingFileIndex)
            'read one line and store it to strTempVar
            Line Input #intScriptingFileIndex, strTempVar
            'append the new line to strScript
            strScript = strScript & vbNewLine & strTempVar
        Wend
        Close #intScriptingFileIndex
        'the hole script is now stored in strScript
        ''there may be an error in the script's code
        ''add an error trap.
        'On Error GoTo lbShowError
        'add the code in order to execute it
        scTarget.AddCode strScript
    End If
End Sub
Public Sub RunScript(ByVal strRoutine As String)
    'this sub is used to run a specific routine from a .vbs file
    Dim CurrentFile As Integer 'the script's file index
    Dim strTempVar As String 'temporary variable used to store the current line inside the file
    Dim strScript As String 'string variable used to store the hole script's code
    Dim SingleSub As Procedure 'variable storing the current procedure inside the script file when we loop through the procedures
    
    'only if scripting is enabled run the routine
    If Options.EnableScripting = False Then
        'Note: we don't use (Not Options.EnableScripting) as it may return TRUE even if the
        'value Options.EnableScripting is TRUE, because Options.EnableScripting can be NULL.
        Exit Sub
    End If
    
    'if there was no routine passed
    If LenB(strRoutine) = 0 Then
        'the routine executed is Main()
        strRoutine = "Main()"
    End If
    
    'load sphere variables
    SetSphereVariables
Script_Loaded:
    'execute the procedure
    'go throught the procedures
    For Each SingleSub In frmMain.ndScript.Procedures
        'if the called procedure is the one we want to execute
        If Strings.LCase$(SingleSub) = Strings.LCase$(strRoutine) Then
            'execute it
            On Error GoTo lbShowError
            frmMain.ndScript.Run strRoutine
            'and don't search any further
            Exit For
        End If
    'go to the next procedure
    Next SingleSub
lbTerminateLoading:
    'don't show any errors
    Exit Sub
lbShowError:
    'there was an error
    'show information about it
    'show the error
    'MsgBox Language(171) & " " & Err.Number & " " & Language(172) & ": " & _
            vbnewline & "main.vbs" & vbnewline & vbnewline & _
            Language(173) & ": " & vbnewline & _
            strRoutine & vbnewline & vbnewline & _
            Language(174) & ": " & vbnewline & _
            Err.Description _
            , vbCritical + vbMsgBoxHelpButton, Language(175), Err.HelpFile, Err.HelpContext
    Resume lbTerminateLoading
End Sub
'#</section>

'#<section id="dialogs">
Public Sub LoadDialog(ByVal FileName As String, Optional ByRef Window As Form)
    'Sub to load elements from
    'external .xml dialog files
    Dim frmNewDialog As frmCustom 'object variable used to store the result window
    Dim strLineInput As String 'string variable used to store the current line imported from the dialog file
    Dim i As Integer 'a counter variable
    Dim intFL As Integer 'variable holding the file index
    
    'If the path of the file includes : it is an absolute path(for example C:\test.xml)
    'if it doesn't it must be a relative path
    If Not InStr(1, FileName, ":") > 0 Then
        'not an absolute path
        'it's relative to App.Path
        FileName = App.Path & "\" & FileName
    End If
    'if the filename doesn't exist
    If Not FS.FileExists(FileName) Then
        '(File not found)
        'do not execute the dialog file
        Exit Sub
    End If
    'this variable is used to store if the dialog file is going to be modal
    BolNextModal = False
    'there may be errors in the code of the .xml file
    'create error trap
    On Error GoTo lbShowError
    'load a new instance of frmCustom
    If Window Is Nothing Then
        Set frmNewDialog = New frmCustom
    Else
        Set frmNewDialog = Window
    End If
    Dim XDialog As DOMDocument
    Dim subElement As IXMLDOMNode
    Set XDialog = New DOMDocument
    If Not XDialog.Load(FileName) Then
        MsgBox "Unable to load the dialog `" & FileName & "' into memory." & vbNewLine & _
               "The XML file of the dialog is invalid.", vbCritical, "Invalid XML File"
        Exit Sub
    End If
    executeDialog XDialog.documentElement, frmNewDialog, 0
    For Each subElement In XDialog.documentElement.childNodes
        executeDialog subElement, frmNewDialog, 0
    Next subElement
lbTerminateLoading:
    'close the dialog file
    Close #intFL
    frmNewDialog.ActiveDialog = GetFileName(FileName)
    'show the projected dialog file (either modally or not)
    frmNewDialog.Show IIf(frmNewDialog.Special, vbModal, vbModeless)
    'don't display any errors
    Exit Sub
lbShowError:
    Resume lbTerminateLoading
End Sub
Public Sub executeDialog(ByVal XMLElement As IXMLDOMNode, ByRef frmDialog As Form, intLineCount As Integer)
    'this sub executes one single tag of an XML file
    
    'Syntax Explanation:
    'Numeric Arguments are enclosed in Parenthesis (argument)
    'String Arguments are enclosed in Single Quotes 'argument'
    'Arguments that their type depents on the other arguments
    'are not enclosed into special characters
    'Any Optional arguments are enclosed in Brakets [argument]
    'Every Statement is seperated from the other by
    'a full line brake(CRLF)
    'The arguments are seperated by spaces
    'Each Statament seems like this
    'Statament [Argument1 [Argument 2 [Argument 3 [...]]]]
    
    Dim objTemp As Object 'object variable used to store the object whose properties are changing or whose methods are being used
    Dim intIndex As Integer 'the index of the object
    Dim i As Integer
    Dim strAttribute As String
    Dim strValue As String
     
    'get the object type
    Select Case Strings.LCase$(XMLElement.nodeName)
        'these are the same objects are for initialise
        Case "static"
            'the object whose properties we are going to set is lblCustom(Index)
            Set objTemp = frmDialog.lblCustom
        Case "button"
            Set objTemp = frmDialog.cmdCustom
        Case "input"
            Set objTemp = frmDialog.txtCustom
        Case "image"
            Set objTemp = frmDialog.imgCustom
        Case "web"
            Set objTemp = frmDialog.wbCustom
        
        'this is a special object used to set the window's properties
        Case "window"
            Set objTemp = frmDialog
            GoTo SetAttributes
        Case Else
            'the object type passed is invalid, display warning
            DialogError "(XML Dialog File)", "Invalid Node XML Dialog Object Type", Language(185) & " `" & _
                            XMLElement.nodeName & "'", XMLElement.xml
            Exit Sub
    End Select
    Do
        On Error Resume Next
        If Len(objTemp.Item(i).Name) And False Then
            'the above if line will always return False.
            'this will cause the code bellow to be executed
            'only when an error occurs because of the
            'Resume Next error handler
            Load objTemp.Item(i)
            Set objTemp = objTemp.Item(i)
            objTemp.Visible = True
            Exit Do
        End If
        i = i + 1
    Loop
    
SetAttributes:
    For i = 0 To XMLElement.Attributes.length - 1
        strAttribute = XMLElement.Attributes.Item(i).nodeName
        strValue = DoReplace(XMLElement.Attributes.Item(i).Text)
        'Special Attributes/Properties
        Select Case Strings.LCase$(strAttribute)
            Case "source"
                objTemp.Picture = LoadPicture(strValue)
            Case "zorder"
                objTemp.ZOrder strValue
            Case "navigate"
                objTemp.Navigate2 strValue
            Case "event"
                objTemp.Tag = strValue
            Case Else
                On Error GoTo Invalid_Attribute
                CallByName objTemp, strAttribute, VbLet, strValue
        End Select
    Next i
    Exit Sub
Invalid_Attribute:
    MsgBox "The XML Attribute `" & strAttribute & "' of the Object `" & XMLElement.nodeName & "' is invalid." & vbNewLine & Err.Description, vbCritical, "XML Error"
End Sub
Private Sub DialogError(Optional ByVal strFileName As String = "(Unknown)", Optional strSimpleError As String = vbNullString, Optional strProgrammer As String = "Unspecified Error", Optional XMLElement As String = "(not specified)")
    'sub used to show a warning message when an error occurs
    'in dialog files
    MsgBox Language(190) & " " & strFileName & vbNewLine & strSimpleError & vbNewLine & vbNewLine & strProgrammer & vbNewLine & "XML Element: " & XMLElement, vbCritical, Language(189) & " `LoadDialog'"
End Sub
Public Sub LoadWizard(ByVal WizardName As String)
    Dim frmWizard As frmCustom
    
    Set frmWizard = New frmCustom
    LoadDialog App.Path & "\data\dialogs\wizard.xml", frmWizard
    frmWizard.wbCustom(1).Navigate2 App.Path & "\data\html\dialogdata\wizards\wizard.html"
    frmWizard.DialogData = WizardName
    frmWizard.DialogData2 = 0

    frmWizard.Show
End Sub
Public Sub Wizard_CallBack(ByRef frmWizard As frmCustom, Optional ByVal EventID As Byte = 0)
    Dim webdocWizard As HTMLDocument
    Dim webdocWizardPage As HTMLDocument
    
    'fired when wbCustom(1)_DocumentComplete is fired
    Set webdocWizard = frmWizard.wbCustom(1).Document
    
    Set webdocWizardPage = webdocWizard.parentWindow.frames(0).frames(1).Document
    
    If frmWizard.DialogData2 = 0 Then
        'load the first page
        Wizard_LoadPage frmWizard, 1
        Exit Sub
    End If
    
    Select Case LCase$(frmWizard.DialogData)
        Case "welcome"
            'loaded a page of the welcome wizard
            Select Case frmWizard.DialogData2
                Case 0
                Case 1
                    xNodeTag webdocWizardPage, "lang_welcomewizard_welcome2", Language(763)
                    xNodeTag webdocWizardPage, "lang_welcome", Replace(Language(285), "%1", VERSION_CODENAME & " " & App.Major & "." & App.Minor)
                Case 2
                    'TO DO: Check those Lang Entries replacement
                    xNodeTag webdocWizardPage, "lang_welcomewizard_importsettings", Language(764)
                    xNodeTag webdocWizardPage, "lang_importsettings", Language(768)
                    xNodeTag webdocWizardPage, "lang_importsettings_title", Language(747)
                    xNodeTag webdocWizardPage, "import_settings", frmWizard.DialogData3, , "checked"
                Case 3
                    xNodeTag webdocWizardPage, "lang_welcomewizard_importsettings2", Language(765)
                    xNodeTag webdocWizardPage, "lang_importsettings_title", Language(747)
                    xNodeTag webdocWizardPage, "import_settings2", WW_ImportSettingsFile, , "value"
                Case 4
                    xNodeTag webdocWizardPage, "lang_personal_settings", Language(766)
                    xNodeTag webdocWizardPage, "lang_nickname", Language(79)
                    xNodeTag webdocWizardPage, "lang_altnickname", Language(80)
                    xNodeTag webdocWizardPage, "lang_altnickname2", Language(81)
                    xNodeTag webdocWizardPage, "lang_realname", Language(128)
                    xNodeTag webdocWizardPage, "lang_email", Language(123)
                    xNodeTag webdocWizardPage, "lang_setdetails", Language(61)
                    xNodeTag webdocWizardPage, "nickname", WW_Nickname, , "value"
                    xNodeTag webdocWizardPage, "altnickname", WW_Alt, , "value"
                    xNodeTag webdocWizardPage, "altnickname2", WW_Alt2, , "value"
                    xNodeTag webdocWizardPage, "realname", WW_RealName, , "value"
                    xNodeTag webdocWizardPage, "email", WW_Email, , "value"
                Case 5
                    xNodeTag webdocWizardPage, "lang_ww_finished", Language(767) & "<br>" & Language(857)
                    xNodeTag webdocWizardPage, "lang_finished", Language(770)
            End Select
    End Select
End Sub
Public Sub Wizard_Finished(ByRef frmWizard As frmCustom)
    'finished
    
    Select Case LCase$(frmWizard.DialogData)
        Case "welcome"
            'check to see if we need to import settings
            If Len(WW_ImportSettingsFile) > 0 Then
                frmOptions.ImportSettings WW_ImportSettingsFile
            Else
                frmOptions.txtNickname.Text = WW_Nickname
                frmOptions.txtAlt.Text = WW_Alt
                frmOptions.txtAltTwo.Text = WW_Alt2
                frmOptions.txtEmail.Text = WW_Email
                frmOptions.txtReal.Text = WW_RealName
                frmOptions.SaveAll
            End If
    End Select
    'and close wizard
    frmWizard.DialogScript "hidedialog"
End Sub
Public Sub Wizard_LoadPage(ByRef frmWizard As frmCustom, ByVal intPageIndex As Integer)
    Dim webdocWizard As HTMLDocument
    Dim webdocButtons As HTMLDocument
    Dim webdocWizardPage As HTMLDocument
    Dim boolImportSettingsChecked As Boolean
    Dim strFileImportSettings As String
    
    Set webdocWizard = frmWizard.wbCustom(1).Document
    Set webdocButtons = webdocWizard.parentWindow.frames(1).Document
    Set webdocWizardPage = webdocWizard.parentWindow.frames(0).frames(1).Document
    
    Select Case LCase$(frmWizard.DialogData)
        Case "welcome"
            Select Case intPageIndex
                Case 1
                    xNodeTagShow webdocButtons, "lang_next_dis", False
                    xNodeTagShow webdocButtons, "lang_next", True
                    xNodeTagShow webdocButtons, "lang_back_dis", True
                    xNodeTagShow webdocButtons, "lang_back", False
                    xNodeTagShow webdocButtons, "lang_finish", False
                Case 2 To 4
                    If intPageIndex = 3 Then
                        'if we're moving from page 2 to page 3
                        If frmWizard.DialogData2 = 2 Then
                            boolImportSettingsChecked = webdocWizardPage.All.Item("xnode_import_settings").Checked
                            frmWizard.DialogData3 = boolImportSettingsChecked
                            If Not boolImportSettingsChecked Then
                                'proceed to page #4 (skip #3)
                                Wizard_LoadPage frmWizard, 4
                                Exit Sub
                            End If
                        ElseIf frmWizard.DialogData2 = 4 Then
                            'we're moving backwards from 4 to 3
                            'we should skip #3 and move to #2
                            Wizard_LoadPage frmWizard, 2
                            Exit Sub
                        End If
                    ElseIf intPageIndex = 4 Then
                        If frmWizard.DialogData2 = 5 Then
                            'we're moving back from page 5 to page 4
                            If frmWizard.DialogData3 Then
                                'proceed to page #3 (skip #4)
                                Wizard_LoadPage frmWizard, 3
                                Exit Sub
                            End If
                        ElseIf frmWizard.DialogData2 = 3 Then
                            'we're moving from page 3 to 4
                            'first check if the import settings
                            'filename is valid
                            strFileImportSettings = webdocWizardPage.All.Item("xnode_import_settings2").value
                            If LenB(strFileImportSettings) = 0 Then
                                MsgBox Language(852), vbExclamation, Language(747)
                                Exit Sub
                            ElseIf Not FS.FileExists(strFileImportSettings) Then
                                MsgBox Language(851), vbExclamation, Language(747)
                                Exit Sub
                            End If
                            'file OK store it
                            
                            WW_ImportSettingsFile = strFileImportSettings
                            
                            'we should skip page #4, as #3
                            'is only shown if Import_Settings
                            'was checked
                            Wizard_LoadPage frmWizard, 5
                            Exit Sub
                        End If
                    End If
                    
                    xNodeTagShow webdocButtons, "lang_next_dis", False
                    xNodeTagShow webdocButtons, "lang_next", True
                    xNodeTagShow webdocButtons, "lang_back_dis", False
                    xNodeTagShow webdocButtons, "lang_back", True
                    xNodeTagShow webdocButtons, "lang_finish", False
                Case 5
                    If frmWizard.DialogData2 = 4 Then
                        'moving from 4 to 5
                        'store personal info
                        WW_ImportSettingsFile = vbNullString
                        WW_Nickname = webdocWizardPage.All.Item("xnode_nickname").value
                        WW_Alt = webdocWizardPage.All.Item("xnode_altnickname").value
                        WW_Alt2 = webdocWizardPage.All.Item("xnode_altnickname2").value
                        WW_RealName = webdocWizardPage.All.Item("xnode_realname").value
                        WW_Email = webdocWizardPage.All.Item("xnode_email").value
                        If LenB(WW_Nickname) = 0 Then
                            MsgBox Language(853), vbExclamation, Language(762)
                            Exit Sub
                        ElseIf LenB(WW_RealName) = 0 Then
                            MsgBox Language(854), vbExclamation, Language(762)
                            Exit Sub
                        ElseIf LenB(WW_Email) = 0 Then
                            MsgBox Language(855), vbExclamation, Language(762)
                            Exit Sub
                        ElseIf LenB(WW_Alt) = 0 Then
                            If MsgBox(Language(856), vbYesNo Or vbQuestion, Language(762)) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    xNodeTagShow webdocButtons, "lang_next_dis", False
                    xNodeTagShow webdocButtons, "lang_next", False
                    xNodeTagShow webdocButtons, "lang_back_dis", False
                    xNodeTagShow webdocButtons, "lang_back", True
                    xNodeTagShow webdocButtons, "lang_finish", True
            End Select
    End Select
    webdocWizard.parentWindow.frames(0).frames(1).location.href = App.Path & "/data/html/dialogdata/wizards/" & frmWizard.DialogData & "/" & intPageIndex & ".html"
    frmWizard.DialogData2 = intPageIndex
End Sub
Public Sub Wizard_NextPage(ByRef frmWizard As frmCustom)
    Wizard_LoadPage frmWizard, frmWizard.DialogData2 + 1
End Sub
Public Sub Wizard_PreviousPage(ByRef frmWizard As frmCustom)
    Wizard_LoadPage frmWizard, frmWizard.DialogData2 - 1
End Sub
'#</section>

'#<section id="skins">
Public Function LoadSkin(ByVal SkinFile As String) As NodeSkin
    Dim intFL As Integer
    Dim strTemp As String
        
    intFL = FreeFile
    If InStr(1, SkinFile, ":") <= 0 Then
        SkinFile = App.Path & "/data/skins/" & SkinFile
    End If
    
    LoadSkin.FileName = SkinFile
    
    'default values
    LoadSkin.Resize_LeftOffset = 90
    LoadSkin.Resize_TabFocusOffset = 90
    LoadSkin.Resize_TabNonFocusOffset = 111
    LoadSkin.FontName = "Arial"
    LoadSkin.Font2Name = "Arial"
    LoadSkin.Font3Name = "Arial"
    LoadSkin.FontSize = 12
    LoadSkin.Font2Size = 12
    LoadSkin.Font3Size = 12
    LoadSkin.Resize_TextRightOffset = 200
    LoadSkin.Resize_TextLeftOffset = 100

    On Error GoTo Set_Default
    Open SkinFile For Input Access Read Lock Write As #intFL
    
    'Skin the first line
    xLineInput intFL
    
    Do Until EOF(intFL)
        Line Input #intFL, strTemp
        strTemp = Strings.Trim$(strTemp)
        If Strings.Left$(strTemp, 1) <> "#" And Len(strTemp) > 0 Then
            SkinAssign LoadSkin, GetStatement(strTemp), GetParameter(strTemp)
        End If
    Loop
    Close #intFL
    
    If Len(LoadSkin.ButtonUp) Then
        Set LoadSkin.private_ButtonUp = LoadPicture(App.Path & "/skins/" & LoadSkin.ButtonUp)
    End If
    ThisSkin = LoadSkin
    Exit Function
Set_Default:
    'file not found...
    MsgBox Language(247), vbCritical, Language(166)
    SaveSetting App.EXEName, "Options", "Skin", App.Path & "\data\skins\default.skin"
    SkinFile = App.Path & "\data\skins\default.skin"
    MsgBox "Please restart Node", vbInformation
    End
End Function
Public Sub ApplySkin(ByRef Skin As NodeSkin)
    Dim bChkThreeDee As Byte
    Dim bTxtThreeDee As Byte
    Dim bFraThreeDee As Byte
    Dim lBackColor As Long 'BackColor = Color1
    Dim lForeColor As Long 'ForeColor = Color2
    Dim i As Integer
    Dim SkinFont As StdFont, Skin2Font As StdFont, Skin3Font As StdFont
    Dim OptionsIcons(12) As String
    Dim lblObject As OptionButton
    Dim objContent As Object
    Dim lnDebug As Integer
    Dim bRegisterCount As Byte
    Dim SkinName As String
    
    SkinName = GetFileName(App.Path & "/temp/" & Skin.FileName)
    SkinName = Left$(SkinName, Len(SkinName) - Len(".skin"))
    If Not FS.FolderExists(SkinName) Then
        FS.CopyFolder App.Path & "/data/skins/" & SkinName, App.Path & "/temp/currentskin"
    End If
    
    On Local Error GoTo Skin_Error
    lnDebug = 1
    bChkThreeDee = Abs(CInt(Skin.CheckBox3D))
    bTxtThreeDee = Abs(CInt(Skin.TextBox3D))
    bFraThreeDee = Abs(CInt(Skin.Frame3D))
    
    lnDebug = 2
    lBackColor = Skin.ForeGround
    lForeColor = Skin.BackGround
    Set SkinFont = New StdFont
    Set Skin2Font = New StdFont
    Set Skin3Font = New StdFont
    SkinFont.Name = Skin.FontName
    SkinFont.Bold = Skin.FontBold
    SkinFont.Italic = Skin.FontItalic
    SkinFont.Size = Skin.FontSize
    
    lnDebug = 3
    Skin2Font.Name = Skin.Font2Name
    Skin2Font.Bold = Skin.Font2Bold
    Skin2Font.Italic = Skin.Font2Italic
    Skin2Font.Size = Skin.Font2Size

    lnDebug = 4
    Skin3Font.Name = Skin.Font3Name
    Skin3Font.Bold = Skin.Font3Bold
    Skin3Font.Italic = Skin.Font3Italic
    Skin3Font.Size = Skin.Font3Size
    
    'this order is important
    'FIRST we want to load frmOptions
    'and THEN to load frmMain
    'so that frmMain can use
    'the previously loaded options
    lnDebug = 5
    
    DB.Enter "Debug Point 5"
    DB.X "Set BackColor/Font on frmOptions"
    frmOptions.BackColor = lBackColor
    frmOptions.tvOptions.Font = SkinFont
    
    DB.X "Checking if Skin is Menu-Skinning-Compatible"
    If Skin.MenuForegroundHOver <> 0 Then
        DB.X "...Yes"
        DB.X "Loading frmMain"
        For Each objContent In frmMain.Controls
            If (TypeOf objContent Is NodeMenu.nMenu) Then
                    DB.X "Skin updating NodeMenu " & objContent.Name
                    objContent.Color_Font_Active_Color = Skin.MenuForegroundHOver
                    If objContent Is frmMain.nmnuMain Then
                        objContent.Color_Font_Inactive_Color = Skin.MenuForegroundNormal
                    Else
                        objContent.Color_Font_Inactive_Color = Skin.MenuForegroundNormalVertical
                    End If
                    objContent.Color_Regular_Background = Skin.MenuBackgroundNormal
                    objContent.Color_Shape_Active_Color = Skin.MenuBackgroundHOver
                    objContent.Color_Shape_Inactive_Color_Horizontal = Skin.MenuBackgroundExpanded
                    objContent.Color_Seperator_BackColor = Skin.MenuSeperatorColor
                    If LenB(Skin.MenuShadow) > 0 Then
                        objContent.Shadow = App.Path & "/data/skins/" & Skin.MenuShadow
                    End If
            End If
        Next objContent
    Else
        DB.X "...No"
    End If
    
    DB.X "frmMain::Refreshing NodeMenus (loading now if skin not menus-compatible)"
    frmMain.nmnuMain.Refresh
    DB.Leave "Debug Point 5"
    
    lnDebug = 6
    'frmOptions.lvList.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.OptionsBackground)
    
    lnDebug = 7
    'the font will be loaded with the next tabs as well
    frmMain.tsTabs(0).Font = SkinFont
    frmMain.lstSuggestions.BackColor = lBackColor
    frmMain.lstSuggestions.ForeColor = lForeColor
    
    lnDebug = 71
    frmMain.picPanelResize.BackColor = Skin.BackGround
    On Error Resume Next
    frmMain.imgClosePanel.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.Icon_Close)
    frmMain.imgPanelBegin.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundStart)
    frmMain.imgPanel.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundMid)
    frmMain.imgPanelEnd.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundEnd)
    
    frmMain.imgClosePane.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.Icon_Close)
    frmMain.imgPaneBegin.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundStart)
    frmMain.imgPane.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundMid)
    frmMain.imgPaneEnd.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.PanelTitleBackgroundEnd)
    frmMain.imgMore.Picture = LoadPicture(App.Path & "/data/skins/" & ThisSkin.MorePic)
    
    On Error GoTo 0
    frmMain.lblPanelTitle.ForeColor = Skin.PanelTitleColor
    frmMain.lblPaneTitle.ForeColor = Skin.PanelTitleColor
    
    lnDebug = 72
    frmMain.ntsPanel.ForeColor = Skin.ForeGround
    frmMain.ntsPanel.ForeColorActive = Skin.HotColor
    frmMain.ntsPanel.ForeColorDisabled = Skin.BackGround
    frmMain.ntsPanel.ForeColorHot = Skin.HotColor
    frmMain.ntsPanel.BackColor = Skin.BackGround
    frmMain.ntsPanel.BackColorScroll = Skin.BackGround
    frmMain.ntsPanel.ScrollArrowColor = Skin.ForeGround
    
    lnDebug = 8
    For Each objContent In frmOptions.Controls
        If (TypeOf objContent Is CheckBox) Or _
           (TypeOf objContent Is OptionButton) Then
            objContent.Appearance = bChkThreeDee
            objContent.Font = SkinFont
        End If
        If (TypeOf objContent Is Frame) Then
            objContent.Appearance = bFraThreeDee
        End If
        If (TypeOf objContent Is TextBox) Then
            objContent.Appearance = bTxtThreeDee
            objContent.Font = Skin3Font
        End If
        If (TypeOf objContent Is Label) Or _
           (TypeOf objContent Is CheckBox) Or _
           (TypeOf objContent Is Frame) Or _
           (TypeOf objContent Is OptionButton) Then
            objContent.BackColor = lBackColor
            objContent.ForeColor = lForeColor
        End If
        If (TypeOf objContent Is Label) Then
            objContent.Font = SkinFont
            objContent.ForeColor = lForeColor
        End If
        If (TypeOf objContent Is CommandButton) Then
            objContent.Font = Skin2Font
        End If
        lnDebug = lnDebug + 1
    Next objContent
    
    Exit Sub
Skin_Error:
    If lnDebug = 5 Then
        'failed to load frmMain
        'probably there is something wrong
        'with the custom ActiveX interfaces
        '(see http://sourceforge.net/forum/forum.php?thread_id=1044032&forum_id=327488 for more information)
                
        Err.Clear
        
        'try to register prjNodeMenu
        'and prjNodeTabs if possible
        'if we haven't already tried before
        bRegisterCount = bRegisterCount + 1
        If bRegisterCount > 2 Then
            'we have already tried
            'terminate and display warning
            MsgBox Language(530), vbCritical, Language(531)
            End
        End If
        
        'try to register them
        On Error Resume Next
        RegisterDLL App.Path & "\misc\nodemenu\prjNodeMenu.ocx"
        RegisterDLL App.Path & "\misc\nodeTab\prjNodeTab.ocx"
        RegisterDLL App.Path & "\misc\PlugInsInterface\prjNodePlugInsInterface.dll"
        
        Reload
        
        'try to reload frmMain
        Resume
        
    Else
        MsgBox Replace(Replace(Language(532), "%1", "mdlNode.ApplySkin"), "%2", CStr(lnDebug)), vbCritical
        ReportBug
        End
    End If
End Sub
Public Sub SkinExecuteCodeBehind(ByVal strRoutine As String, Optional boolReset As Boolean = False)
    Dim intFL As Integer
    Dim strCodeBehind As String
    Dim i As Integer
    
    'only if CodeBehind is enabled run this routine
    If Options.EnableCodeBehind = False Then
        'Note: we don't use (Not Options.EnableCodeBehind) as it may return TRUE even if the
        'value Options.EnableCodeBehind is TRUE, because Options.EnableCodeBehind can be NULL.
        Exit Sub
    End If
    
    If boolReset Then
        If scCodeBehind Is Nothing Then
            Set scCodeBehind = New ScriptControl
            scCodeBehind.Language = "VBScript"
        End If
        InitSphereVariables scCodeBehind, False
        
        intFL = FreeFile
        On Error GoTo CodeBehind_NotFound
        Open App.Path & "/data/skins/" & ThisSkin.CodeBehind For Input Access Read Lock Write As #intFL
        Do Until EOF(intFL)
            strCodeBehind = strCodeBehind & xLineInput(intFL) & vbNewLine
        Loop
        Close #intFL
        
        scCodeBehind.AddCode strCodeBehind
    End If
    
    If Not scCodeBehind Is Nothing Then
        On Error GoTo CodeBehind_Error
        For i = 1 To scCodeBehind.Procedures.Count
            If Strings.LCase$(scCodeBehind.Procedures.Item(i).Name) = Strings.LCase$(strRoutine) Then
                scCodeBehind.ExecuteStatement strRoutine
                Exit For
            End If
        Next i
    End If
    
    Exit Sub
CodeBehind_Error:
    MsgBox Language(222) & " " & ThisSkin.CodeBehind & " " & Language(223) & vbNewLine & Err.Description, vbCritical, Language(224)
    Exit Sub
CodeBehind_NotFound:
    MsgBox Language(228) & " " & ThisSkin.CodeBehind & " " & Language(229) & " " & ThisSkin.FileName & " " & Language(230) & vbNewLine & "(" & Err.Description & ")", vbCritical, Language(224)
End Sub
Public Sub SkinAssign(ByRef Skin As NodeSkin, ByVal Property As String, ByVal value As String)
    'CallByName ThisSkin, Property, VbLet, Value
    'Exit Sub
    Select Case Strings.LCase$(Property)
        Case "codebehind"
            Skin.CodeBehind = value
        Case "checkbox3d"
            Skin.CheckBox3D = value
        Case "textbox3d"
            Skin.TextBox3D = value
        Case "frame3d"
            Skin.Frame3D = value
        
        Case "foreground"
            Skin.ForeGround = value
        Case "background"
            Skin.BackGround = value
        Case "hotcolor"
            Skin.HotColor = value
            
        Case "fontname"
            Skin.FontName = value
        Case "fontbold"
            Skin.FontBold = value
        Case "fontitalic"
            Skin.FontItalic = value
        Case "fontsize"
            Skin.FontSize = value
        
        Case "font2name"
            Skin.Font2Name = value
        Case "font2bold"
            Skin.Font2Bold = value
        Case "font2italic"
            Skin.Font2Italic = value
        Case "font2size"
            Skin.Font2Size = value
        
        Case "font3name"
            Skin.Font3Name = value
        Case "font3bold"
            Skin.Font3Bold = value
        Case "font3italic"
            Skin.Font3Italic = value
        Case "font3size"
            Skin.Font3Size = value
        
        Case "templatefile"
            Skin.TemplateFile = value
        Case "templatefilechannel"
            Skin.TemplateFileChannel = value
        Case "templatefilenicks"
            Skin.TemplateFileNicks = value
        Case "templatefileprivate"
            Skin.TemplateFilePrivate = value
        Case "templatefilemisc"
            Skin.TemplateFileMisc = value
        
        'Case "backgroundfile"
        '    Skin.BackgroundFile = Value
        Case "backgroundcolor"
            Skin.BackgroundColor = value
        Case "skinpic"
            Skin.SkinPic = value
        
        Case "icon_help"
            Skin.Icon_Help = value
        Case "icon_options"
            Skin.Icon_Options = value
        Case "icon_join"
            Skin.Icon_Join = value
        Case "icon_add"
            Skin.Icon_Add = value
        Case "icon_server"
            Skin.Icon_Server = value
        Case "icon_channel"
            Skin.Icon_Channel = value
        Case "icon_web"
            Skin.Icon_Web = value
        Case "icon_close"
            Skin.Icon_Close = value
        Case "icon_delete"
            Skin.Icon_Delete = value
        
        Case "morebutton_down"
            Skin.ButtonDown = value
        Case "morebutton_up"
            Skin.ButtonUp = value
            
        Case "menubackgroundnormal"
            Skin.MenuBackgroundNormal = value
        Case "menubackgroundhover"
            Skin.MenuBackgroundHOver = value
        Case "menuforegroundnormal"
            Skin.MenuForegroundNormal = value
        Case "menuforegroundhover"
            Skin.MenuForegroundHOver = value
        Case "menuforegroundnormalvertical"
            Skin.MenuForegroundNormalVertical = value
        Case "menubackgroundexpanded"
            Skin.MenuBackgroundExpanded = value
        Case "menuseperatorcolor"
            Skin.MenuSeperatorColor = value
        Case "menushadow"
            Skin.MenuShadow = value
            
        Case "panel_titlebackground_start"
            Skin.PanelTitleBackgroundStart = value
        Case "panel_titlebackground_mid"
            Skin.PanelTitleBackgroundMid = value
        Case "panel_titlebackground_end"
            Skin.PanelTitleBackgroundEnd = value
        Case "panel_titlecolor"
            Skin.PanelTitleColor = value
            
        Case "optionsbackground"
            Skin.OptionsBackground = value
        Case "splashdialog"
            Skin.SplashDialog = value
        Case "buddyinimage"
            Skin.BuddyInImage = value
        Case "buddyoutimage"
            Skin.BuddyOutImage = value
            
        Case "optioncategoryuser"
            Skin.OptionCategoryUser = value
        Case "optioncategoryservers"
            Skin.OptionCategoryServers = value
        Case "optioncategoryscripts"
            Skin.OptionCategoryScripts = value
        Case "optioncategorylogs"
            Skin.OptionCategoryLogs = value
        Case "optioncategoryperform"
            Skin.OptionCategoryPerform = value
        Case "optioncategoryignore"
            Skin.OptionCategoryIgnore = value
        Case "optioncategorylanguage"
            Skin.OptionCategoryLanguage = value
        Case "optioncategorybuddys"
            Skin.OptionCategoryBuddys = value
        Case "optioncategorymisc"
            Skin.OptionCategoryMisc = value
        Case "optioncategoryskin"
            Skin.OptionCategorySkin = value
        Case "optioncategorysmileys"
            Skin.OptionCategorySmileys = value
        Case "optioncategorysessions"
            Skin.OptionCategorySessions = value
        
        Case "resize_bottomoffset"
            Skin.Resize_BottomOffset = value
        Case "resize_topoffset"
            Skin.Resize_TopOffset = value
        Case "resize_rightoffset"
            Skin.Resize_RightOffset = value
        Case "resize_leftoffset"
            Skin.Resize_LeftOffset = value
        Case "resize_tabfocusoffset"
            Skin.Resize_TabFocusOffset = value
        Case "resize_tabnonfocusoffset"
            Skin.Resize_TabNonFocusOffset = value
        Case "resize_textleftoffset"
            Skin.Resize_TextLeftOffset = value
        Case "resize_textrightoffset"
            Skin.Resize_TextRightOffset = value
        Case "resize_textbottomoffset"
            Skin.Resize_TextBottomOffset = value
        Case "resize_bottomhiddenoffset"
            Skin.Resize_BottomHiddenOffset = value
        Case "treeexpanedimage"
            Skin.TreeExpanedImage = value
        Case "treecollapsedimage"
            Skin.TreeCollapsedImage = value
        Case "treesubnodeimage"
            Skin.TreeSubNodeImage = value
        Case "treeseperator"
            Skin.TreeSeperator = value
        Case "treesubnodeimage2"
            Skin.TreeSubNodeImage2 = value
        Case "morepic"
            Skin.MorePic = value
            
        Case Else
            MsgBox Language(216) & " `" & Property & "'.", vbCritical, Language(166)
            SaveSetting App.EXEName, "Options", "Skin", App.Path & "\data\skins\default.skin"
    End Select
End Sub
'#</section>

'#<section id="plugins">
Public Sub ListPlugIns()
    Dim i As Integer
    Dim plugInFile As File
    On Error Resume Next
    
    'Lists all installed plugins
    'add all the valid plugins
    For Each plugInFile In FS.GetFolder(App.Path & "/data/plugins").Files
        If Right$(LCase$(plugInFile.Name), Len(".dll")) = ".dll" Then
            'valid plugin - dll extension
            'add to list(do not display the extension)
            NumToPlugIn.Add NumToPlugIn.Count, plugInFile.Name
            If Not Err Then
                ReDim Preserve Plugins(NumToPlugIn.Count - 1)
                Plugins(NumToPlugIn.Count - 1).strName = Left$(plugInFile.Name, Len(plugInFile.Name) - Len(".dll"))
            End If
        End If
    Next plugInFile
End Sub
Public Function LoadPlugIn(ByVal FileName As String) As Boolean
    'Loads a PlugIn
    Dim PlugInLib As TypeLibInfo
    Dim objPlugIn As prjNPIN.clsLittleFinger
    Dim strInterfaceName As String
    'Dim LoadedPluginsCount As Integer
    Dim i As Integer
    Dim i2 As Integer
    Dim boolValidNodePlugin As Boolean
    Dim strPlugInPath As String
    Dim strPlugInName As String
    
    strPlugInPath = App.Path & "/data/plugins/" & FileName
    'LoadedPluginsCount = UBound(LoadedPlugin) + 1
    'ReDim Preserve LoadedPlugin(LoadedPluginsCount)
    On Error GoTo Invalid_Plugin
    Set PlugInLib = TLI.TypeLibInfoFromFile(strPlugInPath)
    
    'PlugInLib.Register
    
    boolValidNodePlugin = False
    
    On Error Resume Next
    'check to see if the plug in is a valid node plugin
    For i = 1 To PlugInLib.TypeInfos.Count
        For i2 = 1 To PlugInLib.TypeInfos(i).Interfaces.Count
            If PlugInLib.TypeInfos(i).Interfaces(i2).Name = "_PlugIn" Then
                'valid
                boolValidNodePlugin = True
                strInterfaceName = PlugInLib.TypeInfos(i).Interfaces(i2).Name
                GoTo Continue
            End If
        Next i2
    Next i
    
Continue:
    If Not boolValidNodePlugin Then
        Exit Function
    End If
    
    'RegisterDLL strPlugInPath
    'Try and create an instance of the plugin
    On Error Resume Next
    Set objPlugIn = CreateObject(PlugInLib.Name & ".PlugIn")
                
    'If we fail, try registering the DLL
    If Err Then
        'This is probably because the DLL isn't registered, so we try and register it
        If RegisterDLL(strPlugInPath) = True Then
            Err.Clear
            On Local Error Resume Next
            Set objPlugIn = CreateObject(PlugInLib.Name & ".PlugIn")
            If Err Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    'to do:
    'store the plugin information
    If Not AtLeastOnePluginLoaded Then
        AtLeastOnePluginLoaded = True
        Set NPINHost = New clsNPINInterface
    End If
    Set Plugins(NumToPlugIn(FileName)).objPlugIn = objPlugIn
    Plugins(NumToPlugIn(FileName)).boolLoaded = True
    
    objPlugIn.PluginInit NPINHost
    
    strPlugInName = Right$(FileName, Len(FileName) - Len("prjPlugIn"))
    strPlugInName = Left$(strPlugInName, Len(strPlugInName) - Len(".dll"))
    frmMain.AddStatus Replace(Language(627), "%1", strPlugInName) & vbNewLine, CurrentActiveServer
    frmMain.AddNews Replace(Language(627), "%1", strPlugInName)
    
    frmOptions.fgPlugins.TextMatrix(NumToPlugIn(FileName) + 1, 1) = Language(280)
    
    'success
    LoadPlugIn = True
Invalid_Plugin:
End Function
Public Function IsPluginLoaded(ByVal FileName As String) As Boolean
    IsPluginLoaded = Plugins(NumToPlugIn(FileName)).boolLoaded
End Function
Public Function UnloadPlugIn(ByVal FileName As String) As Boolean
    Set Plugins(NumToPlugIn("prjPlugIn" & FileName)).objPlugIn = Nothing
    Plugins(NumToPlugIn("prjPlugIn" & FileName)).boolLoaded = False
    UnloadPlugIn = True
End Function
'#</section>

'#<section id="soundschemes">
Public Function LoadSoundScheme(ByVal SoundSkinFile As String) As nodeSoundScheme
    Dim XMLss As DOMDocument
    Dim XMLElement As IXMLDOMElement
    Dim strSSVersion As String
    Dim i As Integer
    Dim strSSNameOfAttribute As String
    Dim strSSValueOfAttribute As String
    Dim ssResult As nodeSoundScheme
    Dim strFileName As String
    
    Set XMLss = New DOMDocument
    
    If Not XMLss.Load(SoundSkinFile) Then
        MsgBox Language(865), vbCritical
        Exit Function
    End If
    
    For i = 0 To XMLss.documentElement.Attributes.length - 1
        Select Case XMLss.documentElement.Attributes.Item(i).nodeName
            Case "version"
                strSSVersion = XMLss.documentElement.Attributes.Item(i).nodeValue
                If strSSVersion <> "0.33" Then
                    MsgBox Language(875), vbCritical
                    Exit Function
                Else
                    Exit For
                End If
        End Select
    Next i
    
    For Each XMLElement In XMLss.documentElement.childNodes
        For i = 0 To XMLElement.Attributes.length - 1
            Select Case XMLElement.Attributes.Item(i).nodeName
                Case "for"
                    strSSNameOfAttribute = XMLElement.Attributes.Item(i).nodeValue
                Case "file"
                    strSSValueOfAttribute = XMLElement.Attributes.Item(i).nodeValue
            End Select
        Next i
        SoundSchemeAssign strSSNameOfAttribute, strSSValueOfAttribute, ssResult
    Next XMLElement
    
    Set XMLElement = Nothing
    Set XMLss = Nothing
    
    strFileName = GetFileName(SoundSkinFile)
    ssResult.FileName = Left$(strFileName, Len(strFileName) - 4)
    
    LoadSoundScheme = ssResult
End Function
Public Sub SoundSchemeAssign(ByVal strAttribute As String, ByVal strValue As String, ByRef SoundScheme As nodeSoundScheme)
    Select Case LCase$(strAttribute)
        Case "connect"
            SoundScheme.SndConnect = strValue
        Case "disconnect"
            SoundScheme.SndDisconnect = strValue
        Case "halfop"
            SoundScheme.SndHalfop = strValue
        Case "dehalfop"
            SoundScheme.SndDehalfop = strValue
        Case "op"
            SoundScheme.SndOp = strValue
        Case "deop"
            SoundScheme.SndDeop = strValue
        Case "voice"
            SoundScheme.SndVoice = strValue
        Case "devoice"
            SoundScheme.SndDevoice = strValue
        Case "join"
            SoundScheme.SndJoin = strValue
        Case "part"
            SoundScheme.SndPart = strValue
        Case "kick"
            SoundScheme.SndKick = strValue
        Case "quit"
            SoundScheme.SndQuit = strValue
        Case "error"
            SoundScheme.SndError = strValue
        Case "start"
            SoundScheme.SndStart = strValue
        Case "end"
            SoundScheme.SndEnd = strValue
        Case "modechange"
            SoundScheme.SndModeChange = strValue
        Case "topicchange"
            SoundScheme.SndTopicChange = strValue
        Case "whisper"
            SoundScheme.SndWhisper = strValue
        Case "invitation"
            SoundScheme.SndInvitation = strValue
    End Select
End Sub
Public Sub ThisSoundSchemePlaySound(ByVal strSound As String)
    SoundSchemePlaySound ThisSoundScheme, strSound
End Sub
Public Sub SoundSchemePlaySound(ByRef SoundScheme As nodeSoundScheme, ByVal strSound As String)
    Dim strR As String
    
    Select Case strSound
        Case "connect"
            strR = SoundScheme.SndConnect
        Case "disconnect"
            strR = SoundScheme.SndDisconnect
        Case "halfop"
            strR = SoundScheme.SndHalfop
        Case "dehalfop"
            strR = SoundScheme.SndDehalfop
        Case "op"
            strR = SoundScheme.SndOp
        Case "deop"
            strR = SoundScheme.SndDeop
        Case "voice"
            strR = SoundScheme.SndVoice
        Case "devoice"
            strR = SoundScheme.SndDevoice
        Case "join"
            strR = SoundScheme.SndJoin
        Case "part"
            strR = SoundScheme.SndPart
        Case "kick"
            strR = SoundScheme.SndKick
        Case "quit"
            strR = SoundScheme.SndQuit
        Case "error"
            strR = SoundScheme.SndError
        Case "start"
            strR = SoundScheme.SndStart
        Case "end"
            strR = SoundScheme.SndEnd
        Case "modechange"
            strR = SoundScheme.SndModeChange
        Case "topicchange"
            strR = SoundScheme.SndTopicChange
        Case "whisper"
            strR = SoundScheme.SndWhisper
        Case "invitation"
            strR = SoundScheme.SndInvitation
    End Select
    
    If LenB(strR) > 0 Then
        PlaySound App.Path & "/data/sounds/" & SoundScheme.FileName & "/" & strR
    End If
End Sub
Public Sub PlaySound(ByVal strFileName As String)
    Dim i As Integer
    Dim wmpObjectToUse As Object
    Static WMP() As Object
    
    strFileName = Replace(strFileName, "/", "\")
    
    On Error GoTo Array_Not_Initialized
    For i = 0 To UBound(WMP)
        On Error GoTo 0
        Select Case WMP(i).PlayState
            Case mpPlaying, mpWaiting
            Case Else
                GoTo Use_This_Object
        End Select
    Next i
    'no free Windows Media Player object found
    'create one
    ReDim Preserve WMP(i)
    Set WMP(i) = CreateObject("MediaPlayer.MediaPlayer.1")
Use_This_Object:
    Set wmpObjectToUse = WMP(i)
Playback_Sound:
    wmpObjectToUse.FileName = strFileName
    wmpObjectToUse.Stop
    wmpObjectToUse.Open strFileName
    'depending on if auto-play is enabled this may cause an error
    'because if it is, .open would cause the file to be played
    'right away
    On Error Resume Next
    wmpObjectToUse.Play
    Exit Sub
Array_Not_Initialized:
    ReDim WMP(0)
    Set WMP(0) = CreateObject("MediaPlayer.MediaPlayer.1")
    Resume
End Sub
'#</section>

'#<section id="smileypacks">
Public Sub LoadSmileyPack(ByRef ReturnTo As nodeSmileyPack, ByVal SmileyPackFileName As String)
    Dim XMLDoc As DOMDocument
    Dim XMLDocElement As IXMLDOMElement
    Dim XMLSmileyElement As IXMLDOMElement
    Dim bolSmileyPackVersionOK As Boolean
    Dim bolIsSpecialSmiley As Boolean
    Dim strFileName As String
    Dim strShortcutText As String
    Dim i As Integer
    Dim intSmileysCount As Integer
    
    DB.Enter "LoadSmileyPack"
    
    DB.X "Loading Smiley Pack " & SmileyPackFileName
    
    Set XMLDoc = New DOMDocument
    If Not XMLDoc.Load(SmileyPackFileName) Then
        DB.X "Warning: Invalid smiley pack XML file, aborting"
        MsgBox Language(896), vbCritical, Language(887)
        Exit Sub
    End If
    
    Set XMLDocElement = XMLDoc.documentElement
    
    If XMLDocElement.nodeName <> "smileypack" Then
        DB.X "Warning: Root element is not `smileypack', aborting"
        MsgBox Language(897), vbCritical, Language(887)
        Exit Sub
    End If
    
    For i = 0 To XMLDocElement.Attributes.length - 1
        Select Case XMLDocElement.Attributes.Item(i).nodeName
            Case "version"
                If XMLDocElement.Attributes.Item(i).nodeValue <> "0.34" Then
                    DB.X "Warning: Smiley Pack is not created for this version of Node"
                    bolSmileyPackVersionOK = False
                    Exit For
                Else
                    bolSmileyPackVersionOK = True
                    Exit For
                End If
        End Select
    Next i
    
    If Not bolSmileyPackVersionOK Then
        'either there is no version attribute, or it's not 0.34
        DB.X "Warning: Version check failed, aborting"
        MsgBox Language(897), vbCritical, Language(887)
        Exit Sub
    End If
    
    'everything Ok, load the file
    
    ReDim ReturnTo.AllSmileys(0)
    ReDim ReturnTo.SpecialSmileys(0)
    
    For Each XMLSmileyElement In XMLDocElement.childNodes
        If XMLSmileyElement.nodeName <> "smiley" Then
            DB.X "Warning: Invalid XML element detected: " & XMLSmileyElement.nodeName & " (element ignored)"
        Else
            For i = 0 To XMLSmileyElement.Attributes.length - 1
                Select Case XMLSmileyElement.Attributes.Item(i).nodeName
                    Case "for"
                        bolIsSpecialSmiley = False
                        strShortcutText = XMLSmileyElement.Attributes.Item(i).nodeValue
                    Case "file"
                        strFileName = XMLSmileyElement.Attributes.Item(i).nodeValue
                    Case "specialsmiley"
                        bolIsSpecialSmiley = True
                        strShortcutText = XMLSmileyElement.Attributes.Item(i).nodeValue
                End Select
            Next i
            If bolIsSpecialSmiley Then
                intSmileysCount = UBound(ReturnTo.SpecialSmileys) + 1
                ReDim Preserve ReturnTo.SpecialSmileys(intSmileysCount)
                ReturnTo.SpecialSmileys(intSmileysCount).FileName = strFileName
                ReturnTo.SpecialSmileys(intSmileysCount).ShortcutText = strShortcutText
            Else
                intSmileysCount = UBound(ReturnTo.AllSmileys) + 1
                ReDim Preserve ReturnTo.AllSmileys(intSmileysCount)
                ReturnTo.AllSmileys(intSmileysCount).FileName = strFileName
                ReturnTo.AllSmileys(intSmileysCount).ShortcutText = strShortcutText
            End If
        End If
    Next XMLSmileyElement

    DB.Leave "LoadSmileyPack"
End Sub
Public Function SpecialSmiley(ByVal SmileyID As String) As String
    Dim strSmileyId As String
    Dim i As Integer
    
    strSmileyId = LCase$(SmileyID)
    
    If Options.UseSmileys Then
        For i = 1 To UBound(ThisSmileyPack.SpecialSmileys)
            If LCase$(ThisSmileyPack.SpecialSmileys(i).ShortcutText) = strSmileyId Then
                SpecialSmiley = img(App.Path & "/data/smileys/" & Options.SmileyPack & "/" & ThisSmileyPack.SpecialSmileys(i).FileName, ThisSmileyPack.SpecialSmileys(i).ShortcutText, True)
                Exit Function
            End If
        Next i
    End If
End Function
'Public Function Smiley(ByVal SRC As String) As String
    'Function used to generate the HTML code for a smiley
'    Smiley = HTML_OPEN & "img src=""" & App.Path & "\data\graphics\smileys\" & SRC & """" & HTML_CLOSE
'End Function

'#</section>

'#<section id="general">
'this sub is used to get the statement from a command
'actually used to remove the parameters from a full statement
'and return only the statement itself
Public Function GetStatement(ByVal strText As String)
    'create a fake statement called `this' and
    'add the rest of the command as parameters
    'then use GetParameter to get the first parameter
    'which is actually the statement.
    'This is done so statements containing spaces
    'are also supported if they are included in double quotes(")
    GetStatement = GetParameter("this " & strText)
End Function
Public Function GetParameter(ByVal strText As String, Optional ByVal intParameterIndex As Integer = 1, Optional ByVal LastParameter As Boolean = False, Optional ByVal IsUnicode As Boolean) As String
    'GetParameter Function
    '
    'Function to make it easier to get a
    'parameter from a statement/command
    'e.g. for command connect nana.irc.gr 7000
    'we only have to do GetParameter("connect nana.irc.gr 7000", 1) and it will return
    '"nana.irc.gr" and GetParameter("connect nana.irc.gr 7000", 2)
    'will return "7000"
    'parameters inside "double quotes" count as one parameter even if they contain spaces
    'for example splay C:\My Documents\file.mp3 will seperate the parameters to
    '1)C:\My and
    '2)Documents\file.mp3
    'on the other hand splay "C:\My Documents\file.mp3" will understand
    'that we are talking for one parameter
    'do not use double quotes in the values themselves
    'it can cause problems, like this line: echo "Everybody "here" is cool"
    'instead use: echo "Everybody 'here' is cool"
    'Quotes are not returned from the function(they are removed while parsing)
    'so a command like echo "Hello World!" will
    'only return Hello World!, not "Hello World!"
    '
    Dim strStatement As String 'string variable used to store the statement from whom we want to get a parameter
    Dim i As Integer, i2 As Integer 'two counter variables
    Dim InQuotes As Boolean 'Determines whether an argument is into guotes(") to avoid spaces in it normally counted as arguments' seperators
    Dim intNextSpacePos As Integer 'variable used to store the position of the next space in the string
    Dim strQuotes As String
    Dim strSpace As String
    Dim strTempString As String 'mozillagodzilla optimization
    
    If intParameterIndex = 0 Then
        GetParameter = GetStatement(strText)
        Exit Function
    End If
    
    If IsUnicode Then
        strQuotes = """" & ChrW$(0)
        strSpace = " " & ChrW$(0)
    Else
        strQuotes = """" ' """" = "
        strSpace = " "
    End If
    
    If Not CBool(InStr(1, strText, strQuotes)) Then
        GetParameter = GetParameterQuick(strText, intParameterIndex, LastParameter, IsUnicode)
        Exit Function
    End If
    
    'get the argument of the function, add a space to the end and store it to strStatement
    strStatement = strText & strSpace
    
    'i is used to store the parameter index we are in
    'go throught the text; start from parameter 1 and
    'go to the parameter we want to return plus one.
    For i = 1 To intParameterIndex + 1
CheckNextSpace:
        'we are not currently in quotation marks
        InQuotes = False
        'get the next space position
        intNextSpacePos = InStr(1, strStatement, strSpace)
        'if there are no more spaces in the string
        'the parameter we are asked for is invalid
        If InStr(1, strStatement, strSpace) <= 0 Then
            'raise error
            GoTo lbInvalidParameter
        End If
        'go throught the rest of the statement
        'character-to-character
        'i2 is used to store the current character index
        For i2 = 1 To Len(strStatement)
            'if the current character is quotation marks
            'mozillagodzilla's note: StrComp is faster than string1=string2
            strTempString = Strings.Mid$(strStatement, i2, Len(strQuotes))
            'Debug.Print "moz: " & StrComp(strTempString, strQuotes, vbTextCompare)
            If StrComp(strTempString, strQuotes, vbTextCompare) = 0 Then
                'either the quotation marks start or end here
                InQuotes = Not InQuotes
            End If
            'if the current character is a space
            If i2 = intNextSpacePos Then
                'and we are not in quotation marks...
                If Not InQuotes Then
                    'if this is the parameter we are looking for
                    If i = intParameterIndex + 1 Then
                        'we found the parameter
                        'if this is the "last parameter"...
                        If Not LastParameter Then
                            'return the string from the current position to the next space
                            'cut out everything else
                            strStatement = Strings.Left$(strStatement, InStr(1, strStatement, strSpace) - Len(strSpace) + IIf(IsUnicode, 1, 0))
                        End If
                        'if lastparameter was set we are going to return the hole string without cutting anything out
                        'use lastparameter only if the last parameter doesn't have quotes in it, or the quotes will be returned as well!
                        
                        GoTo FinishLoops
                    End If
                    strStatement = Strings.Right$(strStatement, Len(strStatement) - InStr(1, strStatement, strSpace) - IIf(IsUnicode, 1, 0))
                    'Go to next space and increase i
                    GoTo CheckNextSpaceInc
                
                'if we are in quotation marks...
                Else
                    'replace the space with the identifier $space so as not to count it as a space
                    'we are going to replace it back later
                    strStatement = Strings.Left$(strStatement, i2 - 2 - IIf(IsUnicode, 1, 0)) & Replace(strStatement, strSpace, "$space", i2 - Len(strSpace), 1)
                    GoTo CheckNextSpace 'Go to next space but do NOT increase i as we remove the space
                End If
            End If
        'move to the next character
        Next i2
CheckNextSpaceInc:
    'move to the next parameter
    Next i
FinishLoops:
    'if the first character of the return string is a quotation mark remove it
    If Strings.Left$(strStatement, Len(strQuotes)) = strQuotes Then
        strStatement = Strings.Right$(Strings.Left$(strStatement, Len(strStatement) - Len(strQuotes)), Len(strStatement) - Len(strQuotes) * 2)
    End If
    'replace $space with spaces again and return the result. Note that the parameter may contain $space itself, but it will also be replace with a space.
    GetParameter = Replace(strStatement, "$space", strSpace)
    'don't show any errors
    Exit Function
lbInvalidParameter:
    'there's no such parameter. Display warning.
    Err.Raise vbObjectError, Language(187) & " `GetParameter'", Language(188)
End Function
Public Function GetParameterQuick(ByVal strText As String, Optional ByVal intParameterIndex As Integer = 1, Optional ByVal LastParameter As Boolean = False, Optional ByVal IsUnicode As Boolean) As String
    On Error GoTo Invalid_Parameter_Index
    GetParameterQuick = Split(strText, " " & IIf(IsUnicode, ChrW$(0), vbNullString), IIf(LastParameter, intParameterIndex + 1, -1))(intParameterIndex)
    Exit Function
Invalid_Parameter_Index:
    Err.Raise vbObjectError + 3, "GetParameterQuick Function", "Invalid Parameter Index"
End Function
Public Function IsParameter(ByVal strText As String, Optional ByVal intParameterIndex As Integer = 1) As Boolean
    'Function used to determine if a parameter exists
    'if an error is caused by GetParameter it's not a valid parameter
    'create an error trap
    On Error Resume Next
    GetParameter strText, intParameterIndex
    'if there was no error it is a valid parameter
    'else it's not
    IsParameter = Err.Number = 0
End Function
'#</section>

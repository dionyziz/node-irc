VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlugIn"
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
Option Compare Text

Private objNode As prjNPIN.clsHostInterface

Private Type UserBeingHelped
    strNickname As String
    hpbPhase As IXMLDOMNode
End Type

Private Users() As UserBeingHelped

Private XMLFile As DOMDocument
Implements prjNPIN.clsLittleFinger
Private Sub Class_Initialize()
    ReDim Users(0)
    Set Users(0).hpbPhase = Nothing
    'load XML data file
    Set XMLFile = New DOMDocument
    If Not XMLFile.Load(App.Path & "/HelperBot/helperbot.xml") Then
        MsgBox "Could not load information XML file!", vbCritical
    End If
End Sub
Private Sub clsLittleFinger_IRCAction(ByVal ActionID As prjNPIN.NodePlugInsNotifications)
    'MsgBox "Action Event " & ActionID & " fired!", vbInformation, "Node Action Event"
End Sub
Private Function clsLittleFinger_PluginInit(objHostInterface As prjNPIN.clsHostInterface) As Boolean
    Set objNode = objHostInterface
End Function
Private Function clsLittleFinger_PluginOptions() As Boolean
    MsgBox "Sorry, but this PlugIn has no options. Thank you for using HelperBot.", vbInformation
End Function
Private Sub clsLittleFinger_Sending()
End Sub
Private Sub clsLittleFinger_SomeoneJoined(ByVal UserName As String, ByVal Channel As String)
    Dim UserID As Integer
    
    UserID = GetHpbIDFromUser(UserName)
    If UserID = -1 Then
        UserID = UBound(Users) + 1
        ReDim Users(UserID)
    End If
    
    Users(UserID).strNickname = UserName
        
    Set Users(UserID).hpbPhase = XMLFile.documentElement

    SendSectionText UserID
End Sub
Private Sub clsLittleFinger_Receiving(ByVal User As String, ByVal Text As String, Optional ByVal Channel As String)
    Dim UserID As Integer
    Dim strlang As String
    Dim subSection As IXMLDOMNode
    Dim xmlAttribute As IXMLDOMNode
    
    If Channel <> "" Then
        'only reply if talking privately
        Exit Sub
    End If
    
    UserID = GetHpbIDFromUser(User)
    If UserID = -1 Then
        'user not in channel
        'or the bot joined later
        'send him/her the welcome message
        'if he/she talks to us
        clsLittleFinger_SomeoneJoined User, Channel
        Exit Sub
    End If

    'check to see if the user tries to enter a new section
    For Each subSection In Users(UserID).hpbPhase.childNodes
        If subSection.nodeName = "section" Then
            For Each xmlAttribute In subSection.Attributes
                If xmlAttribute.nodeName = "enter" Then
                    If hpbTrim(Text) = xmlAttribute.nodeValue Then
                        'entering new section
                        Set Users(UserID).hpbPhase = subSection
                        SendSectionText UserID
                        Exit Sub
                    End If
                End If
            Next xmlAttribute
        End If
    Next subSection

    'not entering new section
    
    'check to see if the user wants to go to the parent section
    If hpbTrim(Text) = "back" Then
        'yes
        'go to parent section only if possible
        If Not Users(UserID).hpbPhase.parentNode.parentNode Is Nothing Then
            Set Users(UserID).hpbPhase = Users(UserID).hpbPhase.parentNode
            SendSectionText UserID
            Exit Sub
        End If
    End If
    'not entering parent section
    
    'trying to enter an invalid section
    'display warning of <default>
    'or, if not set, display the default warning
    For Each subSection In Users(UserID).hpbPhase.childNodes
        If subSection.nodeName = "default" Then
            'there is a default section
            'send its text to the user
            Set Users(UserID).hpbPhase = subSection
            SendSectionText UserID
            Set Users(UserID).hpbPhase = subSection.parentNode
            Exit Sub
        End If
    Next subSection
    
    SendText Users(UserID).strNickname, "This is an invalid option."
End Sub
Public Sub SendSectionText(ByVal UserID As Integer)
    Dim xmlElement As IXMLDOMNode
    Dim strText As String
    Dim xmlNode As IXMLDOMNode
    
    For Each xmlElement In Users(UserID).hpbPhase.childNodes
        If xmlElement.nodeName = "text" Then
            For Each xmlNode In xmlElement.Attributes
                If xmlNode.nodeName = "value" Then
                    SendText Users(UserID).strNickname, xmlNode.nodeValue
                End If
            Next xmlNode
        ElseIf xmlElement.nodeName = "hr" Then
            SendText Users(UserID).strNickname, "%[hr size='1' width='30px'%]"
        End If
    Next xmlElement
    
    If Not Users(UserID).hpbPhase Is XMLFile.documentElement Then
        On Error Resume Next 'no parent element
        If Not Users(UserID).hpbPhase.parentNode Is XMLFile.documentElement Then
            'SendText Users(UserID).strNickname, "%[hr size='1' width='30px'%]"
            SendText Users(UserID).strNickname, "%[small%] To go back, type %bback%b %[/small%]"
        End If
    End If
End Sub
Public Sub SendText(ByVal UserName As String, ByVal strText As String)
    Dim strResultText As String
    
    strResultText = Replace(strText, "%b", Chr(2))
    strResultText = Replace(strResultText, "%nick", UserName)
    strResultText = Replace(strResultText, "%mynick", objNode.MyNick)
    strResultText = Replace(strResultText, "%[", Chr(27))
    strResultText = Replace(strResultText, "%]", Chr(29))
    strResultText = Replace(strResultText, "%%", "%")
    strResultText = Replace(strResultText, "%'", """")
    
    objNode.TxtSend "/privmsg " & UserName & " :" & strResultText & vbCrLf
End Sub
Private Function hpbTrim(ByVal Text As String) As String
    Dim strReturn As String
    strReturn = Trim(LCase(Text))
    If (Left(Text, 1) = """" And Right(Text, 1) = """") Or (Left(Text, 1) = "'" And Right(Text, 1) = "'") Or (Left(Text, 1) = "`" And Right(Text, 1) = "'") Then
        strReturn = Mid(Text, 2, Len(Text) - 2)
    End If
    hpbTrim = strReturn
End Function
Private Function GetHpbIDFromUser(ByVal User As String) As Integer
    Dim i As Integer
    Dim strLUser As String
    
    strLUser = LCase(User)
    
    For i = 0 To UBound(Users)
        If LCase(Users(i).strNickname) = strLUser Then
            GetHpbIDFromUser = i
            Exit Function
        End If
    Next i
    
    GetHpbIDFromUser = -1
End Function

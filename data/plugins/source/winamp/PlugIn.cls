VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
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

Implements prjNPIN.clsLittleFinger
Private Sub clsLittleFinger_IRCAction(ByVal ActionID As prjNPIN.NodePlugInsNotifications)
    'MsgBox "Action Event " & ActionID & " fired!", vbInformation, "Node Action Event"
End Sub
Private Function clsLittleFinger_PluginInit(objHostInterface As prjNPIN.clsHostInterface) As Boolean
    Set objNode = objHostInterface
    'objNode.npinAddStatus "Winamp Plugin is up and running"
    Load frmOptions
    Main
End Function
Private Function clsLittleFinger_PluginOptions() As Boolean
    'there are no options used by this example plugin
    'just display a message
    'MsgBox "This should display the options dialog for this plugin!!", vbInformation, "Sample Plugin"
    frmOptions.Show
End Function
Private Sub clsLittleFinger_Receiving(ByVal User As String, ByVal Text As String, Optional ByVal Channel As String)
    '--
End Sub
Private Sub clsLittleFinger_Sending()
    '--
End Sub
Private Sub clsLittleFinger_SomeoneJoined(ByVal UserName As String, ByVal Channel As String)
    '--
End Sub

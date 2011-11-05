VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLittleFinger"
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

'IRC Actions
Public Enum NodePlugInsNotifications
    npinPrivMsg = 1 'someone is sending a message either to a channel we are in or privately to us
    npinJoin 'someone joins a channel we are in, or we join a new channel
    npinPart 'someone leaves a channel we are in, or we leave a channel
    npinQuit 'someone disconnects from the network we are connected to
    npinKick 'someone is kicked from one of the channels we are in, we kicked someone, or we were kicked by someone
    npinMode 'someone changes the modes of a channel we are in, or we changed the modes of a channel
    npinError 'an error caused us to disconnect from the server
    npinNames 'when joining a channel this event fills the nicklist of it
    npinNick 'someone that is in a channel we are in changes his/her nickname or we changed our nickname
    npinUseAltNick 'the nickname the user chose is in use; we are going to use an alternative one
    npinTopic 'the topic in a channel we are in changes(either by us or by another user) or we are informed about the topic when we join a channel
    npinTopicTime 'if we are being informed about the topic in a particular channel this action lets us know who set the topic and when he did it
    npinNotice 'we recieved a notice from the server
    npinInvite 'someone invited us to join a channel.
    npinChannellist 'this event occurs when using /list command to list the channels of the server we are connected to
    npinBanList 'this event occurs when using /mode #chan +b (without any arguments) to list the bans in a channel
    npinSpecialNotice 'some text describing a `special event', an event that is not in another category, has to appear in a channel
    npinBuddyName 'saving the real name of the buddy to his/her profile file
    npinNamesSpecial 'allows /names to be used
End Enum
Public Sub IRCAction(ByVal ActionID As NodePlugInsNotifications)
    'called when an IRC action(join, part, etc) occurs
End Sub
Public Function PluginInit(objHostInterface As clsHostInterface) As Boolean
    'used to initialize the plug in and connect it with the Node object
End Function
Public Function PluginOptions() As Boolean
    'used to display the options of the plugin
End Function
Public Sub Sending()
    'called before sending
End Sub
Public Sub SomeoneJoined(ByVal UserName As String, ByVal Channel As String)
    'called when someone joins a channel
End Sub
Public Sub Receiving(ByVal User As String, ByVal Text As String, Optional ByVal Channel As String)
    'called when we receive a message
End Sub


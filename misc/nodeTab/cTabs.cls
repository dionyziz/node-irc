VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

Public TabCaption           As String
Public TabLeft              As Long
Public TabTop               As Long
Public TabWidth             As Long
Public TabHeight            As Long
Public TabPicture           As StdPicture
Public TabEn                As Boolean
Public TabAutoSize          As Boolean
Public TabToolTip           As Boolean

Public Property Get TabIcon() As StdPicture
    Set TabIcon = TabPicture
End Property

Public Property Set TabIcon(ByVal NewTabPicture As StdPicture)
    Set TabPicture = NewTabPicture
End Property

Private Sub Class_Terminate()
    Set TabPicture = Nothing
End Sub

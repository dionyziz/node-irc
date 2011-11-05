Attribute VB_Name = "Subclass"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

' A hi-lo type for breaking them up
Private Type HILOWord
  loword As Integer
  hiword As Integer
End Type

' Used to store and retrieve process addresses and handles against a window's handle in
'  the internal Windows database.
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

' Copies memory blocks. I bet you never would have guessed.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
' We need this one to redirect our windows messages... I.e. subclass.
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
' Used to invoke the original/default process for a window.
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Used to get the parent of the control
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' Subclassing stuff
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const GWL_WNDPROC As Long = (-4)            ' Used by SetWindowLong to start subclassing
Private Const WM_CREATE As Long = &H1               ' Sent when a window or control is created.
Private Const WM_DESTROY As Long = &H2              ' Sent when a window is being... destroyed.

Public Sub subclassControl(aRTBName As RichTextBox)
  Dim origProc As Long
  
  ' Make sure that subclassing method doesn't have any typos before invoking it
  '  for real. At this point, we're not subclassing so it won't crash.
  GenericSubCProc 0, 0, 0, 0
  
  SetProp aRTBName.hwnd, "CtrlPtr", ObjPtr(aRTBName.Parent)
  ' Only subclass once.
  If GetProp(aRTBName.hwnd, "OrigWindowProc") = 0 Then
    origProc = SetWindowLong(aRTBName.hwnd, GWL_WNDPROC, AddressOf GenericSubCProc)
    SetProp aRTBName.hwnd, "OrigWindowProc", origProc
  End If
End Sub

Private Function GenericSubCProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  
  Dim origProc As Long            ' Original process address
  Dim aCtrlPtr As Long            ' a pointer to the control
  Dim aControl As CodeEdit        ' a dummy variable to use for the control
  Dim aHiLo    As HILOWord        ' Used to seperate the words.
  Dim aXY      As HILOWord        ' used for the child mousedown events
  
  ' Just used for when I am bug squashing
  If hwnd = 0 And uMsg = 0 And wParam = 0 And lParam = 0 Then Exit Function
  
  ' Get our original process address
  origProc = GetProp(hwnd, "OrigWindowProc")
  
  If uMsg = WM_DESTROY And origProc <> 0 Then
    ' Unhook our control
    SetWindowLong hwnd, GWL_WNDPROC, origProc
    RemoveProp hwnd, "OrigWindowProc"
    RemoveProp hwnd, "CtrlPtr"
    ' Invoke the default process
    GenericSubCProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
  ElseIf origProc <> 0 Then
    ' Invoke the default process
    GenericSubCProc = CallWindowProc(origProc, hwnd, uMsg, wParam, lParam)
    ' Send the message to our control anyways
    aCtrlPtr = GetProp(hwnd, "CtrlPtr")
    If aCtrlPtr <> 0 Then
        CopyMemory aControl, aCtrlPtr, 4&
        aControl.SubclassedMessage uMsg, wParam, lParam
        CopyMemory aControl, 0&, 4&
    End If
  Else
    ' Used as extra protection in case something totally weird happens and
    '  we lose the process address. This code will probably never be
    '  invoked, but I'm freaky-cautious when it comes to subclassing &
    '  hooking :-\
    GenericSubCProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  End If
End Function



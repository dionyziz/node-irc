Attribute VB_Name = "mCommonDialog"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private strfileName As OPENFILENAME

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Sub DialogFilter(WantedFilter As String)
    Dim intLoopCount As Integer
    strfileName.lpstrFilter = ""

    For intLoopCount = 1 To Len(WantedFilter)
        If Mid(WantedFilter, intLoopCount, 1) = "|" Then strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Chr(0) Else strfileName.lpstrFilter = _
        strfileName.lpstrFilter + Mid(WantedFilter, intLoopCount, 1)
    Next intLoopCount
    strfileName.lpstrFilter = strfileName.lpstrFilter + Chr(0)
End Sub

Public Function OpenCommonDialog(Optional strDialogTitle As String = "Open", Optional strFilter As String = "All Files|*.*", Optional strDefaultExtention As String = "*.*") As String
    Dim lngReturnValue As Long
    Dim intRest As Integer
    Dim i As Long
    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259)
    strfileName.nMaxFile = 260
    strfileName.flags = &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetOpenFileName(strfileName)
    strfileName.lpstrFile = Trim(strfileName.lpstrFile)
    i = Len(strfileName.lpstrFile)
    If i <> 1 Then
        OpenCommonDialog = Trim(strfileName.lpstrFile)
    Else
        OpenCommonDialog = ""
    End If
End Function

Public Function GetFileNametoSave(strFilter As String, strDefaultExtention As String, Optional strDialogTitle As String = "Save") As String
    Dim lngReturnValue As Long
    Dim intRest As Integer
    strfileName.lpstrTitle = strDialogTitle
    strfileName.lpstrDefExt = strDefaultExtention
    DialogFilter (strFilter)
    strfileName.hInstance = App.hInstance
    strfileName.lpstrFile = Chr(0) & Space(259)
    strfileName.nMaxFile = 260
    strfileName.flags = &H80000 Or &H4
    strfileName.lStructSize = Len(strfileName)
    lngReturnValue = GetSaveFileName(strfileName)
    GetFileNametoSave = strfileName.lpstrFile
End Function

        





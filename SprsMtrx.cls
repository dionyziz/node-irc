VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSparseMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit

'Private collection to store collections for each row
Private m_RowCollection As New Collection

'Returns the cell value for the given row and column
Public Property Get Cell(nRow As Integer, nCol As Integer)
    Dim ColCollection As Collection
    Dim value As Variant

    On Error Resume Next
    Set ColCollection = m_RowCollection(CStr(nRow))
    'Return empty value if row doesn't exist
    If Err Then Exit Property
    value = ColCollection(CStr(nCol))
    'Return empty value if column doesn't exist
    If Err Then Exit Property
    'Else return cell value
    Cell = value
End Property

'Sets the cell value for the given row and column
Public Property Let Cell(nRow As Integer, nCol As Integer, value As Variant)
    Dim ColCollection As Collection

    On Error Resume Next
    Set ColCollection = m_RowCollection(CStr(nRow))
    'Add row if it doesn't exist
    If Err Then
        Set ColCollection = New Collection
        m_RowCollection.Add ColCollection, CStr(nRow)
    End If
    'Remove cell if it already exists (errors ignored)
    ColCollection.Remove CStr(nCol)
    'Add new value
    ColCollection.Add value, CStr(nCol)
End Property


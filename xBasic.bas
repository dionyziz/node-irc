Attribute VB_Name = "xBasic"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'xBasic Module
'
'some usefull functions

'allow only declared variables to be used
Option Explicit

'Registry Declares
Public Const HKEY_CURRENT_USER = &H80000001
Private Const KEY_READ = &H20019
Private Const ERROR_MORE_DATA = 234
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const ERROR_SUCCESS = &H0
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_MULTI_SZ = 7

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

'Set Window Top Most
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'Get Mouse Cursor Position
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Get Windows Version
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Declare Function GetVersionExA Lib "kernel32" _
   (lpVersionInformation As OSVERSIONINFO) As Integer

'DLL
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Public Function GetVersion() As String
    'routine used to get Window's version
    Dim OSinfo As OSVERSIONINFO 'API type used to store the user's OS version
    
    'initialize the buffers
    OSinfo.dwOSVersionInfoSize = 148
    OSinfo.szCSDVersion = Strings.Space$(128)
    'recieve information
    GetVersionExA OSinfo

    With OSinfo
        'there are two platform types...
        Select Case .dwPlatformId
            'Old-Style Windows
            Case 1
                'there are three versions of old-style windows
                Select Case .dwMinorVersion
                    'win '95
                    Case 0
                        GetVersion = "Windows 95"
                    'win '98
                    Case 10
                        GetVersion = "Windows 98"
                    'win Me
                    Case 90
                        GetVersion = "Windows Mellinnium"
                End Select
            'and Windows NT technology
            Case 2
                'there are three versions of NT-style windows so far
                Select Case .dwMajorVersion
                    'NT 3.51
                    Case 3
                        GetVersion = "Windows NT 3.51"
                    'NT 4.0
                    Case 4
                        GetVersion = "Windows NT 4.0"
                    'New Technology, after Windows NT
                    Case 5
                        'if minor is 0 it's Win 2k
                        If .dwMinorVersion = 0 Then
                            GetVersion = "Windows 2000"
                        'else it must be Win XP or greater
                        Else
                            GetVersion = "Windows XP"
                        End If
                End Select
       End Select
   End With
End Function

'<section name="API">
'Sub xWindowTopMost
'[sets the window on the top of every other window]
Public Sub xWindowTopMost(ByVal Handle As Long, Optional ByVal TopMost As Boolean = True)
    SetWindowPos Handle, IIf(TopMost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

'[returns the MousePosition]
Public Sub GetMousePosition(ByRef OutX As Integer, ByRef OutY As Integer)
    Dim lpXY As POINTAPI
    'get information about where the cursor is
    GetCursorPos lpXY
    'and assign the byRef variables
    OutX = lpXY.X
    OutY = lpXY.Y
End Sub
'</section>

'Function xLet
'[sets a value to a variable and returns the value]
'the cpp = operator
Function xLet(ByRef Var, ByVal value) As Variant
    'assign the value
    Var = value
    'and return the result of the assignement
    xLet = Var
End Function

Function xMax(ByVal Value1 As Double, ByVal Value2 As Double) As Double
    xMax = IIf(Value1 > Value2, Value1, Value2)
End Function
Function xMin(ByVal Value1 As Double, ByVal Value2 As Double) As Double
    xMin = IIf(Value1 < Value2, Value1, Value2)
End Function

'Function UnEscape
'[returns the unescaped text(with the HTML escaping rules)]
'the javascript unescape function
Function UnEscape(ByVal strText As String) As String
    Dim JS As ScriptControl
    'initialize a new script control
    Set JS = New MSScriptControl.ScriptControl
    'we're going to use JS's unescape function
    JS.Language = "JavaScript"
    'call it
    UnEscape = JS.Eval("unescape('" & Replace(strText, "'", "\'") & "');")
End Function

'Function Escape
'[returns the escaped text(with the HTML escaping rules)]
'the javascript escape function
Function Escape(ByVal strText As String) As String
    Dim JS As ScriptControl
    'Here, we're using the same method as in UnEscape()...
    Set JS = New MSScriptControl.ScriptControl
    JS.Language = "JavaScript"
    Escape = JS.Eval("escape('" & Replace(strText, "'", "\'") & "');")
End Function

'Function xLineInput
'[inputs a line from the file passed as a parameter and returns it]
Public Function xLineInput(ByVal File As Integer) As String
    Dim strTemp As String
    'read a line, store it into strTemp...
    Line Input #File, strTemp
    'and return it
    xLineInput = strTemp
End Function

'Function PartCopy
'[inputs a hole file and prints it into another]
Public Sub PartCopy(ByVal intDestinationFile As Integer, ByVal intSourceFile As Integer)
    'go through the lines of the source file
    Do Until EOF(intSourceFile)
        'read one line from the sourcefile and print it to the destination file
        Print #intDestinationFile, xLineInput(intSourceFile)
    Loop
End Sub

'Executes the specified program
'(that's the VB shell statement but as an API it's faster and with less bugs)
Public Function xShell(ByVal strCommand As String, ByVal lCaller As Long)
    Dim strParameters As String
    
    If IsParameter(strCommand) Then
        strParameters = GetParameter(strCommand)
    End If
    ShellExecute lCaller, "open", GetStatement(strCommand), strParameters, App.Path, SW_SHOW
End Function

'[stops program execution for the specified amount of time]
Public Sub Wait(ByVal Seconds As Single)
    Dim t As Single 'the time to stop waiting
    t = Timer + Seconds
    'loop if we have to wait more
    While t > Timer
        'don't crash ;-P
        DoEvents
    Wend
End Sub

'<section name="string_manipulation">

'Function GetFileName
'[returns the filename from a filename including the path]
Public Function GetFileName(ByVal strPath As String) As String
    strPath = Replace(strPath, "/", "\")
    GetFileName = Strings.Right$(strPath, Len(strPath) - InStrRev(strPath, "\"))
End Function

'Get the word on the specified position
'For example in this text:
'This is a test phrase
'           ^
'           |
'          13
'
'If position is 13 the word returned is "test"
Public Function GetWord(ByVal Text As String, ByVal Position As Integer) As String
    Dim intCurrentStart As Integer 'the start position of the word in the string
    Dim intCurrentEnd As Integer 'the end position of the word in the string
    
    'get the start position, it's the position of the last space
    intCurrentStart = InStrRev(Strings.Left$(Text, Position), " ") + 1
    'get the end position, it's the position of the next space
    intCurrentEnd = InStr(Position, Text, " ")
    'if there's no next space...
    If intCurrentEnd <= 0 Then
        '...the end position is the end of the string
        intCurrentEnd = Len(Text) + 1
    End If
    'if there's no start position
    If intCurrentEnd <= intCurrentStart Then
        'there's no word to return
        GetWord = vbNullString
    'there's a word to return
    Else
        'return it
        GetWord = Strings.Mid$(Text, intCurrentStart, intCurrentEnd - intCurrentStart)
    End If
End Function

'[remove the word in position Position]
Public Function RemoveWord(ByVal Text As String, ByVal Position As Integer) As String
    Dim intCurrentStart As Integer 'the start position of the word to remove in the string
    Dim intCurrentEnd As Integer 'the end position of the word to remove in the string
    
    'get the start position; it's the last space again
    intCurrentStart = InStrRev(Strings.Left$(Text, Position), " ") + 1
    'get the end position; the next space
    intCurrentEnd = InStr(Position, Text, " ")
    'if there's no next space
    If intCurrentEnd <= 0 Then
        'the end of the word is the end of the string
        intCurrentEnd = Len(Text) + 1
    End If
    'Use RemovePhrase to remove the word and return the result
    RemoveWord = RemovePhrase(Text, intCurrentStart - 1, intCurrentEnd)
End Function

'Adds a string(Phrase) inside a bigger string(Text) in a certain position(Position)
'For example if text is:
'This is a source program
'         ^
'         |
'         10
'the phrase is "n open", and position is 10
'the result will be "This is an open source program"
'                             ^^^^^^
'                             these were added by the function
Public Function AddPhrase(ByVal Text As String, ByVal Phrase As String, ByVal Position As Integer)
    'if the position in the string is invalid...
    If Position > Len(Text) Then
        'add the phrase at the end
        AddPhrase = Text & Phrase
    ElseIf Position < 0 Then
        AddPhrase = Phrase & Text
    'if it's valid
    Else
        'add the phrase inside the string
        AddPhrase = Strings.Left$(Text, Position) & Phrase & Strings.Right$(Text, Len(Text) - Position)
    End If
End Function

'[removes the specified part of a string]
Public Function RemovePhrase(ByVal Text As String, ByVal StartPosition As Integer, ByVal EndPosition As Integer)
    'get the remaining parts of the string and return them
    RemovePhrase = Strings.Left$(Text, StartPosition) & Strings.Right$(Text, Len(Text) - EndPosition + 1)
End Function

'[determines whether a word is the beggining of another]
Public Function WordMatch(ByVal SmallWord As String, ByVal BigWord As String) As Boolean
    'if the small word is at the beginning of the end word return true; else, return false
    WordMatch = Strings.LCase$(SmallWord) = Strings.LCase$(Strings.Left$(BigWord, Len(SmallWord)))
End Function

'[removes leading tab/space(s) from a command]
'same as Lstrings.trim$()
'reprecated; use Lstrings.trim$ instead.
'Public Function RemoveTabs(ByVal strText As String) As String
'    Dim i As Integer 'the position in the text for the loop
'    For i = 1 To Len(strText)
'        If Strings.Mid$(strText, i, 1) <> " " And ascw(Strings.Mid$(strText, i, 1)) <> vbKeyTab Then
'            RemoveTabs = Strings.Right$(strText, Len(strText) - i + 1)
'            Exit Function
'        End If
'    Next i
'End Function

'</section>

'method used to sort a collection
'OLD
'Public Sub SortCollection(ByRef cCollectionToSort As Collection, Optional ByVal cCriteria As Collection, Optional ByRef cCollection2 As Collection, Optional ByRef cCollection3 As Collection)
'    Dim i As Integer 'the first position in the collection
'    Dim i2 As Integer 'the second position in the collection; always after i
'    Dim cTempCriteria As Collection
'    Dim boolC2Exists As Boolean
'    Dim boolC3Exists As Boolean
'    Dim boolCriteriaIsDifferent As Boolean
'    Dim t As Long
'
'    t = GetTickCount
'
'    boolCriteriaIsDifferent = True
'
'    On Error GoTo Stop_To_Debug
'    If cCriteria Is Nothing Then
'        Set cCriteria = cCollectionToSort
'        boolCriteriaIsDifferent = False
'    End If
'    boolC2Exists = Not cCollection2 Is Nothing
'    boolC3Exists = Not cCollection3 Is Nothing
'
'    Set cTempCriteria = cCriteria
'
'    'go through the items of the collection
'    For i = 1 To cCollectionToSort.Count - 1
'        'go through the items of the collection after i
'        For i2 = i + 1 To cCollectionToSort.Count
'            'if i is higher than i2
'            If CompareEntries(cCriteria.Item(i), cCriteria.Item(i2)) Then
'                'swap the two entries
'                SwapEntries cCollectionToSort, i, i2
'                If boolCriteriaIsDifferent Then
'                    SwapEntries cCriteria, i, i2
'                End If
'                If boolC2Exists Then
'                    SwapEntries cCollection2, i, i2
'                End If
'                If boolC3Exists Then
'                    SwapEntries cCollection3, i, i2
'                End If
'            End If
'        Next i2
'    Next i
'
'    Set cCriteria = cTempCriteria
'
'    Debug.Print "SortCollection Old Time: " & GetTickCount - t
'
'    Exit Sub
'Stop_To_Debug:
'    Stop
'End Sub
Public Sub SortCollection2(ByVal cCollectionToSort As Collection)
    Dim AArrayToSort() As String
    
    CollectionToArray cCollectionToSort, AArrayToSort
    
'new sort code ---------------------

    mdlSort.QuickSortStrings AArrayToSort, LBound(AArrayToSort), UBound(AArrayToSort), 1, vbTextCompare
    
'new sort code ---------------------
    ArrayToCollection AArrayToSort, cCollectionToSort
End Sub
Public Sub SortCollection(ByRef cCollectionToSort As Collection, Optional ByVal cCriteria As Collection, Optional ByRef cCollection2 As Collection, Optional ByRef cCollection3 As Collection)
    Dim i As Integer 'the first position in the collection
    Dim i2 As Integer 'the second position in the collection; always after i
    Dim cTempCriteria As Collection
    Dim boolC2Exists As Boolean
    Dim boolC3Exists As Boolean
    Dim boolCriteriaIsDifferent As Boolean
    Dim AArrayToSort() As String
    Dim AArrayCriteria() As String
    Dim AArray2() As String
    Dim AArray3() As String
    
    boolCriteriaIsDifferent = True
    
    On Error GoTo Stop_To_Debug
    If cCriteria Is Nothing Then
        Set cCriteria = cCollectionToSort
        boolCriteriaIsDifferent = False
    End If
    boolC2Exists = Not cCollection2 Is Nothing
    boolC3Exists = Not cCollection3 Is Nothing
    
    CollectionToArray cCollectionToSort, AArrayToSort
    CollectionToArray cCriteria, AArrayCriteria
    
    If boolC2Exists Then
        CollectionToArray cCollection2, AArray2
    End If
    
    If boolC3Exists Then
        CollectionToArray cCollection3, AArray3
    End If

'still old sort code: new available in SortCollection2 ---------------------
'TO DO: update channel list sorting to work with SortCollection2

    'go through the items of the collection
    For i = 0 To UBound(AArrayToSort)
        'go through the items of the collection after i
        For i2 = i + 1 To UBound(AArrayToSort)
            'if i is higher than i2
            If CompareEntries(AArrayCriteria(i), AArrayCriteria(i2)) Then
                'swap the two entries
                SwapEntriesArray AArrayToSort, i, i2
                SwapEntriesArray AArrayCriteria, i, i2
                If boolC2Exists Then
                    SwapEntriesArray AArray2, i, i2
                End If
                If boolC3Exists Then
                    SwapEntriesArray AArray3, i, i2
                End If
            End If
        Next i2
        DoEvents 'to avoid ping timeout
    Next i
    
    ArrayToCollection AArrayToSort, cCollectionToSort
    
    If boolC2Exists Then
        ArrayToCollection AArray2, cCollection2
    End If
    
    If boolC3Exists Then
        ArrayToCollection AArray3, cCollection3
    End If
    
    Exit Sub
Stop_To_Debug:
    'Stop
End Sub
Public Sub SortCollection3(ByRef cCollectionToSort As Collection, ByRef cCollection2 As Collection, ByRef cCollection3 As Collection)
    Dim AArrayToSort() As String
    Dim AArray2() As String
    Dim AArray3() As String
    
    CollectionToArray cCollectionToSort, AArrayToSort
    CollectionToArray cCollection2, AArray2
    CollectionToArray cCollection3, AArray3

    MyQuickSort LBound(AArrayToSort), UBound(AArrayToSort), AArrayToSort, AArray2, AArray3
    
    ArrayToCollection AArrayToSort, cCollectionToSort
    ArrayToCollection AArray2, cCollection2
    ArrayToCollection AArray3, cCollection3
End Sub
'<quicksort>
'here's the quicksort function (the optimal sorting algorithm) in its general form:
'Public Sub MyQuickSort(ByVal Left As Integer, ByVal Right As Integer)
'    Dim Pivot As Integer
'    Dim i As Integer
'    Dim j As Integer
'
'    If Left < Right Then
'        i = Left
'        j = Right + 1
'        Pivot = Left
'        Do
'            Do
'                i = i + 1
'                If i > UBound(myArray) Then
'                    Exit Do
'                End If
'            Loop Until QCompare(i, Pivot) >= 0
'            Do
'                j = j - 1
'                If j < LBound(myArray) Then
'                    Exit Do
'                End If
'            Loop Until QCompare(j, Pivot) <= 0
'            If i < j Then
'                QSwap i, j
'            End If
'        Loop Until j <= i
'        QSwap Left, j
'        MyQuickSort Left, j - 1
'        MyQuickSort j + 1, Right
'    End If
'End Sub
'Public Sub QSwap(ByVal one As Integer, ByVal two As Integer)
'    Dim t As Integer
'
'    t = myArray(one)
'    myArray(one) = myArray(two)
'    myArray(two) = t
'End Sub
'Public Function QCompare(ByVal one As Integer, ByVal two As Integer) As Integer
'    If myArray(one) > myArray(two) Then
'        QCompare = 1
'        Exit Function
'    End If
'    If myArray(one) < myArray(two) Then
'        QCompare = -1
'        Exit Function
'    End If
'    QCompare = 0
'End Function
'and specifically for our arrays:
Public Sub MyQuickSort(ByVal Left As Integer, ByVal Right As Integer, ByRef a() As String, ByRef b() As String, ByRef c() As String)
    Dim Pivot As Integer
    Dim i As Integer
    Dim j As Integer
    
    If Left < Right Then
        i = Left
        j = Right + 1
        Pivot = Left
        Do
            Do
                i = i + 1
                If i > UBound(a) Then
                    Exit Do
                End If
            Loop Until QCompare(i, Pivot, a) >= 0
            Do
                j = j - 1
                If j < LBound(a) Then
                    Exit Do
                End If
            Loop Until QCompare(j, Pivot, a) <= 0
            If i < j Then
                QSwap i, j, a
                QSwap i, j, b
                QSwap i, j, c
            End If
        Loop Until j <= i
        QSwap Left, j, a
        QSwap Left, j, b
        QSwap Left, j, c
        MyQuickSort Left, j - 1, a, b, c
        MyQuickSort j + 1, Right, a, b, c
    End If
End Sub
Public Sub QSwap(ByVal one As Integer, ByVal two As Integer, ByRef a() As String)
    Dim t As String
    
    t = a(one)
    a(one) = a(two)
    a(two) = t
End Sub
Public Function QCompare(ByVal one As Integer, ByVal two As Integer, ByRef a() As String) As Integer
    If CompareEntries(a(one), a(two)) Then
        QCompare = 1
    Else
        QCompare = -1
    End If
End Function
'</quicksort>
Private Function CollectionToArray(ByRef cCollection As Collection, outAArray() As String)
    Dim i As Integer
    
    ReDim outAArray(cCollection.Count - 1)
    For i = 1 To cCollection.Count
        outAArray(i - 1) = cCollection.Item(i)
    Next i
End Function
Public Sub CollectionAppend(ByRef cCollection As Collection, ByVal strItem As String)
    'TO DO: Just fit the item in (so that we don't have to re-sort)
    
    cCollection.Add strItem
    SortCollection2 cCollection
End Sub
Private Function ArrayToCollection(ByRef AArray() As String, ByRef outCCollection As Collection)
    Dim i As Integer
    
    For i = 1 To outCCollection.Count
        outCCollection.Remove 1
    Next i
    
    For i = 0 To UBound(AArray)
        outCCollection.Add AArray(i)
    Next i
End Function
Private Sub SwapEntriesArray(ByRef AArray() As String, Entry1 As Integer, ByVal Entry2 As Integer)
    Dim Temp As String
    
    Temp = AArray(Entry1)
    AArray(Entry1) = AArray(Entry2)
    AArray(Entry2) = Temp
End Sub
Private Function CompareEntries(ByVal EntryOne As String, ByVal EntryTwo As String) As Boolean
    Dim lE1 As String
    Dim lE2 As String
    Dim dE1 As Double
    Dim dE2 As Double
    Dim boolNumeric As Boolean
    
    If Left$(EntryOne, 2) = "%@" Then
        EntryOne = Right$(EntryOne, Len(EntryOne) - 1)
    End If
    
    If Left$(EntryTwo, 2) = "%@" Then
        EntryTwo = Right$(EntryTwo, Len(EntryTwo) - 1)
    End If
        
    If Left$(EntryOne, 2) = "+%" Then
        EntryOne = Right$(EntryOne, Len(EntryOne) - 1)
    End If
    
    If Left$(EntryTwo, 2) = "+%" Then
        EntryTwo = Right$(EntryTwo, Len(EntryTwo) - 1)
    End If
    
    If Left$(EntryOne, 2) = "+@" Then
        EntryOne = Right$(EntryOne, Len(EntryOne) - 1)
    End If
    
    If Left$(EntryTwo, 2) = "+@" Then
        EntryTwo = Right$(EntryTwo, Len(EntryTwo) - 1)
    End If
    
    boolNumeric = IsNumeric(EntryOne) And IsNumeric(EntryTwo)
    If boolNumeric Then
        dE1 = Val(EntryOne)
        dE2 = Val(EntryTwo)
        CompareEntries = dE1 > dE2
    Else
        lE1 = LCase$(EntryOne)
        lE2 = LCase$(EntryTwo)
        If Left$(lE1, 1) = "@" And Left$(lE2, 1) <> "@" Then
            CompareEntries = False
        ElseIf Left$(lE1, 1) <> "@" And Left$(lE2, 1) = "@" Then
            CompareEntries = True
        Else
            CompareEntries = lE1 > lE2
        End If
    End If
End Function
Public Sub SwapEntries(ByRef cCollection As Collection, ByVal Entry1 As Integer, ByVal Entry2 As Integer)
    Dim i As Integer 'the position in the collection
    Dim cReturn As Collection 'the result collection
    
    'create a new collection
    Set cReturn = New Collection
    
    'go through the input collection
    For i = 1 To cCollection.Count
        'if this is the entry we have to swap
        If i = Entry1 Then
            'add the other entry instead
            cReturn.Add cCollection.Item(Entry2)
        'if this is the other entry we have to swap
        ElseIf i = Entry2 Then
            'add the first entry instead
            cReturn.Add cCollection.Item(Entry1)
        'it's a normal entry that we don't have to swap
        Else
            'simply add it
            cReturn.Add cCollection.Item(i)
        End If
    Next i
    'Copy the result collection into the source collection which is going to be returned
    'we don't want a reference so we'll have to copy data
    CopyCollection cReturn, cCollection
    'clear the "result" collection
    Set cReturn = Nothing
End Sub
Public Sub CopyCollection(ByRef SourceCollection As Collection, ByRef DestinationCollection As Collection)
    'copy items from Source collection to Destination collection
    'without refering from the one collection to the other
    Dim i As Integer 'the index in the collections
    
    'clear the destination collection
    ClearList DestinationCollection
    'go through the source collection
    For i = 1 To SourceCollection.Count
        'and copy all items from it to the destination collection
        DestinationCollection.Add SourceCollection.Item(i)
    Next i
End Sub
Public Sub AddEntry(ByVal strEntry As String, ByRef cCollection As Collection, ByVal intPosition As Integer)
    On Local Error Resume Next
    cCollection.Add strEntry, , intPosition
    If Err.Number <> 0 Then
        Err.Number = 0
        cCollection.Add strEntry, , intPosition - 1
        If Err.Number <> 0 Then
            cCollection.Add strEntry
        End If
    End If
End Sub
Public Sub ClearList(ByRef List As Collection)
    Dim i As Integer
    'go through the target collection
    For i = 1 To List.Count
        'and remove all entries
        List.Remove 1
    Next i
End Sub
Public Function Tween(ByVal StartValue As Integer, ByVal EndValue As Integer, ByVal Position As Double) As Integer
    'sub used to evaluate animation values
    Tween = StartValue + Position * (EndValue - StartValue)
End Function
'<section name="bytes">
Public Function UpperLowerToInt(ByVal Upper As Byte, ByVal Lower As Byte) As Integer
    'convert two bytes(lower/upper) to an integer
    Dim strUpperBinary As String
    Dim strLowerBinary As String
    strUpperBinary = ByteToBinary(Upper)
    strLowerBinary = ByteToBinary(Lower)
    UpperLowerToInt = BinaryToByte(strUpperBinary & FixLeadingZero(strLowerBinary, 8)) - 32768
End Function
Public Sub IntToUpperLower(ByVal intInput As Integer, ByRef OutUpper As Byte, ByRef OutLower As Byte)
    'convert an integer to two bytes(lower/upper)
    Dim strIntBinary As String
    strIntBinary = FixLeadingZero(ByteToBinary(CLng(intInput) + 32768), 16)
    OutUpper = BinaryToByte(Left$(strIntBinary, 8))
    OutLower = BinaryToByte(Right$(strIntBinary, 8))
End Sub
Public Function ByteToBinary(ByVal bInput As Long) As String
    'convert byte/integer/long to binary
    Dim strResult As String
    Dim bPower As Long
    Dim bRemaining As Long
    Dim i As Byte
    Dim bEx As Byte
    'get bEx
    Do
        i = i + 1
        If bInput - 2 ^ i < 0 Then
            bEx = i
            Exit Do
        End If
    Loop
    bRemaining = bInput
    'starting from the final bit, 128
    bPower = 2 ^ bEx
    '8 bits in each byte, bEx should be 8 for a byte
    For i = 0 To bEx
        'we will have to convert bPower to double, so as to do a double
        'evaluation and not a binary one
        If bRemaining - Val(bPower) >= 0 Then
            'this bit is true in the byte, binary 1
            strResult = strResult & "1"
            bRemaining = bRemaining - bPower
        Else
            'this bit is false in the byte, binary 0
            strResult = strResult & "0"
        End If
        'going to the lower bit
        bPower = bPower \ 2
    Next i
    ByteToBinary = strResult
End Function
Public Function BinaryToByte(strBinary As String) As Long
    'convert binary to byte/integer/long
    Dim lResult As Long
    Dim i As Byte
    Dim bEx As Byte
    'bEx should be 8 for a byte
    bEx = Len(strBinary)
    For i = 0 To bEx - 1
        lResult = lResult + Mid$(strBinary, bEx - i, 1) * (2 ^ i)
    Next i
    BinaryToByte = lResult
End Function
'</section>

Public Function EnumerateRegistryValues(ByVal lHKey As Long, ByVal strKeyName As String) As Collection
    'EnumerateRegistryValues Function
    'This function contains code based on
    'parts of code by Gregory Mazarakis
    
    Dim lHandle As Long
    Dim lValueType As Long
    Dim lNameLen As Long
    Dim lDataLen As Long
    Dim lIndex As Long
    Dim lRes As Long
    Dim strName As String
    Dim strRes As String
    Dim valueInfo(1) As Variant
    Dim lRetVal As Long
    Dim bResBinary() As Byte

    Set EnumerateRegistryValues = New Collection
    
    If LenB(strKeyName) > 0 Then
        'open registry key
        If CBool(RegOpenKeyEx(lHKey, strKeyName, 0, KEY_READ, lHandle)) Then
            'the return value is different from zero
            'an error occured
            Err.Raise vbObjectError, "EnumerateRegistryValues", "An error occured while trying to enumerate registry values: RegOpenKey() failed."
            Exit Function
        End If
    Else
        Err.Raise vbObjectError, "EnumerateRegistryValues", "Please provide a valid strKeyName parameter"
    End If
    
    Do
        lNameLen = 260
        strName = Space$(lNameLen)
        lDataLen = 4096
        ReDim bResBinary(lDataLen - 1)
        lRetVal = RegEnumValue(lHandle, lIndex, strName, lNameLen, ByVal 0&, lValueType, bResBinary(0), lDataLen)
        
        If lRetVal = ERROR_MORE_DATA Then
            ReDim bResBinary(lDataLen - 1)
            lRetVal = RegEnumValue(lHandle, lIndex, strName, lNameLen, ByVal 0&, lValueType, bResBinary(0), lDataLen)
        End If
        
        'there was an error, either directly after reading the value (other than ERROR_MORE_DATA)
        'or after getting More Data. This means that there aren't any more values to read
        'return the collection and exit function
        If lRetVal Then
            Exit Do
        End If
        
        valueInfo(0) = Left$(strName, lNameLen)
        
        Select Case lValueType
            Case REG_DWORD
                CopyMemory lRes, bResBinary(0), 4
                valueInfo(1) = lRes
            Case REG_SZ, REG_EXPAND_SZ
                strRes = Space$(lDataLen - 1)
                CopyMemory ByVal strRes, bResBinary(0), lDataLen - 1
                valueInfo(1) = strRes
            Case REG_BINARY
                If lDataLen < UBound(bResBinary) + 1 Then
                    ReDim Preserve bResBinary(lDataLen - 1)
                End If
                valueInfo(1) = bResBinary()
            Case REG_MULTI_SZ
                strRes = Space$(lDataLen - 2)
                CopyMemory ByVal strRes, bResBinary(0), lDataLen - 2
                valueInfo(1) = strRes
        End Select
        EnumerateRegistryValues.Add valueInfo, valueInfo(0)
        
        lIndex = lIndex + 1
    Loop
    
    RegCloseKey lHandle
End Function
Public Function RegisterDLL(ByVal strFileName As String) As Boolean
    Dim DLLTypeLib As TypeLibInfo
    Set DLLTypeLib = TLI.TypeLibInfoFromFile(strFileName)
    DLLTypeLib.Register
    Set DLLTypeLib = Nothing
    
    RegisterDLL = RegisterServer(strFileName)
End Function
Public Function GetFile(ByVal FileName As String) As String
    Dim intFL As Integer
    Dim strResult As String
    
    If FS.FileExists(FileName) Then
        intFL = FreeFile
        Open FileName For Input As intFL
        If Not EOF(intFL) Then
            Do Until EOF(intFL)
                strResult = strResult & xLineInput(intFL) & vbNewLine
            Loop
            strResult = Left$(strResult, Len(strResult) - 2)
        End If
        Close #intFL
        GetFile = strResult
    End If
End Function
Public Sub SetFile(ByVal FileName As String, Data As String)
    Dim intFL As Integer
    Dim strData() As String
    Dim i As Integer
    
    strData = Split(Data, vbLf)
    intFL = FreeFile
    Open FileName For Output As intFL
    For i = 0 To UBound(strData)
        Print #intFL, strData(i)
    Next i
    Close #intFL
End Sub
Public Function RegisterServer(ByVal DllServerPath As String, Optional ByVal hwnd As Long = 0, Optional ByVal bRegister As Boolean = True) As Boolean
    On Error Resume Next
    
    Dim lb As Long, pa As Long
    
    lb = LoadLibrary(DllServerPath)
    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If
    
    If CallWindowProc(pa, hwnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
        RegisterServer = True
    Else
        RegisterServer = False
    End If
    FreeLibrary lb
End Function
Public Function IsCompiled() As Boolean
    Static boolNotFirstCheck As Boolean
    Static boolLastReturn As Boolean
    
    If Not boolNotFirstCheck Then
        boolNotFirstCheck = True
        On Error GoTo vb_ide
        Debug.Print 1 / 0
        IsCompiled = True
        boolLastReturn = True
        Exit Function
vb_ide:
        IsCompiled = False
        boolLastReturn = False
    Else
        IsCompiled = boolLastReturn
    End If
End Function
Public Function ObjectCollectionItemExists(ByRef ObjectCollection As Object, ByVal IndexOfObject As Integer) As Boolean
    On Error GoTo No_Object
    If ObjectCollection(IndexOfObject).index Or True Then
        ObjectCollectionItemExists = True
    End If
    Exit Function
No_Object:
    ObjectCollectionItemExists = False
End Function
Public Function FixLeadingZero(ByVal strValue As String, Optional ByVal Digits As Integer = 2) As String
    On Error GoTo More_Digits_Than_Zeros
    FixLeadingZero = String$(Digits - Len(strValue), "0") & strValue
    Exit Function
More_Digits_Than_Zeros:
    FixLeadingZero = Right$(strValue, Digits)
End Function
Public Function DecodeFile(strInputFile As String, ByVal IsUnicode As Boolean) As String
    Dim bTemp() As Byte
    Dim fh As Long
    
    fh = FreeFile(0)
    Open strInputFile For Binary Access Read As fh
    
    ' Check for empty file and read the file
    If LOF(fh) Then
        ReDim bTemp(0 To LOF(fh) - 1)
        Get fh, , bTemp
    End If
    
    Close fh
    
    'Convert to a byte array then convert.
    'This is faster the repetitive calls to (w)ascw() or ChrW$()
    If IsUnicode Then
        DecodeFile = StrConv(bTemp, vbUnicode)
    Else
        DecodeFile = StrConv(bTemp, vbUnicode)
    End If
End Function
Public Function DecodedLineInput(ByRef strFileData As String)
    Dim intLineBreak As Integer
    Dim strReturn As String
    intLineBreak = InStr(1, strFileData, vbNewLine) 'StrConv(vbnewline, vbUnicode))
    If intLineBreak > 0 Then
        strReturn = Left$(strFileData, intLineBreak - 1)
        strFileData = Right$(strFileData, Len(strFileData) - Len(strReturn) - 2) '4
    Else
        strReturn = strFileData
        strFileData = vbNullString 'Right$(strFileData, Len(strFileData) - Len(strReturn))
    End If
    DecodedLineInput = strReturn
    Exit Function
Decode_Error:
End Function

Attribute VB_Name = "mdlObjList"
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Public Sub BuildFromTypeLib()
    Dim DLLTypeLib As TypeLibInfo
    Dim MyNode As MSComctlLib.Node
    Dim i As Integer

    'Get information from type library
    Set DLLTypeLib = TLI.TypeLibInfoFromFile(App.Path & "/../../node.exe")

    'create the central node
    Set MyNode = frmMain.tvObjects.Nodes.Add(, , , "Object Browser")

    'go through the classes of the program
    For i = 1 To DLLTypeLib.CoClasses.Count
        'for each class call AddClass()
        AddClass DLLTypeLib.CoClasses.Item(i), MyNode
    Next i
    
    Set DLLTypeLib = Nothing
End Sub
Public Sub AddClass(ByRef Element As CoClassInfo, ByRef tvNode As MSComctlLib.Node)
    Dim MyNode As MSComctlLib.Node
    Dim i As Integer
    
    'create a new node for this class
    Set MyNode = frmMain.tvObjects.Nodes.Add(tvNode, tvwChild, "~" & Element.Name, Element.Name, 1)
    
    'go through the interfaces of the class
    For i = 1 To Element.Interfaces.Count
        'for each interface, call AddInterface()
        AddInterface Element.Interfaces.Item(i), MyNode
    Next i
End Sub
Public Sub AddInterface(ByRef Element As InterfaceInfo, ByRef tvNode As MSComctlLib.Node)
    Dim MyNode As MSComctlLib.Node
    Dim i As Integer
    
    'we don't have to create a new node for each interface, we will add all
    'members from all interfaces of this class under the node of the class
    
    'go through the members of this interface of the class
    For i = 1 To Element.Members.Count
        'call AddMember() for each member of the interface
        AddMember Element.Members.Item(i), tvNode, Element.Parent.Name
    Next
End Sub
Public Sub AddMember(ByRef Element As MemberInfo, ByRef tvNode As MSComctlLib.Node, ByRef ClassName As String)
    'add a new node for this member under the node of the class
    On Error Resume Next 'dublicate key, so that we don't add an item twice
    frmMain.tvObjects.Nodes.Add tvNode, tvwChild, ClassName & "->" & Element.Name, Element.Name, GetImageFromInvokeKind(Element.InvokeKind)
End Sub
Public Function GetImageFromInvokeKind(ByVal InvokeKind As InvokeKinds)
    Select Case InvokeKind
        Case INVOKE_CONST, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF
            GetImageFromInvokeKind = 3
        Case INVOKE_FUNC, INVOKE_EVENTFUNC
            GetImageFromInvokeKind = 2
        Case INVOKE_UNKNOWN
            GetImageFromInvokeKind = 1
    End Select
End Function

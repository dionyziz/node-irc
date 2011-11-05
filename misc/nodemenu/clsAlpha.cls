VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Alpha Chanel Class
' 2002 Ed Preston - epreston@selectedsystems.com
'
Option Explicit
 
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function RedrawWindow Lib "user32" _
    (ByVal hwnd As Long, _
    lprcUpdate As RECT, _
    ByVal hrgnUpdate As Long, _
    ByVal fuRedraw As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
    
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2

Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_FRAME = &H400
Private Const RDW_INVALIDATE = &H1
Friend Sub SetLayered(ByVal hwnd As Long, ByVal bolSetAs As Boolean, ByVal bAlpha As Byte)
    ' Toggle layered and set the alpha chanel
    
    Dim nullRect As RECT
    Dim lret As Long
       
    lret = GetClassLong(hwnd, GWL_EXSTYLE)
    If bolSetAs = True Then
        lret = lret Or WS_EX_LAYERED
    Else
        lret = lret And Not WS_EX_LAYERED
    End If
    
    SetWindowLong hwnd, GWL_EXSTYLE, lret
    If bolSetAs Then
        SetLayeredWindowAttributes hwnd, 0, bAlpha, LWA_ALPHA
    Else
        RedrawWindow hwnd, nullRect, 0&, RDW_ALLCHILDREN Or RDW_ERASE Or RDW_FRAME Or RDW_INVALIDATE
    End If
End Sub
Friend Sub ReleaseDisplay(ByVal hwnd As Long)
    ' Release this layered window display
    SetLayered hwnd, False, 255
End Sub


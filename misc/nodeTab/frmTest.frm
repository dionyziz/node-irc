VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00E0F0F0&
   Caption         =   "Tab Test"
   ClientHeight    =   6630
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   557
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Change Color Scheme"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
   End
   Begin TabTest.xpTab xpWellsTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   10398
      Alignment       =   0
      TabHeight       =   22
      BackColor       =   14741744
      BackColorScroll =   14741744
      ForeColor       =   0
      ForeColorActive =   9982008
      ForeColorHot    =   16711680
      ForeColorDisabled=   12110024
      FrameColor      =   12110024
      ScrollArrowColor=   10794164
      MaskColor       =   16711935
      TabHotStripColor=   2658536
      SelectedTab     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberOfTabs    =   1
      TabPicture1     =   "frmTest.frx":0000
      AutoSize1       =   -1  'True
      TabWidth1       =   60
      TabText1        =   "Tab1"
      TabEnabled1     =   -1  'True
      Begin TabTest.xpTab xpWellsTab2 
         Height          =   4815
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   8493
         Alignment       =   1
         TabHeight       =   22
         BackColor       =   14741744
         BackColorScroll =   14741744
         ForeColor       =   0
         ForeColorActive =   9982008
         ForeColorHot    =   16711680
         ForeColorDisabled=   12110024
         FrameColor      =   12110024
         ScrollArrowColor=   10794164
         MaskColor       =   16711935
         TabHotStripColor=   2658536
         SelectedTab     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumberOfTabs    =   1
         AutoSize1       =   -1  'True
         TabWidth1       =   60
         TabText1        =   "Tab1"
         TabEnabled1     =   -1  'True
         Begin TabTest.xpTab xpWellsTab3 
            Height          =   4095
            Left            =   960
            TabIndex        =   4
            Top             =   360
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   7223
            Alignment       =   3
            TabHeight       =   22
            BackColor       =   14741744
            BackColorScroll =   14741744
            ForeColor       =   0
            ForeColorActive =   9982008
            ForeColorHot    =   16711680
            ForeColorDisabled=   12110024
            FrameColor      =   12110024
            ScrollArrowColor=   10794164
            MaskColor       =   16711935
            TabHotStripColor=   2658536
            SelectedTab     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumberOfTabs    =   3
            TabPicture1     =   "frmTest.frx":001C
            AutoSize1       =   -1  'True
            TabWidth1       =   60
            TabText1        =   "&AutoSize Property"
            TabEnabled1     =   -1  'True
            TabPicture2     =   "frmTest.frx":0038
            AutoSize2       =   -1  'True
            TabWidth2       =   60
            TabText2        =   "&Enabled Property"
            TabEnabled2     =   -1  'True
            TabPicture3     =   "frmTest.frx":0054
            AutoSize3       =   -1  'True
            TabWidth3       =   60
            TabText3        =   "Alignment &Property"
            TabEnabled3     =   -1  'True
            Begin TabTest.xpTab xpWellsTab4 
               Height          =   3015
               Left            =   480
               TabIndex        =   6
               Top             =   360
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   5318
               Alignment       =   2
               TabHeight       =   22
               BackColor       =   14741744
               BackColorScroll =   14741744
               ForeColor       =   0
               ForeColorActive =   9982008
               ForeColorHot    =   16711680
               ForeColorDisabled=   12110024
               FrameColor      =   12110024
               ScrollArrowColor=   10794164
               MaskColor       =   16711935
               TabHotStripColor=   2658536
               SelectedTab     =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumberOfTabs    =   3
               TabPicture1     =   "frmTest.frx":0070
               AutoSize1       =   -1  'True
               TabWidth1       =   60
               TabText1        =   "Right Align"
               TabEnabled1     =   -1  'True
               TabPicture2     =   "frmTest.frx":03C2
               AutoSize2       =   -1  'True
               TabWidth2       =   60
               TabText2        =   "Tab2"
               TabEnabled2     =   -1  'True
               TabPicture3     =   "frmTest.frx":0714
               AutoSize3       =   -1  'True
               TabWidth3       =   60
               TabText3        =   "Tab3"
               TabEnabled3     =   -1  'True
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Alignment"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   1575
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   6120
      Picture         =   "frmTest.frx":0A66
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim al As Long
Private Sub Command1_Click()
    xpWellsTab1.Alignment = al
    al = al + 1
    If al = 4 Then
        al = 0
    End If
End Sub

Private Sub Command2_Click()
Dim BackClr As OLE_COLOR
Dim FrameClr As OLE_COLOR
Dim ScrollClr As OLE_COLOR
Dim ArrowClr As OLE_COLOR
Dim NormalTxt As OLE_COLOR
Dim StripClr As OLE_COLOR
Dim ActiveTxt As OLE_COLOR
    BackClr = RGB(80, 104, 128)
    FrameClr = vbWhite
    ScrollClr = RGB(128, 128, 255)
    ArrowClr = vbWhite
    NormalTxt = RGB(192, 208, 216)
    
    Me.BackColor = BackClr
    
    xpWellsTab1.BackColor = BackClr
    xpWellsTab2.BackColor = BackClr
    xpWellsTab3.BackColor = BackClr
    
    xpWellsTab1.FrameColor = FrameClr
    xpWellsTab2.FrameColor = FrameClr
    xpWellsTab3.FrameColor = FrameClr
    
    xpWellsTab1.BackColorScroll = ScrollClr
    xpWellsTab2.BackColorScroll = ScrollClr
    xpWellsTab3.BackColorScroll = ScrollClr
    
    xpWellsTab1.ScrollArrowColor = ArrowClr
    xpWellsTab2.ScrollArrowColor = ArrowClr
    xpWellsTab3.ScrollArrowColor = ArrowClr
    
    xpWellsTab1.ForeColor = NormalTxt
    xpWellsTab2.ForeColor = NormalTxt
    xpWellsTab3.ForeColor = NormalTxt
    
    xpWellsTab1.TabHotStripColor = RGB(56, 112, 224)
    xpWellsTab2.TabHotStripColor = RGB(56, 112, 224)
    xpWellsTab3.TabHotStripColor = RGB(56, 112, 224)
    
    xpWellsTab1.ForeColorActive = vbWhite
    xpWellsTab2.ForeColorActive = vbWhite
    xpWellsTab3.ForeColorActive = vbWhite
End Sub

Private Sub Form_Load()
Dim i As Long
Dim j As Long
    al = 1
    For i = 1 To 20
        
        xpWellsTab1.AddTab ((56 * Rnd) + 5), , , False
    Next i
    j = 65
    For i = 1 To 25
        xpWellsTab2.AddTab
    Next i
    For i = 1 To 26
        xpWellsTab2.TabCaption(i) = Chr(j)
        j = j + 1
    Next i
End Sub

Private Sub xpWellsTab1_DblClick()
    MsgBox "In The Body", , "DblClick Event"
End Sub

Private Sub xpWellsTab1_TabDblClick(Index As Long)
Dim sCaption As String
    sCaption = InputBox("Rename Tab Number " & Index, "TabDblClick Event")
    If sCaption = "" Then
        Exit Sub
    Else
        xpWellsTab1.TabCaption(Index) = sCaption
    End If
End Sub

Private Sub xpWellsTab2_TabDblClick(Index As Long)
Dim result As Integer
    result = MsgBox("Disable Tab Number " & Index & " ?", vbYesNo, "Disable Tab")
    If result = vbYes Then
        xpWellsTab2.TabEnabled(Index) = False
    End If
End Sub


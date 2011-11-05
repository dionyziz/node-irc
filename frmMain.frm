VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{FC4172EC-F5A0-4265-B25E-9E25EF63128D}#1.0#0"; "prjNodeTab.ocx"
Object = "{26AD3DAD-35EF-4D74-92B0-D106F68C32EC}#94.0#0"; "prjNodeMenu.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00332222&
   Caption         =   "Node"
   ClientHeight    =   7950
   ClientLeft      =   3885
   ClientTop       =   3240
   ClientWidth     =   12030
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Node"
   ScaleHeight     =   7950
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLag 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   4560
      Top             =   1200
   End
   Begin VB.Timer tmrAway 
      Interval        =   60000
      Left            =   4560
      Top             =   720
   End
   Begin VB.Frame fraWebTab 
      BackColor       =   &H00C7C7C7&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtURL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Text            =   "(URL)"
         Top             =   0
         Width           =   5415
      End
   End
   Begin NodeMenu.nMenu nmnuMain 
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin VB.Frame fraPane 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   3360
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox picPaneResize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   2700
         MousePointer    =   9  'Size W E
         ScaleHeight     =   2535
         ScaleWidth      =   45
         TabIndex        =   17
         Top             =   0
         Width           =   50
      End
      Begin ComctlLib.TreeView tvConnections 
         Height          =   2415
         Left            =   0
         TabIndex        =   26
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   4260
         _Version        =   327682
         Indentation     =   706
         LabelEdit       =   1
         Style           =   7
         Appearance      =   0
      End
      Begin VB.Label lblPaneTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My Pane"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00854A34&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   810
      End
      Begin VB.Image imgClosePane 
         Height          =   165
         Left            =   2040
         MouseIcon       =   "frmMain.frx":1CFA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1E4C
         ToolTipText     =   "Close Panel"
         Top             =   10
         Width           =   180
      End
      Begin VB.Image imgPaneEnd 
         Height          =   195
         Left            =   0
         Picture         =   "frmMain.frx":201A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15
      End
      Begin VB.Image imgPaneBegin 
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":20AC
         Stretch         =   -1  'True
         Top             =   0
         Width           =   45
      End
      Begin VB.Image imgPane 
         Height          =   195
         Left            =   360
         Picture         =   "frmMain.frx":213E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   45
      End
   End
   Begin MSComctlLib.ImageList ilMenu 
      Left            =   5280
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2270
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2317
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":273E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2802
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2894
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2911
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2998
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrClearToolbar 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   4560
      Top             =   240
   End
   Begin VB.ListBox lstSuggestions 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   7560
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer tmrRefreshBuddy 
      Interval        =   30000
      Left            =   4080
      Top             =   240
   End
   Begin MSComctlLib.Toolbar tbMain 
      Height          =   540
      Left            =   1080
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   953
      ButtonWidth     =   979
      ButtonHeight    =   900
      Style           =   1
      ImageList       =   "ilTbMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add a smiley"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Bold"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Underlined"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Italic"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame fraMore 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   7080
      Width           =   255
      Begin VB.Image imgMore 
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape shpMore 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Timer tmrHideURLBar 
      Interval        =   2500
      Left            =   4080
      Top             =   720
   End
   Begin MSComctlLib.Toolbar tbText 
      Height          =   360
      Left            =   1080
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add a smiley"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Bold"
            Object.Tag             =   "Unpressed"
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Italic"
            Object.Tag             =   "Unpressed"
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Underline"
            Object.Tag             =   "Unpressed"
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Make text Colored"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Picture"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Hyperlink"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mutliline Message"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ilTbText 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3311
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3763
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E19
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrPanelRefreshSoon 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4080
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog cdPickAvatar 
      Left            =   5400
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "JPEG Images|*.jpg;*.jpeg"
      Flags           =   4
   End
   Begin VB.Frame fraPanel 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   9120
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Frame fraPanelTitle 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   200
         Left            =   50
         TabIndex        =   27
         Top             =   0
         Width           =   2295
         Begin VB.Label lblPanelTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "My Panel"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00854A34&
            Height          =   195
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   870
         End
         Begin VB.Image imgClosePanel 
            Height          =   165
            Left            =   2040
            MouseIcon       =   "frmMain.frx":80ED
            MousePointer    =   99  'Custom
            Picture         =   "frmMain.frx":823F
            ToolTipText     =   "Close Panel"
            Top             =   0
            Width           =   180
         End
         Begin VB.Image imgPanel 
            Height          =   195
            Left            =   15
            Picture         =   "frmMain.frx":840D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   45
         End
         Begin VB.Image imgPanelEnd 
            Height          =   195
            Left            =   1440
            Picture         =   "frmMain.frx":849F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   45
         End
      End
      Begin TabTest.xpTab ntsPanel 
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Alignment       =   3
         TabHeight       =   22
         BackColor       =   0
         BackColorScroll =   0
         ForeColor       =   0
         ForeColorActive =   9982008
         ForeColorHot    =   16711680
         ForeColorDisabled=   0
         FrameColor      =   0
         ScrollArrowColor=   0
         MaskColor       =   16711935
         TabHotStripColor=   2658536
         SelectedTab     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   161
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
      End
      Begin VB.PictureBox picPanelResize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   0
         MousePointer    =   9  'Size W E
         ScaleHeight     =   2535
         ScaleWidth      =   45
         TabIndex        =   13
         Top             =   0
         Width           =   50
      End
      Begin SHDocVwCtl.WebBrowser wbPanel 
         Height          =   2895
         Left            =   -30
         TabIndex        =   12
         Top             =   165
         Width           =   2325
         ExtentX         =   4110
         ExtentY         =   5106
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.Image imgPanelBegin 
         Height          =   195
         Left            =   0
         Picture         =   "frmMain.frx":85D1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Timer tmrRegularOperations 
      Interval        =   100
      Left            =   3600
      Top             =   240
   End
   Begin ComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   7695
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Welcome to Node"
            TextSave        =   "Welcome to Node"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTray 
      Height          =   495
      Left            =   6360
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin ComctlLib.TabStrip tsTabs 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Node IRC"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Timer tmrScriptingTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4080
      Top             =   2880
   End
   Begin MSComDlg.CommonDialog cdfile 
      Left            =   5400
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   4
   End
   Begin VB.Timer tmrShowSendToolTip 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3600
      Top             =   2400
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   350
      Left            =   240
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   609
      _Version        =   393217
      BackColor       =   16777215
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":8663
   End
   Begin SHDocVwCtl.WebBrowser wbStatus 
      Height          =   1935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   3375
      ExtentX         =   5953
      ExtentY         =   3413
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer tmrRefreshSoon 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   1200
   End
   Begin VB.Timer tmrMakeItQuicker 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   720
   End
   Begin SHDocVwCtl.WebBrowser wbBack 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      ExtentX         =   4683
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer tmrNDCAudioRecord 
      Interval        =   1000
      Left            =   4080
      Top             =   1200
   End
   Begin MSComctlLib.ImageList ilTbMain 
      Left            =   6960
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":86E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD43
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin NodeMenu.nMenu nmnuFile 
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuScript 
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuBrowse 
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuView 
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuHelp 
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin NodeMenu.nMenu nmnuIRC 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
   End
   Begin ComctlLib.ImageList ilTabs 
      Left            =   5280
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   10
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   20
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D175
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D307
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D499
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D62B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D7BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D94F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":DAE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E657
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":F1CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":FD43
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":108B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1142F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11FA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":12B1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13691
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14207
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":14D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":158F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16469
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16FDF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuNickPop 
      Caption         =   "&NickPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuMode 
         Caption         =   "&Mode"
         Begin VB.Menu mnuGiveOp 
            Caption         =   "Give Op"
         End
         Begin VB.Menu mnuTakeOp 
            Caption         =   "Take Op"
         End
         Begin VB.Menu mnuNicklistSeperator3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGiveHalfOp 
            Caption         =   "Give HalfOp"
         End
         Begin VB.Menu mnuTakeHalfOp 
            Caption         =   "Take HalfOp"
         End
         Begin VB.Menu mnuNicklistSeperator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGiveVoice 
            Caption         =   "Give Voice"
         End
         Begin VB.Menu mnuTakeVoice 
            Caption         =   "Take Voice"
         End
      End
      Begin VB.Menu mnuKick 
         Caption         =   "&Kick..."
      End
      Begin VB.Menu mnuNickListSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInformation 
         Caption         =   "&Information"
         Begin VB.Menu mnuInfo 
            Caption         =   "Nickname Info"
         End
         Begin VB.Menu mnuWhoIs 
            Caption         =   "WhoIs"
         End
         Begin VB.Menu mnuUserHost 
            Caption         =   "UserHost"
         End
         Begin VB.Menu mnuCTCP 
            Caption         =   "CTCP"
            Begin VB.Menu mnuCTCPVer 
               Caption         =   "Version"
            End
            Begin VB.Menu mnuCTCPTime 
               Caption         =   "Time"
            End
            Begin VB.Menu mnuCTCPPing 
               Caption         =   "Ping"
            End
         End
      End
      Begin VB.Menu mnuDCC 
         Caption         =   "&DCC"
         Begin VB.Menu mnuDCCSend 
            Caption         =   "Send File"
         End
         Begin VB.Menu mnuDCCChat 
            Caption         =   "Chat"
         End
      End
      Begin VB.Menu mnuNDC 
         Caption         =   "&NDC"
         Begin VB.Menu mnuNDCConnect 
            Caption         =   "&Connect"
         End
         Begin VB.Menu mnuNDCStartProgram 
            Caption         =   "Start Program"
            Begin VB.Menu mnuNDCStartNetMeeting 
               Caption         =   "NetMeeting"
            End
         End
         Begin VB.Menu mnuNDCAudio 
            Caption         =   "Start &Audio Conversation"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuNDCVideo 
            Caption         =   "Start Video Conversation"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuNickSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddBuddy 
         Caption         =   "Add Buddy"
      End
      Begin VB.Menu mnuNickIgnore 
         Caption         =   "Ignore"
      End
      Begin VB.Menu mnuNickListSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWhisper 
         Caption         =   "&Whisper"
      End
      Begin VB.Menu mnuNickListSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNickClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuNickViewLogs 
         Caption         =   "View &Logs"
      End
   End
   Begin VB.Menu mnuTabsPop 
      Caption         =   "&TabsPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCloseTab 
         Caption         =   "&Close Tab"
      End
      Begin VB.Menu mnuTabConnect 
         Caption         =   "&Connect..."
      End
      Begin VB.Menu mnuTabDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnuWebTabRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuWebTabStop 
         Caption         =   "&Stop"
      End
      Begin VB.Menu mnuWebTabBack 
         Caption         =   "&Back"
      End
      Begin VB.Menu mnuWebTabForward 
         Caption         =   "&Forward"
      End
      Begin VB.Menu mnuWebTabFav 
         Caption         =   "&Add to Favorites"
      End
      Begin VB.Menu mnuChanTabLeave 
         Caption         =   "&Leave"
      End
      Begin VB.Menu mnuChanTabRejoin 
         Caption         =   "Re&join"
      End
   End
   Begin VB.Menu mnuIRCPop 
      Caption         =   "&IRCPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuViewLogs 
         Caption         =   "View &Logs"
      End
      Begin VB.Menu mnuIRCPopSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChanProperties 
         Caption         =   "Channel &Properties..."
      End
      Begin VB.Menu mnuChanModes 
         Caption         =   "Channel Modes..."
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPanelTab 
      Caption         =   "PanelTab"
      Visible         =   0   'False
      Begin VB.Menu mnuClosePanelTab 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuConnectionsPop 
      Caption         =   "&ConnectionsPopUP"
      Visible         =   0   'False
      Begin VB.Menu mnuNewServer 
         Caption         =   "&New Server Connection"
      End
      Begin VB.Menu mnuRemoveServer 
         Caption         =   "&Remove Server Connection"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'Allow only declared variables
Option Explicit
'Make text comparison the default one
Option Compare Text

'Constants
Private Const WebBrowserIndex_Priv = 1
Private Const WebBrowserIndex_Chan = 2
Private Const WebBrowserIndex_DCC = 3
Private Const WebBrowserIndex_Loading = 4

Private Const wbLeft = 120 'the webbrowser left
Private Const wbTop = 480 'and top
Private Const txLeft = 120 'the textbox left

Private Const Panel_Title_Height = 200 'the height of the title of the panel on the right

'Scripting Object
Public ndScript As ScriptControl 'object variable used to execute scripts

Public PortToUse As Long 'this is the port number that is sent back from ChooseOpenPort
Public MsgRcvTxt As String 'this is the message that we received
Public CurrentNick As String 'the current selected nickname of a buddy in the nicklist
Private GiveFocus As Boolean 'give focus to the text after pressing tab or not?
Private MaxMode As Boolean 'are we working on full screen mode? (under construction)
Private currentWB As Integer 'the current(visible) webbrowser object index
Private MessageHistory() As String 'previous sent messages
Private NDCRandomCurrent As Integer
Private NDCGlobalEventID As Long
Public ScriptingTimerNextCall As String
Public ChanProps As frmCustom
Private BanListIndex As Integer
Public webdocChanFrameSet As HTMLDocument 'Channels' frameset
Public webdocChanMain As HTMLDocument 'Channels' main text
Public webdocChanNicklist As HTMLDocument 'Channels' nicklist
Public webdocChanTopic As HTMLDocument 'Channels' topic
Public webdocPrivates As HTMLDocument 'DOM document of Private Windows and the Status Window
Public webdocDCCs As HTMLDocument 'DOM document of the DCC windows
Private strDCCPresetFile As String
Private webdocTemp As HTMLDocument 'Tempoarary DOM document object
Private webdocWebSite() As HTMLDocument 'DOM documents storing the active web sites
Private webdocPanel As HTMLDocument 'DOM document storing the active panel
Private webdocProfile As HTMLDocument 'DOM document for the buddys profile
Private frmMinor As Form
Private frmOrganize As frmCustom '"Organize My Servers"
Private frmFullScreen As frmCustom 'The full screen form
Private frmChanModes As frmCustom 'Channel Modes (not Channel Topic/Ban)
Private boolSaveNextType As Boolean 'save the next text we type in the message history list?
Public boolRestoring As Boolean 'are we currently restoring a session?
Private ThisTip As NodeInfoTips
Private strLastBalloonInfo As String
Private e_KeyCode As Long

Public WithEvents xpBalloon As clsBalloon
Attribute xpBalloon.VB_VarHelpID = -1
Public WithEvents webdocFullScreen As HTMLDocument 'DOM document of the Full Screen HTML file.
Attribute webdocFullScreen.VB_VarHelpID = -1
Public WithEvents webdocFullScreenChanMain As HTMLDocument
Attribute webdocFullScreenChanMain.VB_VarHelpID = -1
Public WithEvents webdocFullScreenChanNicks As HTMLDocument
Attribute webdocFullScreenChanNicks.VB_VarHelpID = -1
Public WithEvents webdocCurrentIRCWindow As HTMLDocument
Attribute webdocCurrentIRCWindow.VB_VarHelpID = -1
Public webdocSplit As HTMLDocument
Private SplitEnabled As Boolean
Private SplitIndex As Integer

Private intPanelResizeStartX As Integer
Private intPaneResizeStartX As Integer
Private boolPanelResizing As Boolean
Private boolPaneResizing As Boolean
Public strCurrentPanel As String
Private IsPaneOpen As Boolean
Private LoadedPanels() As String 'keys of the loaded panels
Public namesboolean As Boolean 'so that /names command will work
Public showwhois As Boolean 'whether or not to show the whois messages
Public showison As Boolean 'whether or not to show the ison messages
Public LastMessage As String
Public IsActive As Boolean
Public MultipleInstances As Boolean
Public MultilineText As Boolean
Private boolBuildingPrimary As Boolean
Private ChannelsCount As Integer 'the number of the channels we have retrieved so far from the channel list
Private MinimumPanelWidth As Integer
Private BSCodeCall As Boolean
Private BolLstSuggstionsCodeClick As Boolean
Public AwayMins As Integer

Public WithEvents wsIdentD As clsWSArray
Attribute wsIdentD.VB_VarHelpID = -1
Public WithEvents wsNDC As clsWSArray
Attribute wsNDC.VB_VarHelpID = -1
Public WithEvents wsDCC As clsWSArray
Attribute wsDCC.VB_VarHelpID = -1
Public WithEvents wsDCCSend As clsWSArray
Attribute wsDCCSend.VB_VarHelpID = -1
Public WithEvents wsDCCChat As clsWSArray
Attribute wsDCCChat.VB_VarHelpID = -1
Private Sub Form_Activate()
    'DB.Enter "frmMain.Form_Activate"
    
    IsActive = True
    If MaxMode Then
        MaxMode = False
        frmFullScreen.Hide
        Set frmFullScreen = Nothing
        buildStatus
    End If
    If Not frmMinor Is Nothing Then
        Unload frmMinor
    End If

    'DB.Leave "frmMain.Form_Activate"
End Sub
Private Sub Form_Deactivate()
    'DB.Enter "frmMain.Form_Deactivate"
    IsActive = False
    'DB.Leave "frmMain.Form_Deactivate"
End Sub
Private Sub Form_Initialize()
    'Warn the user for
    'multiple Node instances
    DB.Enter "frmMain.Form_Initialize"
    If App.PrevInstance Then
        DB.X "Previous Instance of the Program detected"
        'Note: App.PrevInstance will not work inside mdlNode.Main
        '      don't move this code from here if possible.
        DB.X "Checking for potential Reload() message"
        If GetSetting("Node", "Remember", "Reload", False) Then
            DB.X "Reload() message detected, there's no need to worry about Previous Instances"
            SaveSetting "Node", "Remember", "Reload", False
        Else
            DB.X "No Reload() message detected"
            DB.X "Warning User"
            If MsgBox(Language(504), vbYesNo Or vbQuestion, Language(505)) = vbNo Then
                DB.X "User Cancelled Program Execution"
                End
            End If
            DB.X "User chose to keep the program open"
            MultipleInstances = True
        End If
    End If
    DB.X "InitCommonControls"
    On Error GoTo Failed_To_Init_Commons
    InitCommonControls
    DB.Leave "frmMain.Form_Initialize"
    Exit Sub
Failed_To_Init_Commons:
    DB.Leave "frmMain.Form_Initialize", "Failed to InitCommonControls"
    MsgBox "Critical Error in procedure frmMain.Form_Initialize. Error block: 0.", vbInformation
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intFL As Integer
    Dim strHotKeyLine As String
    
    'DB.Enter "frmMain.Form_KeyDown"
    'DB.X "(KeyCode := " & KeyCode & ", Shift := " & Shift & ")"
    'execute hot key
    intFL = FreeFile
    Open App.Path & "/conf/hotkeys.dat" For Input As intFL
    Do Until EOF(intFL)
        Line Input #intFL, strHotKeyLine
        If KeysMatch(GetStatement(strHotKeyLine), KeyCode, Shift) Then
            DB.X "Detected as Hot Key (strHotKeyLine = " & strHotKeyLine & ")"
            ExecuteAction GetParameter(strHotKeyLine)
        End If
    Loop
    Close intFL
    'DB.Leave "frmMain.Form_KeyDown"
End Sub
Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    'first when a DDE link is executed
    'i.e. the user clicks on an irc://
    'link
    
    Dim strServer As String
    Dim strChannel As String
    Dim strCmd As String
    Dim strSessionPerformFilename As String
    Dim intPos As Integer
    Dim intPos2 As Integer
    Dim lPort As Long
    Dim TheServer As clsActiveServer
    Dim intFL As Integer
    
    DB.Enter "LinkExecute"
    DB.X "Received DDE Message: " & CmdStr
    DB.X "Parsing"
    
    'the parameter is passed by reference (without an obvious reason!)
    'but we don't want to modify it.
    'We'll have to use a local variable to
    'store it while we're parsing it
    strCmd = CmdStr
    
    'CmdStr = "irc://irc.freenode.org/node-irc"
    'Get the position of :// here instead of just
    'counting the letters including "irc" so as
    'to make it compatible with other "protocol shortcuts"
    '(which we may implement in the future)
    intPos = InStr(1, strCmd, "://")
    If intPos > 0 Then
        DB.X "Valid DDE Link"
        'the text :// exists in the Link Command String
        'continue
        'if the link is like this irc://irc.myserver.org/
        '(but not like this irc://irc.myserver.org/mychannel/ )
        'remove the end slash
        intPos2 = InStr(intPos + 3, strCmd, "/")
        
        If intPos2 = Len(strCmd) - 1 Then
            strCmd = Left$(strCmd, Len(strCmd) - 1)
            intPos2 = InStr(intPos + 1, strCmd, "/")
        End If
        
        'get the server, it must be between :// and the next /
        If intPos2 <= 0 Then
            'the link is of type irc://irc.myserver.org
            'so there's no channel
            'get the server
            '
            DB.X "Only Server Connection (no channel found)"
            strServer = Right$(strCmd, Len(strCmd) - intPos - 2)
        Else
            DB.X "Server Connection + Channel Join"
            'we have a channel and a server
            'get them
            strServer = Mid$(strCmd, intPos, intPos2 - intPos)
            strChannel = Right$(strCmd, Len(strCmd) - intPos2)
            If Left$(strChannel, 1) = "#" Then
                strChannel = Right$(strChannel, Len(strChannel) - 1)
            End If
        End If
    
        'check if there are existing connected servers;
        'if there aren't, we shall use the current active server
        If Not ConnectionIsPresent Then
            DB.X "No connection is present. Using CAS."
            Set TheServer = CurrentActiveServer
        Else
            DB.X "Connection is present. Using NewServer()."
            'the server will be activated on-create
            Set TheServer = NewServer()
        End If
        
        'if we were asked to join a channel, we shall do this using SessionPerform
        If LenB(strChannel) > 0 Then
            DB.X "We have to join a channel. Building DDE Session Perform."
            intFL = FreeFile
            strSessionPerformFilename = App.Path & "/temp/dde-perform-" & FixLeadingZero(Int(Rnd * 1000), 4) & ".dat"
            DB.X "Writting to file " & strSessionPerformFilename
            Open strSessionPerformFilename For Output Access Write Lock Read Write As #intFL
            Print #intFL, "# This file was automatically created by Node in order to join a channel after a DDE request"
            Print #intFL, "/join #" & strChannel
            Print #intFL, "# End of file"
            Close #intFL
            DB.X "Session Perform File Created"
            DB.X "Setting CAS attributes for perform"
            CurrentActiveServer.DoSessionPerform = True
            CurrentActiveServer.SessionPerformFile = strSessionPerformFilename
            DB.X "Done. DDE Session Perform will be executed uppon RealConnect()."
        End If
    Else
        DB.X "Warning: DDE Link Parsing Error: Cannot detect :// in Command String"
    End If
    'MsgBox "DDE-ing at frmMain:: " & vbnewline & CmdStr, vbInformation, "DDE Debug"
    DB.Leave "LinkExecute"
End Sub
Private Sub Form_Load()
    Dim strRestoreFile As String
    Dim frmSessionsAsk As frmCustom
    Dim ndLastStatus As NodeStatus
    Dim lnDebug As Integer
    Dim i As Integer
       
    DB.Enter "frmMain.Form_Load"
    
    lblPanelTitle.Font.Charset = LangCharSet
    lblPaneTitle.Font.Charset = LangCharSet
    
    On Error GoTo Form_Load_Error
    lnDebug = 2
    'set the title to `Node'
    Me.Caption = Language(0) 'Node
    
    lnDebug = 3
    SplashScreen.lblCustom(1).Caption = Language(19) '"Starting..."
    SplashScreen.lblCustom(1).Refresh
    
    lnDebug = 4
    
    Set tvConnections.ImageList = ilTabs
    
    tvConnections.Nodes.Add(, tvwFirst, "root", Language(793), TabImage_Status).Expanded = True
    
    'create default server and activate it
    DB.X "NewServer"
    
    ReDim ActiveServers(0)
    
    NewServer , True
       
    DB.X "Removing first tab of ntsPanel"
    ntsPanel.DeleteTab
    
    lnDebug = 5
    DB.X "Redim-entioning Arrays"
    'initialize the arrays
    ReDim MessageHistory(0)
    ReDim webdocWebSite(0)
    ReDim NDCConnections(0)
    ReDim NDCConnectionsny(0)
    ReDim AllowedHidden(0)
    ReDim LoadedPanels(0)
    ReDim LangMultiNames(0)
    ReDim AllBuddies(0)
    
    lnDebug = 6
    'initialize variables
    DB.X "Setting Variables to Default Values"
    boolSaveNextType = True
    namesboolean = False
    showwhois = True
    showison = True
    
    lnDebug = 7
    DB.X "NDCGenerateRandomID"
    NDCRandomCurrent = NDCGenerateRandomID()
    DB.X "(NDCRandomCurrent = " & NDCRandomCurrent & ")"
    
    lnDebug = 8
    If Options.StartPage Then
        DB.X "Loading Web Site Tab"
        'the first tab is connected with webbrowser which has index 2
        'its type is WebSite
        CurrentActiveServer.GetStatusID
        CurrentActiveServer.TabInfo.Add WebBrowserIndex_Loading + 1 'first tab
        CurrentActiveServer.TabType.Add TabType_WebSite
        CurrentActiveServer.Tabs.Tabs.Add , , "Web Site"
        CreateWSRoot
        For i = 1 To tvConnections.Nodes.Count
            If tvConnections.Nodes.Item(i).Key = "wbroot" Then
                tvConnections.Nodes.Add tvConnections.Nodes.Item(i), tvwChild, "w_" & GetServerIndexFromActiveServer(CurrentActiveServer) & "_" & WebBrowserIndex_Loading + 1, Language(38), TabImage_WebSite
            End If
        Next i
    End If
    
    lnDebug = 9
        
    lnDebug = 10
    
    lnDebug = 11
    SplashScreen.lblCustom(1).Caption = Language(20) '"Loading Script Engine..."
    SplashScreen.lblCustom(1).Refresh
    'load the script control.
    DB.X "Initializing ndScript ScriptControl"
    Set ndScript = New ScriptControl
    'initialize it to VBS
    ndScript.Language = "VBScript"
    
    lnDebug = 13
    SplashScreen.lblCustom(1).Caption = Language(22) '"Loading Script Routines..."
    SplashScreen.lblCustom(1).Refresh
    If CBool(Len(ThisSkin.CodeBehind)) Then
        DB.X "Executing CodeBehind `Skin_Init'"
        SkinExecuteCodeBehind "Skin_Init", True
    End If
    
    'SplashScreen.lblCustom(1).Caption = Language(23) '"Loading Classes..."
    'SplashScreen.lblCustom(1).Refresh
       
    lnDebug = 14
    SplashScreen.lblCustom(1).Caption = Language(24) '"Loading Menus..."
    SplashScreen.lblCustom(1).Refresh
    
    'load the menus
    lnDebug = 15
    DB.X "LoadMenus"
    LoadMenus
    DB.X "Loading Language Entries for frmMain"
    MainLoadLanguage 'conatins mainly menus language items
    
    lnDebug = 16
    'the default working mode is not full-screen
    MaxMode = False
    'make the controls on the form fit the new window size
    DB.X "Correcting controls placing on form"
    Form_Resize
    SplashScreen.lblCustom(1).Caption = Language(25) '"Loading Background..."
    SplashScreen.lblCustom(1).Refresh
    'load the background of the main form
    'wbBack.Navigate2 App.Path & "\data\skins\" & ThisSkin.BackgroundFile
    DB.X "Loading background, " & App.Path & "\data\html\back.html"
    wbBack.Navigate2 App.Path & "\data\html\back.html"
    
    lnDebug = 161
    SplashScreen.lblCustom(1).Caption = Language(21) '"Loading WMP..."
    SplashScreen.lblCustom(1).Refresh
    ThisSoundSchemePlaySound "start"
    
    lnDebug = 17
    SplashScreen.lblCustom(1).Caption = Language(26) '"Loading HTML Parser..."
    SplashScreen.lblCustom(1).Refresh
    'load a web browser object
    On Error Resume Next
    DB.X "Loading wbStatus( Load )"
    Load wbStatus(WebBrowserIndex_Loading)
    DB.X "Loading wbStatus( Chan )"
    Load wbStatus(WebBrowserIndex_Chan)
    DB.X "Loading wbStatus( DCC )"
    Load wbStatus(WebBrowserIndex_DCC)
    DB.X "Loading wbStatus( Priv )"
    Load wbStatus(WebBrowserIndex_Priv)
    
    lnDebug = 18
    'load method won't make it visible; we'll have to do it
    DB.X "Making wbStatus-es visible"
    wbStatus(WebBrowserIndex_Loading).Visible = True
    wbStatus(WebBrowserIndex_Chan).Visible = True
    wbStatus(WebBrowserIndex_DCC).Visible = True
    wbStatus(WebBrowserIndex_Priv).Visible = True
    
    
    SplashScreen.lblCustom(1).Caption = Language(719) '"Loading WinSock..."
    SplashScreen.lblCustom(1).Refresh
    DB.X "Loading wsIdentD, wsDCC, wsDCCChat, wsDCCSend, wsNDC"
    Set wsIdentD = New clsWSArray
    Set wsDCC = New clsWSArray
    Set wsDCCChat = New clsWSArray
    Set wsDCCSend = New clsWSArray
    Set wsNDC = New clsWSArray
    DB.X "Loading wsIdentD Server"
    wsIdentD.LoadNew
    DB.X "Loading wsNDC Server"
    wsNDC.LoadNew
    
    lnDebug = 19
    'if the user wants a page displayed on startup...
    If Options.StartPage Then
        'load another one; make it the current(selected) web browser
        DB.X "Loading new wbStatus to load Web Site"
        Load wbStatus(xLet(currentWB, WebBrowserIndex_Loading + 1))
        DB.X "Refering DOM Document to webdocWebSite(currentwb := " & currentWB & ")"
        Set webdocWebSite(currentWB) = wbStatus(currentWB).Document
        'also make it visible
        DB.X "Making visible"
        wbStatus(WebBrowserIndex_Loading + 1).Visible = True
    End If
    
    SplashScreen.lblCustom(1).Caption = Language(879) '"Loading Primary HTMLs..."
    SplashScreen.lblCustom(1).Refresh
    lnDebug = 20
    DB.X "BuildPrimary-ing"
    BuildPrimary
    
    SplashScreen.lblCustom(1).Caption = Language(880) '"Loading Secondary HTMLs..."
    SplashScreen.lblCustom(1).Refresh
    DB.X "BuildSecondary-ing"
    BuildSecondary
    
    lnDebug = 21
    SplashScreen.lblCustom(1).Caption = Language(27) '"Preloading Node Web Site..."
    SplashScreen.lblCustom(1).Refresh
    'use the previously created webbrowser object to navigate
    'to the loading web site -- webbrowser with index one is always used to display the "loading" screen
    DB.X "Loading loading.html"
    wbStatus(WebBrowserIndex_Loading).Navigate2 App.Path & "/data/html/imports/loading.html"
    If Options.StartPage Then
        'load start web site
        DB.X "Loading Start Web Site"
        wbStatus(WebBrowserIndex_Loading + 1).Navigate2 Options.StartPageURL
    End If
    
    lnDebug = 22
    SplashScreen.lblCustom(1).Caption = Language(28) '"Refreshing..."
    SplashScreen.lblCustom(1).Refresh
    'resize the background of the window so as it fits the hole window
    DB.X "ZOrder wbBack to the back"
    wbBack.ZOrder 1
        
    lnDebug = 23
    'select the appropritate tab; make the right webbrowser object visible
    DB.X "tsTabs_Click"
    tsTabs_Click 0
    'refresh the form
    DB.X "Form_Resize"
    Form_Resize
    DB.X "Refreshing View"
    Me.Refresh
    'the first tab is a website tab
    tsTabs(0).Tabs.Item(1).Image = TabImage_WebSite
    
    lnDebug = 24
    SplashScreen.lblCustom(1).Caption = Language(29) '"Please wait..."
    SplashScreen.lblCustom(1).Refresh
    'wait 200 milliseconds
    DB.X "DoEvents"
    DoEvents
    'Wait 0.2 'so as WebBrowser object has enough time to initialize
    
    lnDebug = 25
    SplashScreen.lblCustom(1).Caption = vbNullString
    SplashScreen.lblCustom(1).Refresh
    
    'load IRC messages
    DB.X "Loading IRCMsg"
    LoadIRCMsg
    
    lnDebug = 26
    'load status tab
    DB.X "Loading Status Tab: GetStatusID()"
    CurrentActiveServer.GetStatusID
    
    'update the tabsbar: resize the tabs control so as it takes up only the necessary space
    DB.X "UpdateTabsBar"
    UpdateTabsBar
    
    lnDebug = 27
    'show the main window
    DB.X "Showing Main Window"
    Me.Visible = True
    
    lnDebug = 28
    If Options.FadeTransaction Then
        DB.X "Fading In"
        InitWindow Me.hwnd
        SetLayered Me.hwnd, 0&
    End If
    
    lnDebug = 29
    DB.X "Loading TextToolbarPics"
    TextToolbarLoadPics
    ShowMe 'start fade transaction
    
    lnDebug = 30
    DB.X "Loading xpBalloon object: New clsBalloon"
    Set xpBalloon = New clsBalloon
    DB.X "xpBalloon.Init"
    xpBalloon.Init Me.Icon, Me, picTray.hwnd
    If GetSetting(App.EXEName, "InfoTips", App.Major * 100 + App.Minor, "0") = "0" Then
        'show welcome
        DB.X "Showing Welcome Info Tip"
        ShowInfoTip WelcomeToNode
        SaveSetting App.EXEName, "InfoTips", App.Major * 100 + App.Minor, "1"
    End If
    
    lnDebug = 31
    On Error Resume Next 'for any case(the webbrowser object may fail to load)
    'wbBack.Refresh2
    DB.X "wbStatus Refresh2-ing"
    wbStatus(WebBrowserIndex_Loading + 1).Refresh2 'refresh both webbrowser objects
    wbStatus(WebBrowserIndex_Loading).Refresh2
    'add the default text to the Status window
              '"Welcome to Node IRC :)"
    DB.X "AddStatus Welcome Text"
    AddStatus Replace(Replace(Replace(Replace(Language(657), "%1", VERSION_CODENAME & " " & App.Major & "." & App.Minor), _
                "%2", "<a href='NodeScript:/browse sourceforge.net'>SourceForge.net</a>"), "<", HTML_OPEN), ">", HTML_CLOSE) & vbNewLine, CurrentActiveServer
              '"This is Node Version (CodeName) Major.Minor hosted by SourceForge.net"
    
    SplashScreen.lblCustom(1).Caption = Language(383) '"Loading Sessions Info..."
    SplashScreen.lblCustom(1).Refresh
    
    lnDebug = 32
    If mdlNode.bCrash And Not MultipleInstances Then
        DB.X "Handling Crash"
        If Options.SessionC = 0 Then
            DB.X "Resuming Session"
            boolRestoring = True
            strRestoreFile = App.Path & "\temp\crash.xml"
        ElseIf Options.SessionC = 2 Then
            'ask
            DB.X "Asking User"
            Set frmSessionsAsk = New frmCustom
            mdlScripting.xNodeTempValue = Language(277) & vbNewLine & Language(279)
            LoadDialog App.Path & "\data\dialogs\sessions.xml", frmSessionsAsk
            frmSessionsAsk.Hide
            frmSessionsAsk.Show vbModal
            If GetSetting(App.EXEName, "Options", "Temp", "no") = "yes" Then
                'the user agreed to resume
                DB.X "Yes: Resuming"
                boolRestoring = True
                strRestoreFile = App.Path & "\temp\crash.xml"
            End If
        Else
            DB.X "Ignoring Crash"
            If GetSetting(App.EXEName, "InfoTips", "CrashDis", "0") = "0" Then
                DB.X "Ignoring Crash for first time; displaying tip"
                ShowInfoTip CrashDis
                SaveSetting App.EXEName, "InfoTips", "CrashDis", "1"
            End If
        End If
    ElseIf FS.FileExists(App.Path & "\temp\normal.xml") Then
        DB.X "Normal Session Resuming"
        If Options.SessionN = 0 Then
            DB.X "Resuming Session"
            boolRestoring = True
            strRestoreFile = App.Path & "\temp\normal.xml"
        ElseIf Options.SessionN = 2 Then
            'ask
            DB.X "Asking User"
            Set frmSessionsAsk = New frmCustom
            mdlScripting.xNodeTempValue = Language(278) & vbNewLine & Language(279)
            LoadDialog App.Path & "\data\dialogs\sessions.xml", frmSessionsAsk
            frmSessionsAsk.Hide
            frmSessionsAsk.Show vbModal
            If GetSetting(App.EXEName, "Options", "Temp", "no") = "yes" Then
                'the user agreed to resume
                DB.X "Yes: Resuming"
                boolRestoring = True
                strRestoreFile = App.Path & "\temp\normal.xml"
            End If
        End If
    End If
    
    lnDebug = 33
    'TO DO:
    'I think we've already Initialized Sphere Variables
    'Is this really necessary?
    DB.X "InitSphereVariables"
    InitSphereVariables
    
    lnDebug = 34
    If boolRestoring Then
        DB.X "Loading Past Session"
        LoadSession strRestoreFile
    Else
        'check to see if this is the latest version of node.
        If Options.CheckLatest Then
            DB.X "Checking for Latest Version"
            LoadDialog App.Path & "\data\dialogs\latest.xml"
        End If
        'show tip of the day
        If Options.TOD Then
            DB.X "Showing TOD"
            LoadDialog App.Path & "\data\dialogs\tod.xml"
        End If
        'auto-connect to server if enabled
        If Options.StartupConnect Then
            CurrentActiveServer.preExecute "/connect " & Options.StartupConnectHostname & " " & Options.StartupConnectPort & " """ & frmOptions.cboConnectServer.List(frmOptions.cboConnectServer.ListIndex) & """"
        End If
    End If
    lnDebug = 35
    DB.X "Running Script Procedure Begin()"
    RunScript "Begin"
    
    lnDebug = 36
    'init IdentD server
    
    Err.Clear
    
    If GetSetting("Node", "Options", "EnableIdentD", True) Then
        DB.X "wsIdentD(0): Listening"
        wsIdentD.Item(0).LocalPort = 113
        If Not PortIsInUse(113) Then
            wsIdentD.Item(0).Listen
        Else
            DB.XWarning "Port 113 is in use. Could not initialize the IdentD server."
        End If
    Else
        DB.X "IdentD server is disabled."
    End If
    
    If GetSetting("Node", "Options", "NDCServer", True) Then
        'init NDC server
        DB.X "wsNDC(0): Listening"
        wsNDC.Item(0).LocalPort = 8752
        If Not PortIsInUse(8752) Then
            wsNDC.Item(0).Listen
        Else
            DB.XWarning "Port 8752 is in use. Could not initialize the raverix server."
        End If
    Else
        DB.X "Raverix server is disabled."
    End If
    
    lnDebug = 37
    
    'read previous status, and set it
    DB.X "Reading LastStatus"
    ndLastStatus = GetSetting(App.EXEName, "Remember", "Status", 0)
    DB.X "(ndLastStatus = " & ndLastStatus & ")"
    If Options.RestoreStatus Then
        DB.X "Restoring Status"
        If ndLastStatus <> Status_Online Then
            DB.X "Informing User"
            AddStatus SPECIAL_PREFIX & Replace(Language(437), "%1", Language(397 + ndLastStatus)) & SPECIAL_SUFFIX & vbNewLine, CurrentActiveServer
        End If
        nmnuIRC_MenuClick 1, (ndLastStatus)
    Else
        DB.X "User has chosen not to restore status; setting to Online"
        nmnuIRC_MenuClick 1, Status_Online
    End If
    CurrentActiveServer.preExecute "/focus last"
    
    If GetSetting("Node", "Remember", "WelcomeWizard", False) = False Then
        DB.X "First Program Start!! :-)"
        DB.X "Loading Welcome Wizard"
        SaveSetting "Node", "Remember", "WelcomeWizard", True
        LoadWizard "welcome"
    End If
    
    ShowPane
        
    DB.Leave "frmMain.Form_Load"
    Exit Sub
Form_Load_Error:
    DB.XWarning "Critical Error at Error Block " & lnDebug
    DB.Leave "frmMain.Form_Load", "Critical Error; Error Block: " & lnDebug
    MsgBox "Critical Error in procedure frmMain.Form_Load. Error block:" & lnDebug & ".", vbInformation
    CriticalError
End Sub
Private Sub LoadMenus()
    'this sub is used to load the menus of the main window
    'the main menu is horizontal(it contains the menu titles)
    nmnuMain.Horizontal = True
    nmnuMain.Charset = LangCharSet
        'but the file menu is not; it's vertical
        nmnuFile.Horizontal = False
        nmnuFile.Charset = LangCharSet
        'add the file menu items
        nmnuFile.AddMenu Language(3) ', , ilMenu.ListImages(6).Picture 'LoadPicture(App.Path & "/data/graphics/menuicons/disconnect.gif")  '"Disconnect"
        'this menu is a seperator; it cannot be clicked but it counts in the indeses
        nmnuFile.AddMenu "-"
        nmnuFile.AddMenu Language(4) & "..." ', , LoadPicture(App.Path & "/data/graphics/menuicons/options.gif") '"Options..."
        'nmnuFile.AddMenu Language(156) '"Smileys..."
        nmnuFile.AddMenu "-"
        nmnuFile.AddMenu Language(5) ', , LoadPicture(App.Path & "/data/graphics/menuicons/exit.gif") '"Exit"
        nmnuFile.EndMenu
        'bring it to front(so as it's not hidden behind other ActiveXs)
        nmnuFile.ZOrder 0
    'add the file menu to the main menu
    nmnuMain.AddMenu Language(1), nmnuFile '"File"
        'create View menu
        nmnuView.Horizontal = False
        nmnuView.Charset = LangCharSet
        nmnuView.AddMenu Language(2), , ilMenu.ListImages(5).Picture 'connect
        nmnuView.AddMenu Language(323) ' join
        nmnuView.AddMenu Language(130), , ilMenu.ListImages(2).Picture 'buddies
        nmnuView.AddMenu Language(488), , ilMenu.ListImages(4).Picture 'favorites
        nmnuView.AddMenu Language(456) 'avatars
        nmnuView.AddMenu "-" 'seperator
        nmnuView.AddMenu Language(793) 'connections
        nmnuView.EndMenu
        nmnuView.ZOrder 0
    nmnuMain.AddMenu Language(775), nmnuView
        'create Script vertical menu
        nmnuScript.Horizontal = False
        nmnuScript.Charset = LangCharSet
        'add the script menu items
        nmnuScript.AddMenu Language(7), , ilMenu.ListImages(3).Picture '"Main"
            'the script menu contains another submenu: Browse
            nmnuBrowse.Horizontal = False
            nmnuBrowse.Charset = LangCharSet
            nmnuBrowse.AddMenu Language(11) '"Scripts Folder"
            nmnuBrowse.AddMenu Language(12) '"Application Folder"
            nmnuBrowse.AddMenu Language(13) '"Sounds Folder"
            nmnuBrowse.AddMenu Language(127) '"Languages Folder"
            nmnuBrowse.AddMenu Language(847) '"Downloads Folder"
            nmnuBrowse.AddMenu Language(848) '"Logs Folder"
            nmnuBrowse.AddMenu Language(849) '"Dialogs Folder"
            nmnuBrowse.AddMenu Language(850) '"PlugIns Folder"
            nmnuBrowse.EndMenu
            nmnuBrowse.ZOrder 0
        nmnuScript.AddMenu Language(126), nmnuBrowse '"Browse"
        'add the other items
        nmnuScript.AddMenu Language(214) '"Reload"
        nmnuScript.EndMenu
        nmnuScript.ZOrder 0
    'add the script menu to the main menu
    nmnuMain.AddMenu Language(6), nmnuScript '"Script"
        nmnuIRC(0).Horizontal = False
        nmnuIRC(0).Charset = LangCharSet
            Load nmnuIRC(1) 'Status
            nmnuIRC(1).Horizontal = False
            nmnuIRC(1).Charset = LangCharSet
            nmnuIRC(1).AddMenu Language(397) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(398) ', , LoadPicture(App.Path & "/data/graphics/menuicons/away.gif")
            nmnuIRC(1).AddMenu Language(399) ', , LoadPicture(App.Path & "/data/graphics/menuicons/brb.gif")
            nmnuIRC(1).AddMenu Language(400) ', , LoadPicture(App.Path & "/data/graphics/menuicons/sleep.gif")
            nmnuIRC(1).AddMenu Language(401) ', , LoadPicture(App.Path & "/data/graphics/menuicons/busy.gif")
            nmnuIRC(1).AddMenu Language(402) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(403) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(404) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(405) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(406) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).AddMenu Language(407) ', , LoadPicture(App.Path & "/data/graphics/menuicons/online.gif")
            nmnuIRC(1).EndMenu
            nmnuIRC(1).ZOrder 0
        nmnuIRC(0).AddMenu Language(408), nmnuIRC(1)
        nmnuIRC(0).AddMenu "-"
            Load nmnuIRC(2) 'Nicknames
            nmnuIRC(2).Horizontal = False
            nmnuIRC(2).Charset = LangCharSet
            nmnuIRC(2).AddMenu Language(329) & "..." 'identify
            nmnuIRC(2).AddMenu Language(497) & "..." 'register
            nmnuIRC(2).AddMenu Language(797) & "..." 'ghost
            nmnuIRC(2).AddMenu Language(838) & "..." 'release
            nmnuIRC(2).AddMenu Language(839) & "..." 'drop
            nmnuIRC(2).EndMenu
            nmnuIRC(2).ZOrder 0
        nmnuIRC(0).AddMenu Language(321), nmnuIRC(2)
            Load nmnuIRC(3) 'Memos
            nmnuIRC(3).Horizontal = False
            nmnuIRC(3).Charset = LangCharSet
            nmnuIRC(3).AddMenu Language(351), , ilMenu.ListImages(10).Picture  'list
            nmnuIRC(3).AddMenu Language(422) & "...", , ilMenu.ListImages(8).Picture 'send
            nmnuIRC(3).AddMenu Language(352) & "...", , ilMenu.ListImages(9).Picture 'read
            nmnuIRC(3).AddMenu Language(353) & "...", , ilMenu.ListImages(11).Picture 'delete
            nmnuIRC(3).EndMenu
            nmnuIRC(3).ZOrder 0
        nmnuIRC(0).AddMenu Language(322), nmnuIRC(3) 'memos
            Load nmnuIRC(5) 'Channels
            nmnuIRC(5).Horizontal = False
            nmnuIRC(5).Charset = LangCharSet
            nmnuIRC(5).AddMenu Language(497) & "..."  'register
            nmnuIRC(5).AddMenu Language(329) & "..." 'identify
            nmnuIRC(5).AddMenu Language(839) & "..." 'drop
                Load nmnuIRC(4) 'Channel Access
                nmnuIRC(4).Horizontal = False
                nmnuIRC(4).Charset = LangCharSet
                nmnuIRC(4).AddMenu Language(84) 'add
                nmnuIRC(4).AddMenu Language(197) & "..." 'delete
                nmnuIRC(4).AddMenu Language(775) & "..." 'view
                nmnuIRC(4).EndMenu
                nmnuIRC(4).ZOrder 0
            nmnuIRC(5).AddMenu Language(799), nmnuIRC(4) 'channel access
            nmnuIRC(5).EndMenu
            nmnuIRC(5).ZOrder 0
        nmnuIRC(0).AddMenu Language(320), nmnuIRC(5) 'channels
        nmnuIRC(0).EndMenu
        nmnuIRC(0).ZOrder 0
    nmnuMain.AddMenu Language(319), nmnuIRC(0) 'IRC
        nmnuHelp.Horizontal = False
        nmnuHelp.Charset = LangCharSet
        nmnuHelp.AddMenu Language(137), , ilMenu.ListImages.Item(1).Picture '"Help Topics"
        nmnuHelp.AddMenu "-"
                           '"View Changes Log"
        nmnuHelp.AddMenu Language(14) & " " & VERSION_CODENAME & " " & App.Major & "." & App.Minor, , ilMenu.ListImages(7).Picture
        nmnuHelp.AddMenu "-"
        nmnuHelp.AddMenu Language(520), , ilMenu.ListImages(6).Picture 'Contact Node Dev Team
        nmnuHelp.AddMenu Language(138) & " " & Language(0) '"About Node"
        nmnuHelp.EndMenu
        nmnuHelp.ZOrder 0
    nmnuMain.AddMenu Language(136), nmnuHelp '"Help"
    nmnuMain.EndMenu
    
    'finally, bring the Main menu in front of other controls
    nmnuMain.ZOrder 0
End Sub
Private Sub MenusToFront()
    'won't be used...
    nmnuFile.ZOrder 0
    nmnuBrowse.ZOrder 0
    nmnuScript.ZOrder 0
    nmnuHelp.ZOrder 0
    nmnuMain.ZOrder 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = vbFormControlMenu Then
        Ending = True
    End If
End Sub
Public Sub Form_Resize()
    'When the form is being resized
    Dim i As Integer
    Dim intTemp As Integer
    'change form's objects' sizes
    'so as they fit the window's size
    'we don't have to do that if the window is minimized
       
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    If wbStatus(currentWB).Top <> ThisSkin.Resize_TopOffset Then
        wbStatus(currentWB).Top = ThisSkin.Resize_TopOffset
    End If
    If wbStatus(currentWB).Left <> ThisSkin.Resize_LeftOffset + IIf(IsPaneOpen, fraPane.Width, 0) Then
        wbStatus(currentWB).Left = ThisSkin.Resize_LeftOffset + IIf(IsPaneOpen, fraPane.Width, 0)
    End If
    'if the Tabs left is incorrect...
    If CurrentActiveServer.Tabs.Left <> wbStatus(currentWB).Left Then
        'set it to be the same with the currently visible webbrowser
        CurrentActiveServer.Tabs.Left = wbStatus(currentWB).Left
        'if the left is incorrect its height may also be wrong; correct it as well
        'se it to be the same as the current webbrowser's top minus the tabs height
        CurrentActiveServer.Tabs.Top = wbStatus(currentWB).Top - CurrentActiveServer.Tabs.Height
        'also set the textbox left to its default value
        txtSend.Left = ThisSkin.Resize_TextLeftOffset + IIf(IsPaneOpen, fraPane.Width, 0)
        
        'this block of code should only be executed once...
    End If
    'in the case focusing is impossible add an error trap
    On Error Resume Next
    'focus the textbox
    If txtSend.Visible Then
        txtSend.SetFocus
    End If
    'if the webbrowser width is incorrect...
    If wbStatus(currentWB).Width <> Me.Width - ThisSkin.Resize_RightOffset - IIf(fraPanel.Visible, fraPanel.Width, 0) - IIf(IsPaneOpen, fraPane.Width, 0) Then
        'fix it...
        wbStatus(currentWB).Width = Me.Width - ThisSkin.Resize_RightOffset - IIf(fraPanel.Visible, fraPanel.Width, 0) - IIf(IsPaneOpen, fraPane.Width, 0)
        fraWebTab.Width = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Width
        'wbStatus(TabInfo(tsTabs.SelectedItem.index)).Width = Me.Width - 340
        'the main menu width may also be wrong... correct it as well
        nmnuMain.Width = Me.Width
        'the textbox width depends on the webbrowser width...
        txtSend.Width = Me.Width - ThisSkin.Resize_TextRightOffset - fraMore.Width - 40 - IIf(fraPanel.Visible, fraPanel.Width + 200, 0) - IIf(IsPaneOpen, fraPane.Width, 0)
        fraMore.Left = txtSend.Left + txtSend.Width
        tbText.Left = txtSend.Left + txtSend.Width + fraMore.Width - tbText.Width
    End If
    'do the same with the webbrowser's height so it fits the hole window
    If wbStatus(currentWB).Height <> Me.Height - sbar.Height - IIf(txtSend.Visible, ThisSkin.Resize_BottomOffset + 750, ThisSkin.Resize_BottomOffset + 400 + ThisSkin.Resize_BottomHiddenOffset) Then
        wbStatus(currentWB).Height = Me.Height - sbar.Height - IIf(txtSend.Visible, ThisSkin.Resize_BottomOffset + 750, ThisSkin.Resize_BottomOffset + 400 + ThisSkin.Resize_BottomHiddenOffset)
    End If
    If txtSend.Top <> Me.Height - sbar.Height - ThisSkin.Resize_BottomOffset - IIf(MultilineText, 2500, 0) Then
        txtSend.Top = Me.Height - sbar.Height - ThisSkin.Resize_BottomOffset - ThisSkin.Resize_TextBottomOffset - IIf(MultilineText, 2500, 0)
        If MultilineText Then
            txtSend.Height = 2850
        Else
            txtSend.Height = 350
        End If
        fraMore.Top = txtSend.Top
        tbText.Top = txtSend.Top - tbText.Height
    End If
    If wbBack.Width <> Me.Width + 270 - IIf(fraPanel.Visible, fraPanel.Width, 0) - IIf(IsPaneOpen, fraPane.Width, 0) Then
        wbBack.Width = Me.Width + 270 - IIf(fraPanel.Visible, fraPanel.Width, 0) - IIf(IsPaneOpen, fraPane.Width, 0)
        wbBack.Left = IIf(IsPaneOpen, fraPane.Width - 30, -30)
    End If
    If wbBack.Height <> Me.Height - sbar.Height + 250 Then
        wbBack.Height = Me.Height - sbar.Height + 250
    End If
    intTemp = Me.Width - 5000
    If intTemp < 5000 Then
        intTemp = 5000
    End If
    If sbar.Panels(1).Width <> intTemp Then
        sbar.Panels(1).Width = intTemp
    End If
    If sbar.Panels(2).Width <> 5000 Then
        sbar.Panels(2).Width = 5000
    End If
    If lstSuggestions.Top <> txtSend.Top - lstSuggestions.Height Then
        lstSuggestions.Top = txtSend.Top - lstSuggestions.Height
    End If
    If lstSuggestions.Left <> txtSend.Left Then
        lstSuggestions.Left = txtSend.Left
    End If
    If fraPanel.Height <> Me.ScaleHeight - sbar.Height - Panel_Title_Height - 60 Then
        fraPanel.Top = nmnuMain.Height
        fraPanel.Height = Me.ScaleHeight - sbar.Height - Panel_Title_Height - 60
        wbPanel.Height = fraPanel.Height - wbPanel.Top + 360
        picPanelResize.Height = fraPanel.Height - picPanelResize.Top
        imgClosePanel.Top = Panel_Title_Height \ 2 - imgClosePanel.Height \ 2
        ntsPanel.Top = fraPanel.Height - ntsPanel.Height
    End If
    If fraPanel.Left <> Me.ScaleWidth - fraPanel.Width Then
        fraPanel.Left = Me.ScaleWidth - fraPanel.Width
        imgClosePanel.Left = fraPanel.Width - imgClosePanel.Width
        imgPanelBegin.Left = 0
        imgPanel.Left = imgPanelBegin.Left + imgPanelBegin.Width
        imgPanelEnd.Left = fraPanel.Width - imgPanelEnd.Width
        imgPanel.Width = fraPanel.Width - imgPanelEnd.Width - imgPanelBegin.Width
        wbPanel.Width = fraPanel.Width + 360
        fraPanelTitle.Width = wbPanel.Width
        picPanelResize.Left = 0
        ntsPanel.Left = 60
        ntsPanel.Width = fraPanel.Width
    End If
    If fraPane.Height <> Me.ScaleHeight Or imgClosePane.Left <> fraPane.Width - imgClosePane.Width Then
        fraPane.Height = Me.ScaleHeight - sbar.Height - fraPanel.Top
        tvConnections.Height = fraPane.Height - tvConnections.Top
        imgPane.Width = fraPane.Width - imgPaneBegin.Width - imgPaneBegin.Left - imgPaneEnd.Width
        imgPaneEnd.Left = imgPane.Left + imgPane.Width
        imgClosePane.Left = fraPane.Width - imgClosePane.Width
        picPaneResize.Left = fraPane.Width - picPaneResize.Width
        picPaneResize.Height = fraPane.Height
        tvConnections.Width = fraPane.Width
    End If
    UpdateTabsBar
    SkinExecuteCodeBehind "Skin_Resize"
End Sub
Private Sub imgClosePane_Click()
    ShowPane
End Sub
Public Sub imgMore_Click()
    If imgMore.Tag = "6" Then
        imgMore.Tag = "5"
        tbText.Visible = False
    Else
        imgMore.Tag = "6"
        tbText.Visible = True
    End If
End Sub
Private Sub lstSuggestions_Click()
    Dim intCurrentPosition As Integer
    Dim intInitLen As Integer
    Dim intRemovedLen As Integer
    Dim strCompletedText As String
    Dim strRemovedPhrase As String
    
    If BolLstSuggstionsCodeClick Then
        Exit Sub
    End If
    
    intCurrentPosition = txtSend.SelStart
    'if the cursor is at the beginning of the textbox we don't have anything to complete
    strCompletedText = lstSuggestions.List(lstSuggestions.ListIndex)
    GiveFocus = True
    If intCurrentPosition = 0 Then
        'skip textcompleter
        txtSend.Text = strCompletedText & txtSend.Text
        txtSend.SelStart = intCurrentPosition + Len(strCompletedText)
        Exit Sub
    End If
    strRemovedPhrase = RemoveWord(txtSend.Text, intCurrentPosition)
    intInitLen = Len(txtSend.Text)
    intRemovedLen = Len(strRemovedPhrase)
    txtSend.Text = AddPhrase(strRemovedPhrase, strCompletedText, intCurrentPosition - (intInitLen - intRemovedLen))
    txtSend.SelStart = intCurrentPosition + Len(strCompletedText) - 1
    lstSuggestions.Visible = False
End Sub
Private Sub lstSuggestions_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        lstSuggestions_Click
    End If
End Sub
Private Sub mnuChanModes_Click()
    CurrentActiveServer.preExecute "/chanmodes", False
End Sub
Private Sub mnuChanProperties_Click()
    CurrentActiveServer.preExecute "/chanproperties", False
End Sub
Private Sub mnuChanTabLeave_Click()
    mnuCloseTab_Click
End Sub
Private Sub mnuChanTabRejoin_Click()
    CurrentActiveServer.preExecute "/hop"
End Sub
Private Sub mnuClear_Click()
    CurrentActiveServer.preExecute "/clear", False
End Sub
Private Sub mnuGiveHalfOp_Click()
    If CurrentNick = CurrentActiveServer.myNick Then
        'the user clicked on his/her own nickname
        'use chanserv to get operator status
        CurrentActiveServer.preExecute "/chanserv halfop " & CurrentActiveServer.Tabs.SelectedItem.Caption & " " & CurrentNick
    Else
        'the user clicked on a nick on the nicklist and chose to give him/her operator status
        'execute the command /mode #channel +o nickname
        '(when the user clicked on that nickname it was stored in the variable CurrentNick)
        CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " +h " & CurrentNick
    End If
End Sub
Private Sub mnuNewServer_Click()
    NewServer
    nmnuView_MenuClick (0)
End Sub
Private Sub mnuNickClear_Click()
    mnuClear_Click
End Sub
Private Sub mnuNickIgnore_Click()
    Dim i As Integer
    For i = 0 To frmOptions.lstIgnore(0).ListCount - 1
        If frmOptions.lstIgnore(0).List(i) = CurrentNick Then
            frmOptions.lstIgnore(0).RemoveItem i
            Exit Sub
        End If
    Next i
    
    frmOptions.lstIgnore(0).AddItem CurrentNick
End Sub
Private Sub mnuNickViewLogs_Click()
    mnuViewLogs_Click
End Sub
Private Sub mnuRemoveServer_Click()
    tvConnections_NodeClick tvConnections.SelectedItem
    DeleteServer CurrentActiveServer
End Sub
Private Sub mnuTakeHalfOp_Click()
    'remove operator status from a user
    'execute /mode #channel -o nickname
    CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " -h " & CurrentNick
End Sub
Private Sub mnuUserHost_Click()
    CurrentActiveServer.preExecute "/userhost " & CurrentNick
End Sub
Private Sub mnuViewLogs_Click()
    Dim strNet As String
    Dim strFile As String
    
    If Options.LogByNetwork Then
        If LenB(CurrentActiveServer.WinSockConnection.RemoteHost) = 0 Then
            strNet = vbNullString
        Else
            strNet = "\" & CurrentActiveServer.WinSockConnection.RemoteHost
        End If
    Else
        strNet = vbNullString
    End If
    strFile = CurrentActiveServer.Tabs.SelectedItem.Caption
    
    'some characters that channels contain
    'cannot be used as filenames
    'and they must be replaced before
    'we can save anything
    strFile = Replace(strFile, "/", "_")
    strFile = Replace(strFile, "\", "_")
    strFile = Replace(strFile, ":", "_")
    strFile = Replace(strFile, "*", "_")
    strFile = Replace(strFile, "|", "_")
    strFile = Replace(strFile, "<", "_")
    strFile = Replace(strFile, ">", "_")
    strFile = Replace(strFile, """", "_")
    
    'open the appropriate log file
    xShell App.Path & "\logs" & strNet & "\" & strFile & ".log.html """"", 0
End Sub

'Private Sub mnuClosePanelTab_Click()
'    ntsPanel.DeleteTab ntsPanel.SelectedTab
'End Sub
Private Sub mnuWebTabFav_Click()
    Dim thiswb As WebBrowser
    Dim intFL As Integer
    Dim strTemp As String
    Set thiswb = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index))
    If Len(CurrentActiveServer.Tabs.Tabs(CurrentActiveServer.Tabs.SelectedItem.index).Caption) > 25 Then
        strTemp = Strings.Left$(CurrentActiveServer.Tabs.Tabs(CurrentActiveServer.Tabs.SelectedItem.index).Caption, 25) & "..."
    Else
        strTemp = CurrentActiveServer.Tabs.Tabs(CurrentActiveServer.Tabs.SelectedItem.index).Caption
    End If
    intFL = FreeFile
    Open App.Path & "/conf/favwebs.lst" For Append As #intFL
        Print #intFL, strTemp
        Print #intFL, thiswb.LocationURL
    Close #intFL
    nmnuView_MenuClick 3
End Sub
Private Sub nmnuView_MenuClick(SubMenuIndex As Integer)
    Dim intFL As Integer
    
    Select Case SubMenuIndex
        Case 0
            'connect
            LoadPanel "connect"
        Case 1
            'join
            LoadPanel "join"
        Case 2
            'buddy list
            LoadPanel "buddylist"
        Case 3
            'favs
            LoadPanel "favorites"
        Case 4
            'avatars
            LoadPanel "avatar"
        Case 6
            ShowPane
    End Select
End Sub
Private Sub ntsPanel_TabPressed(PreviousTab As Integer)
    LoadPanel LoadedPanels(ntsPanel.SelectedTab)
End Sub

'RELNOTE:
'Index parameter in TabDblClick event causes errors
'Should be fixed before releasing

'Private Sub ntsPanel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RELNOTE:
'    MouseDown event doesn't exist yet.
'    We need to add it before releasing.
'    RELNOTE:
'    HitTest doesn't exist yet.
'    Create it.
'    ntsPanel.SelectedTab = ntsPanel.HitTest(X, Y)
'    If Button = 2 Then
'        PopupMenu mnuPanelTab, , X, Y, mnuClosePanelTab
'    End If
'End Sub

Private Sub picPanelResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
'    MinimumPanelWidth = 0
'    For i = 1 To ntsPanel.NumberOfTabs
'        MinimumPanelWidth = MinimumPanelWidth + ScaleX(ntsPanel.TabWidth(i), vbPixels, vbTwips)
'    Next i
    'MinimumPanelWidth = ScaleX(ntsPanel.TabWidth(ntsPanel.NumberOfTabs), vbPixels, vbTwips)

    If Button = 1 Then
        intPanelResizeStartX = X
        boolPanelResizing = True
    End If
End Sub
Private Sub picPanelResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CurrentPOINT As POINTAPI
    Dim intNewWidth As Integer
    Static intPrvWidth As Integer
    
    GetCursorPos CurrentPOINT
    If Button = 1 And boolPanelResizing Then
        intNewWidth = Me.ScaleWidth - ScaleX(CurrentPOINT.X, vbPixels, vbTwips) + intPanelResizeStartX + Me.Left
        If intNewWidth > MinimumPanelWidth Then
            On Error Resume Next
            fraPanel.Width = intNewWidth
        End If
        If intNewWidth <> intPrvWidth Then
            Form_Resize
        End If
        intPrvWidth = intNewWidth
    End If
End Sub
Private Sub picPanelResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        boolPanelResizing = False
        SaveSetting "Node", "Remember", "Panel_" & strCurrentPanel, fraPanel.Width
    End If
End Sub
Private Sub picPaneResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        boolPaneResizing = False
        SaveSetting "Node", "Remember", "Pane", fraPane.Width
    End If
End Sub
Private Sub picPaneResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
'    MinimumPanelWidth = 0
'    For i = 1 To ntsPanel.NumberOfTabs
'        MinimumPanelWidth = MinimumPanelWidth + ScaleX(ntsPanel.TabWidth(i), vbPixels, vbTwips)
'    Next i
    'MinimumPanelWidth = ScaleX(ntsPanel.TabWidth(ntsPanel.NumberOfTabs), vbPixels, vbTwips)

    If Button = 1 Then
        intPaneResizeStartX = X
        boolPaneResizing = True
    End If
End Sub
Private Sub picPaneResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim CurrentPOINT As POINTAPI
    Dim intNewWidth As Integer
    
    GetCursorPos CurrentPOINT
    If Button = 1 And boolPaneResizing Then
        intNewWidth = ScaleX(CurrentPOINT.X, vbPixels, vbTwips) - Me.Left
        If intNewWidth > MinimumPanelWidth Then
            On Error Resume Next
            fraPane.Width = intNewWidth
        End If
        'LockWindowUpdate Me.hwnd
        Form_Resize
        'LockWindowUpdate 0
    End If
End Sub
Private Sub imgClosePanel_Click()
    Dim i As Integer
    
    fraPanel.Visible = False
    picPanelResize.Visible = False
    strCurrentPanel = vbNullString
    'unload all panels
    ReDim LoadedPanels(0)
    For i = 1 To ntsPanel.NumberOfTabs
        ntsPanel.DeleteTab
    Next i
    Form_Resize
End Sub
Private Sub mnuDCCChat_Click()
    CurrentActiveServer.RequestDCCChat CurrentNick
End Sub
Private Sub mnuDCCSend_Click()
    'initiate the DCC Send request
    If LenB(strDCCPresetFile) = 0 Then
        On Error GoTo Canceled
        cdfile.ShowOpen
    Else
        cdfile.FileName = strDCCPresetFile
        strDCCPresetFile = vbNullString
    End If
    'check the file that we want to send
    'make sure it exists and if it is a .lnk file
    'ask if the user really wants to send that file
    If Not FS.FileExists(cdfile.FileName) Then
        'The file does not exist
        MsgBox Replace(Language(263), "%1", cdfile.FileName), vbInformation
        Exit Sub
    ElseIf Strings.Right$(Strings.LCase$(cdfile.FileName), 4) = ".lnk" Then
        'The file you are trying to send is a shortcut...
        If MsgBox(Language(264), vbYesNo Or vbQuestion, Language(265)) = vbNo Then
            Exit Sub
        End If
    ElseIf LCase$(Replace(cdfile.FileName, "/", "\")) = LCase$(Replace(App.Path & "\logs\raw.dat", "/", "\")) Then
        'The file you are trying to send is an active log...
        If MsgBox(Language(710), vbYesNo Or vbQuestion, Language(265)) = vbNo Then
            Exit Sub
        End If
    End If
    CurrentActiveServer.DCCTransfer_SendFile CurrentNick, cdfile.FileName, False
Canceled:
End Sub
Private Sub mnuAddBuddy_Click()
    Dim buddynum As Integer
    Dim intFL As Integer
    Dim buddytext As String
    Dim i As Integer
    Dim nameflg As Integer
    nameflg = 0 'set this "flag" to 0
    'check to see if the nick is already in the list
    For i = 0 To frmOptions.lstBdyNk.ListCount - 1
        If CurrentNick = frmOptions.lstBdyNk.List(i) Then
            If mnuAddBuddy.Caption <> Language(796) Then
                MsgBox Language(157) 'You have already added this name.
            End If
            nameflg = 1 'set the "flag" to 1 - already added
        End If
    Next i
    'if you haven't already added this nick then add it
    If nameflg = 0 Then
        'popup for entering welcome text
        buddytext = InputBox(Language(143), Language(84))
        'add the currently selected nick to the buddy list
        frmOptions.lstBdyNk.AddItem CurrentNick
        'add the entered welcome text to the list
        frmOptions.lstBdyWT.AddItem buddytext
        intFL = FreeFile 'get a free file number
        'open buddy.info where all the buddy list info is stored
        Open App.Path & "\conf\buddy.info" For Output As #intFL
            'if the buddy list has any nicks in it, then save them to the file
            If frmOptions.lstBdyNk.ListCount > 0 Then
                'go through the buddy list
                For buddynum = 0 To frmOptions.lstBdyNk.ListCount - 1
                    'save the nick list
                    Print #intFL, frmOptions.lstBdyNk.List(buddynum)
                    'save the welcome text list
                    Print #intFL, frmOptions.lstBdyWT.List(buddynum)
                Next buddynum
            End If
        Close #intFL 'close buddy.info
    Else
        For i = 0 To frmOptions.lstBdyNk.ListCount - 1
            If frmOptions.lstBdyNk.List(i) = CurrentNick Then
                frmOptions.lstBdyNk.RemoveItem i
                frmOptions.lstBdyWT.RemoveItem i
                Exit For
            End If
        Next i
    End If
End Sub
Private Sub mnuCloseTab_Click()
    'the user right-clicked on a tab and from the pop up menu he/she chose to close it
    'execute the command /close
    CurrentActiveServer.preExecute "/close"
End Sub
Private Sub mnuEnd_Click()
    Ending = True
    Unload Me
End Sub
Private Sub mnuGiveOp_Click()
    If CurrentNick = CurrentActiveServer.myNick Then
        'the user clicked on his/her own nickname
        'use chanserv to get operator status
        CurrentActiveServer.preExecute "/chanserv op " & CurrentActiveServer.Tabs.SelectedItem.Caption & " " & CurrentNick
    Else
        'the user clicked on a nick on the nicklist and chose to give him/her operator status
        'execute the command /mode #channel +o nickname
        '(when the user clicked on that nickname it was stored in the variable CurrentNick)
        CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " +o " & CurrentNick
    End If
End Sub
Private Sub mnuGiveVoice_Click()
    'the user wants to give voice to a buddy
    'execute /mode #channel +v nickname
    'using CurrentNick variable again
    CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " +v " & CurrentNick
End Sub
Private Sub mnuKick_Click()
    'this time the user chose to kick someone
    'load the kick dialog and show it
    LoadDialog App.Path & "\data\dialogs\kick.xml"
End Sub
'Private Sub mnuNDCAudio_Click()
'    Dim intNDCIndex As Integer
'    'get index of the NDC connection
'    intNDCIndex = GetNDCFromNickname(CurrentNick)
'    'if there's no NDC connection
'    If intNDCIndex = -1 Then
'        'display warning
'        AddStatus EVENT_PREFIX & Language(426) & EVENT_SUFFIX & vbnewline, GetChanID(tsTabs.SelectedItem.Caption)
'        'and don't continue
'        Exit Sub
'    End If
'    'if there is an NDC connection
'    'request Audio Conversation
'    wsNDC.Item(intNDCIndex).SendData "a"
'    NDCConnections(intNDCIndex).AudioRequested = True
'End Sub
Private Sub mnuNDCConnect_Click()
    Dim intTemp As Integer
    'check to see if a connection is already present.
    intTemp = GetNDCFromNickname(CurrentNick)
    If intTemp <> -1 Then
        If wsNDC.Item(intTemp).State <> sckConnected Then
            intTemp = -1
        End If
    End If
    'only if an NDC connection isn't already present...
    If intTemp = -1 Then
        'Send NDC* CTCP so that the remote client
        'sends back NDC if it is Node. In case it
        'is another client it should simply show
        ' "Please disregard this message"
        CurrentActiveServer.preExecute "/PRIVMSG " & CurrentNick & " :" & Strings.ChrW$(1) & "NDC* Please disregard this message" & Strings.ChrW$(1)
    End If
End Sub
Private Sub mnuNDCStartNetMeeting_Click()
    Dim intTemp As Integer
    Dim TheServer As clsActiveServer
    
    intTemp = GetNDCFromNickname(CurrentNick)
    If intTemp = -1 Then
        'TO DO:
        'Disable this menu item if no
        'NDC connection is present!
        MsgBox Language(785), vbExclamation
    Else
        '                                        01 09 = NetMeeting
        Set TheServer = NDCConnections(intTemp).ActiveServer
        If wsNDC.Item(intTemp).State = sckConnected Then
            wsNDC.Item(intTemp).SendData "M?" & ChrW$(1) & ChrW$(9)
        Else
            MsgBox Language(785), vbExclamation
            NDCConnections(intTemp).strNicknameA = vbNullString
            Exit Sub
        End If
        AddStatus Replace(Replace(Language(789), "%1", CurrentNick), "%2", Language(783)) & vbNewLine, TheServer, TheServer.GetChanID(CurrentNick)
        NDCConnections(intTemp).MMNetMeetingStatus = 1
    End If
End Sub
Private Sub mnuShow_Click()
    Me.Show
End Sub
Private Sub mnuTabConnect_Click()
    LoadPanel "connect"
End Sub
Private Sub mnuTabDisconnect_Click()
    nmnuFile_MenuClick 0 'click disconnect from the file menu
End Sub
Private Sub mnuTakeOp_Click()
    'remove operator status from a user
    'execute /mode #channel -o nickname
    CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " -o " & CurrentNick
End Sub
Private Sub mnuTakeVoice_Click()
    'and remove voice: /mode #channel -v nickname
    CurrentActiveServer.preExecute "/mode " & CurrentActiveServer.Tabs.SelectedItem.Caption & " -v " & CurrentNick
End Sub
Private Sub mnuWebTabBack_Click()
    On Error Resume Next
    wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).GoBack
End Sub
Private Sub mnuWebTabForward_Click()
    On Error Resume Next
    wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).GoForward
End Sub
Private Sub mnuWebTabRefresh_Click()
    On Error Resume Next
    wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Refresh2
End Sub
Private Sub mnuWebTabStop_Click()
    On Error Resume Next
    wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Stop
End Sub
Private Sub mnuWhisper_Click()
    'we want to talk privately to someone
    'use the command /query nickname
    'to do this.
    CurrentActiveServer.preExecute "/query " & CurrentNick
End Sub
Private Sub mnuInfo_Click()
    'This procedure will show info for the
    'specified nickname by using the /nickserv info command
    CurrentActiveServer.preExecute "/nickserv info " & CurrentNick
End Sub
Private Sub mnuWhoIs_Click()
    'Execeute /whois nickname
    'command in order to get information
    'about the clicked nick, as the user asked.
    CurrentActiveServer.preExecute "/whois " & CurrentNick
End Sub
Private Sub mnuCtcpVer_Click()
    CurrentActiveServer.preExecute "/ctcp " & CurrentNick & " VERSION"
End Sub
Private Sub mnuCtcpTime_Click()
    CurrentActiveServer.preExecute "/ctcp " & CurrentNick & " TIME"
End Sub
Private Sub mnuCTCPPing_Click()
    CurrentActiveServer.preExecute "/ctcp " & CurrentNick & " PING"
End Sub
Private Sub nmnuBrowse_MenuClick(SubMenuIndex As Integer)
    'an item inside Scripts\Browse menu has been clicked
    Select Case SubMenuIndex
        Case 0 'scripts folder
            'use Windows Explorer to explore this folder
            Shell "explorer """ & App.Path & "\scripts""", vbMaximizedFocus
        Case 1 'application folder
            Shell "explorer """ & App.Path & """", vbMaximizedFocus
        Case 2 'sounds folder
            Shell "explorer """ & App.Path & "\data\sounds""", vbMaximizedFocus
        Case 3 'languages folder
            Shell "explorer """ & App.Path & "\data\languages""", vbMaximizedFocus
        Case 4 'downloads folder
            Shell "explorer """ & App.Path & "\downloads""", vbMaximizedFocus
        Case 5 'logs folder
            Shell "explorer """ & App.Path & "\logs""", vbMaximizedFocus
        Case 6 'dialogs folder
            Shell "explorer """ & App.Path & "\data\dialogs""", vbMaximizedFocus
        Case 7 'plugins folder
            Shell "explorer """ & App.Path & "\data\plugins""", vbMaximizedFocus
    End Select
End Sub
Private Sub nmnuBrowse_PopupMove()
    'this menu should also be moved to the correct position using the FixPosition method
    nmnuBrowse.FixPosition nmnuBrowse, nmnuScript
End Sub
Public Sub nmnuFile_MenuClick(SubMenuIndex As Integer)
    Dim intFL As Integer
    Dim intLnCount As Integer
    Dim a As Long
    Dim strQuit As String
    
    'the user selected an item from the file menu
    Select Case SubMenuIndex
        Case 0 'disconnect
            If CurrentActiveServer.WinSockConnection.State = sckConnected Then
                If Options.QuitMultiple Then
                    If FS.FileExists(Options.QuitFile) Then
                        intFL = FreeFile
                        Open Options.QuitFile For Input As #intFL
                        Do Until EOF(intFL)
                            xLineInput intFL
                            intLnCount = intLnCount + 1
                        Loop
                        Close #intFL
                        
                        a = Rnd() * intLnCount
                        
                        intLnCount = 0
                        intFL = FreeFile
                        Open Options.QuitFile For Input As #intFL
                        Do Until EOF(intFL)
                            If intLnCount = a Then
                                Line Input #intFL, strQuit
                                Exit Do
                            Else
                                xLineInput intFL
                            End If
                            intLnCount = intLnCount + 1
                        Loop
                        Close #intFL
                    Else
                        strQuit = Options.QuitMsg
                        DB.XWarning "Quit List File does not exist!"
                    End If
                Else
                    strQuit = Options.QuitMsg
                End If
                On Error Resume Next
                CurrentActiveServer.SendData "QUIT :" & strQuit & vbNewLine
                
                'close the connection if we are not already disconnected
                '(see event wsIRC_Close for more information)
                'wsIRC_Close
            Else
                CurrentActiveServer.WinSockConnection_Close
            End If
        Case 2 'options
            'show the options dialog
            frmOptions.Show
        Case 4 'exit
            'quiting... unload the current form
            '(see event Form_Unload for actions after this)
            Ending = True
            Unload Me
    End Select
End Sub
Public Sub nmnuHelp_MenuClick(SubMenuIndex As Integer)
    Select Case SubMenuIndex
        Case 0 'Help topics
            CurrentActiveServer.preExecute "/browse node.sourceforge.net/?view=home&head=doc"
        Case 2 'view changes logs
            'use EditScript method to display this textfile
            Shell "notepad.exe """ & App.Path & "\doc\changelog.txt""", vbMaximizedFocus
        Case 4 'Contact Node Dev Team
            CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=contact"
        Case 5 'About Node
            LoadDialog App.Path & "/data/dialogs/info.xml"
    End Select
End Sub
Public Sub nmnuIRC_MenuClick(index As Integer, SubMenuIndex As Integer)
    Dim i As Integer
    Dim strInfo As String
    Dim strInfo2 As String
    Dim strInfo3 As String
    
    'IRC >
    Select Case index
        Case 2
            'Nicks >
            Select Case SubMenuIndex
                Case 0
                    'Identify
                    strInfo = InputBox(Language(330), Language(329))
                    If LenB(strInfo) > 0 Then
                        CurrentActiveServer.preExecute "/nickserv identify " & strInfo
                    End If
                Case 1
                    'Register
                    strInfo = InputBox(Language(330), Language(497))
                    If LenB(strInfo) > 0 And LenB(frmOptions.txtEmail.Text) > 0 Then
                        CurrentActiveServer.preExecute "/nickserv register " & strInfo & " " & frmOptions.txtEmail.Text
                    End If
                Case 2
                    'Ghost
                    strInfo = InputBox(Language(798), Language(797))
                    If LenB(strInfo) > 0 Then
                        strInfo2 = InputBox(Language(330), Language(797))
                        If LenB(strInfo2) > 0 Then
                            CurrentActiveServer.preExecute "/nickserv ghost " & strInfo & " " & strInfo2
                        End If
                    End If
                Case 3
                    'Release
                    strInfo = InputBox(Language(798), Language(838))
                    If LenB(strInfo) > 0 Then
                        strInfo2 = InputBox(Language(330), Language(838))
                        If LenB(strInfo2) > 0 Then
                            CurrentActiveServer.preExecute "/nickserv release " & strInfo & " " & strInfo2
                        End If
                    End If
                Case 4
                    'Drop
                    strInfo = MsgBox(Language(840), vbYesNo, Language(839))
                    If strInfo = vbYes Then
                        CurrentActiveServer.preExecute "/nickserv drop " & CurrentActiveServer.myNick
                    End If
            End Select
        Case 3
            'TO DO:
            'Create send-a-memo dialog
            
            'Memos >
            Select Case SubMenuIndex
                Case 0
                    'List memos
                    CurrentActiveServer.preExecute "/memoserv list"
                Case 1
                    strInfo = InputBox(Language(423), Language(422))
                    If LenB(strInfo) = 0 Then
                        Exit Sub
                    End If
                    strInfo2 = InputBox(Language(424), Language(422))
                    If LenB(strInfo2) = 0 Then
                        Exit Sub
                    End If
                    CurrentActiveServer.preExecute "/memoserv send " & strInfo & " " & strInfo2
                Case 2
                    'Read memo
                    strInfo = InputBox(Language(354), Language(352), "1")
                    If LenB(strInfo) = 0 Then
                        Exit Sub
                    End If
                    If Not (IsNumeric(strInfo) And Val(strInfo) = Int(Val(strInfo))) Then
                        MsgBox Language(355), vbExclamation, Language(352)
                        Exit Sub
                    End If
                    CurrentActiveServer.preExecute "/memoserv read " & strInfo
                Case 3
                    'Delete memo
                    strInfo = InputBox(Language(846), Language(353), "1")
                    If LenB(strInfo) = 0 Then
                        Exit Sub
                    End If
                    If Not (IsNumeric(strInfo) And Val(strInfo) = Int(Val(strInfo))) And LCase$(strInfo) <> "all" Then
                        MsgBox Language(355), vbExclamation, Language(353)
                        Exit Sub
                    End If
                    CurrentActiveServer.preExecute "/memoserv del " & strInfo
            End Select
        Case 1
            'Status >
            'store status
            If CurrentActiveServer.MyStatus <> SubMenuIndex Then
                CurrentActiveServer.MyStatus = SubMenuIndex
                'check the correct menu
                For i = 0 To 10
                    If nmnuIRC(1).Checked(i) <> (SubMenuIndex = i) Then
                        nmnuIRC(1).Checked(i) = SubMenuIndex = i
                    End If
                Next i
            End If
        Case 4
            'Channel Access >
            Select Case SubMenuIndex
                Case 0
                    'Add
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo3 = InputBox(Language(844), Language(846))
                        If LenB(strInfo3) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo3, 1) <> "#" Then
                            strInfo3 = "#" & strInfo3
                        End If
                    Else
                        strInfo3 = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    strInfo = InputBox(Language(798), Language(846))
                    If LenB(strInfo) > 0 Then
                        strInfo2 = InputBox(Language(800), Language(846))
                        If LenB(strInfo2) = 0 Then
                            Exit Sub
                        End If
                        If Val(strInfo2) <> strInfo2 Then
                            MsgBox Language(845), vbExclamation, Language(846)
                            Exit Sub
                        End If
                        CurrentActiveServer.preExecute "/chanserv access " & strInfo3 & " add " & strInfo & " " & strInfo2
                    End If
                Case 1
                    'del
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo2 = InputBox(Language(843), Language(845))
                        If LenB(strInfo2) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo2, 1) <> "#" Then
                            strInfo2 = "#" & strInfo2
                        End If
                    Else
                        strInfo2 = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    strInfo = InputBox(Language(801), Language(845))
                    If LenB(strInfo) > 0 Then
                        CurrentActiveServer.preExecute "/chanserv access " & strInfo2 & " del " & strInfo
                    End If
                Case 2
                    'view
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo = InputBox(Language(842), Language(844))
                        If LenB(strInfo) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo, 1) <> "#" Then
                            strInfo = "#" & strInfo
                        End If
                    Else
                        strInfo = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    
                    CurrentActiveServer.preExecute "/chanserv access " & strInfo & " list"
            End Select
        Case 5
            'Channel >
            Select Case SubMenuIndex
                Case 0
                    'register
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo3 = InputBox(Language(844), Language(871))
                        If LenB(strInfo3) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo3, 1) <> "#" Then
                            strInfo3 = "#" & strInfo3
                        End If
                    Else
                        strInfo3 = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    strInfo = InputBox(Language(869), Language(871))
                    If LenB(strInfo) > 0 Then
                        strInfo2 = InputBox(Language(870), Language(871))
                        If LenB(strInfo2) > 0 Then
                            CurrentActiveServer.preExecute "/chanserv register " & strInfo3 & " " & strInfo & " " & strInfo2
                        End If
                    End If
                Case 1
                    'identify
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo3 = InputBox(Language(844), Language(872))
                        If LenB(strInfo3) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo3, 1) <> "#" Then
                            strInfo3 = "#" & strInfo3
                        End If
                    Else
                        strInfo3 = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    strInfo = InputBox(Language(874), Language(872))
                    If LenB(strInfo) > 0 Then
                        CurrentActiveServer.preExecute "/chanserv identify " & strInfo3 & " " & strInfo
                    End If
                Case 2
                    'drop
                    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_Channel Then
                        strInfo3 = InputBox(Language(844), Language(873))
                        If LenB(strInfo3) = 0 Then
                            Exit Sub
                        End If
                        If Left$(strInfo3, 1) <> "#" Then
                            strInfo3 = "#" & strInfo3
                        End If
                    Else
                        strInfo3 = CurrentActiveServer.Tabs.SelectedItem.Caption
                    End If
                    strInfo = MsgBox(Language(867), vbYesNo, Language(873))
                    If strInfo = vbYes Then
                        CurrentActiveServer.preExecute "/chanserv drop " & strInfo3
                    End If
            End Select
    End Select
End Sub
Private Sub nmnuIRC_PopupMove(index As Integer)
    Static CodeCall As Boolean
    
    If CodeCall Then
        Exit Sub
    End If
    
    CodeCall = True
    If index = 4 Then
        nmnuIRC(index).RaisePopupMove
        nmnuIRC(index).FixPosition nmnuIRC(index), nmnuIRC(5)
    ElseIf index <> 0 Then
        nmnuIRC(index).RaisePopupMove
        nmnuIRC(index).FixPosition nmnuIRC(index), nmnuIRC(0)
    End If
    CodeCall = False
End Sub
Private Sub nmnuScript_MenuClick(SubMenuIndex As Integer)
    'the user clicked on an item inside the Script menu
    Select Case SubMenuIndex
        Case 0 'edit main script
            'call the EditScript routine in order to edit that script
            EditScript App.Path & "\scripts\main.vbs"
        Case 2 'reload
            InitSphereVariables
            RunScript "Begin"
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Options.KeepTrayRunning = False Or Ending Then
        If Not Restarting Then
            'the user is closing the program
            'set Cancel to False so the program doesn't end immediately
            '(we will end it later, after the fade transaction is completed)
            'then call the ShowMe method with the doShow boolean parameter
            'set to false
            xpBalloon.Unload
            ShowMe Not xLet(Cancel, True)
        Else
            currentWB = 0
        End If
    Else
        Cancel = True
        Me.Hide
        If GetSetting(App.EXEName, "InfoTips", "trayexit", "0") = "0" Then
            ShowInfoTip TrayExit
            SaveSetting App.EXEName, "InfoTips", "trayexit", "1"
        End If
    End If
End Sub
Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xpBalloon.HandleEvent X
End Sub
Public Sub tbText_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim intPrevPosition As Integer
    Dim strSymbolAdd As String
    Dim strSmileys As String
    Dim i As Integer
    Dim strSmileyPic As String
    Dim strPrevSmileyPic As String
    Dim strSmileyText As String
    Dim strSmileyData As String
    Dim ColIndex As Byte
    Dim strURL As String
    
    intPrevPosition = txtSend.SelStart
    
    Select Case Button.index
        Case 1
            'insert a smiley
            Set frmMinor = New frmCustom
            
            'build smileys table
            strSmileys = "<table><tr>"
            For i = 1 To UBound(ThisSmileyPack.AllSmileys)
                strSmileyText = ThisSmileyPack.AllSmileys(i).ShortcutText
                strSmileyPic = ThisSmileyPack.AllSmileys(i).FileName
                If strSmileyPic <> strPrevSmileyPic Then
                    strPrevSmileyPic = strSmileyPic
                    
                    If ColIndex > 5 Then
                        ColIndex = 0
                        strSmileys = strSmileys & "</tr><tr>"
                    End If
                    
                    ColIndex = ColIndex + 1
                    
                    strSmileys = strSmileys & "<td align=""center"" valign=""middle"">"
                    strSmileys = strSmileys & "<a href=""JavaScript:addsmiley('" & strSmileyText & "');"">"
                    strSmileys = strSmileys & img(App.Path & "/data/smileys/" & Options.SmileyPack & "/" & strSmileyPic, ThisSmileyPack.AllSmileys(i).ShortcutText, False, False)
                    strSmileys = strSmileys & "</a>"
                    strSmileys = strSmileys & "</td>"
                End If
            Next i
            strSmileys = strSmileys & "</tr></table>"
            
            If tbText.Visible Then
                imgMore_Click
            End If
            frmMinor.DialogData = strSmileys
            LoadDialog App.Path & "\data\dialogs\smiley.xml", frmMinor
            frmMinor.Left = txtSend.Left + txtSend.Width - frmMinor.Width
            frmMinor.Top = txtSend.Top - frmMinor.Height
        
        Case 3
            If Button.value = tbrPressed Then
                'user wants bold text
                strSymbolAdd = MIRC_BOLD
            ElseIf Button.value = tbrUnpressed Then
                'user wants to stop bold text
                If Len(txtSend.Text) > 0 Then
                    If intPrevPosition <> 0 Then
                        If Mid$(txtSend.Text, intPrevPosition, 1) = MIRC_BOLD Then
                            txtSend.Text = Left$(txtSend.Text, intPrevPosition - 1) & _
                                           Right$(txtSend.Text, Len(txtSend.Text) - intPrevPosition)
                            strSymbolAdd = vbNullString
                            intPrevPosition = intPrevPosition - 1
                        Else
                            strSymbolAdd = MIRC_BOLD
                        End If
                    Else
                        strSymbolAdd = MIRC_BOLD
                    End If
                Else
                    strSymbolAdd = MIRC_BOLD
                End If
            End If
        Case 4
            If Button.value = tbrPressed Then
                'user wants italic text
                strSymbolAdd = MIRC_ITALIC
            ElseIf Button.value = tbrUnpressed Then
                'user wants to stop italic text
                If Len(txtSend.Text) > 0 Then
                    If intPrevPosition <> 0 Then
                        If Mid$(txtSend.Text, intPrevPosition, 1) = MIRC_ITALIC Then
                            txtSend.Text = Left$(txtSend.Text, intPrevPosition - 1) & _
                                           Right$(txtSend.Text, Len(txtSend.Text) - intPrevPosition)
                            strSymbolAdd = vbNullString
                            intPrevPosition = intPrevPosition - 1
                        Else
                            strSymbolAdd = MIRC_ITALIC
                        End If
                    Else
                        strSymbolAdd = MIRC_ITALIC
                    End If
                Else
                    strSymbolAdd = MIRC_ITALIC
                End If
            End If
        Case 5
            If Button.value = tbrPressed Then
                'user wants underlined text
                strSymbolAdd = MIRC_UNDERLINE
            ElseIf Button.value = tbrUnpressed Then
                'user wants to end underline text
                If Len(txtSend.Text) > 0 Then
                    If intPrevPosition <> 0 Then
                        If Mid$(txtSend.Text, intPrevPosition, 1) = MIRC_UNDERLINE Then
                            txtSend.Text = Left$(txtSend.Text, intPrevPosition - 1) & _
                                           Right$(txtSend.Text, Len(txtSend.Text) - intPrevPosition)
                            strSymbolAdd = vbNullString
                            intPrevPosition = intPrevPosition - 1
                        Else
                            strSymbolAdd = MIRC_UNDERLINE
                        End If
                    Else
                        strSymbolAdd = MIRC_UNDERLINE
                    End If
                Else
                    strSymbolAdd = MIRC_UNDERLINE
                End If
            'need pic added to imagelist for palette, then the color selection
            'code will be added here. Also need a dialog file created for showing
            'all the smileys, then i can add the code for sending it here too.
            End If
    
        Case 7
            'insert colored text
            Set frmMinor = New frmCustom
            If tbText.Visible Then
                imgMore_Click
            End If
            LoadDialog App.Path & "\data\dialogs\color.xml", frmMinor
            frmMinor.Left = txtSend.Left + txtSend.Width - frmMinor.Width
            frmMinor.Top = txtSend.Top - frmMinor.Height
            
        Case 9
            'insert picture
            strURL = InputBox(Language(213), Language(562))
            If tbText.Visible Then
                imgMore_Click
            End If
            If LenB(strURL) > 0 Then
                strSymbolAdd = HTML_OPEN & "img src=""" & strURL & """" & HTML_CLOSE
            End If
            
        Case 10
            'insert hyperlink
            If tbText.Visible Then
                imgMore_Click
            End If
            LoadDialog App.Path & "\data\dialogs\hyperlink.xml"
            
        Case 11
            If tbText.Visible Then
                imgMore_Click
            End If
            MultilineText = Not MultilineText
            Form_Resize
            
    End Select
    If LenB(strSymbolAdd) > 0 Then
        txtSend.Text = Strings.Left$(txtSend.Text, intPrevPosition) & strSymbolAdd & Strings.Right$(txtSend.Text, Len(txtSend.Text) - intPrevPosition)
    End If
    txtSend.SelStart = intPrevPosition + Len(strSymbolAdd)
End Sub


Private Sub tmrAway_Timer()
    If Options.AwayEnabled Then
        AwayMins = AwayMins + 1
        If AwayMins >= Options.AwayMinutes Then
            GoAway True
            tmrAway.Enabled = False
        End If
    End If
End Sub
Private Sub tmrClearToolbar_Timer()
    sbar.Panels(1).Text = vbNullString
    tmrClearToolbar.Enabled = False
End Sub
Private Sub tmrHideURLBar_Timer()
    tmrHideURLBar.Enabled = False
    fraWebTab.Visible = False
End Sub

Private Sub tmrLag_Timer(index As Integer)
    'count the milliseconds
    tmrLag(index).Tag = tmrLag(index).Tag + 1
End Sub

Public Sub tmrMakeItQuicker_Timer()
    'this timer is used so that the webbrowser item
    'that contains irc items(channels, status, privates)
    'don't just after one new message appears
    'but in some time.
    'This will cause the program not to refresh
    'for EVERY new message if there are many
    'messages added together, but only once(at the end)
    'if this timer was called, the current channel/private or the status
    'will be updated, so we'll have to disable the timer
    '(it doesn't need to refresh again)
    tmrMakeItQuicker.Enabled = False
    'update the irc item: call the method buildStatus
    buildStatus
End Sub
Private Sub tmrNDCAudioRecord_Timer()
'    NDCAudioRegularOperations
End Sub
Private Sub tmrPanelRefreshSoon_Timer()
    wbPanel_DocumentComplete Nothing, vbNullString
    tmrPanelRefreshSoon.Enabled = False
    If Not frmOrganize Is Nothing Then
        If frmOrganize.Visible Then
            frmOrganize.SetFocus
        End If
    End If
End Sub

Private Sub tmrRefreshBuddy_Timer()
    Dim i As Integer, i2 As Integer
    Dim strTemp As String
    Dim ActiveServer As Variant
    
    '**** JOSH /REMOVE/ THIS IF YOU NEED TO DEVELOP!!! ****
    
    ' -------
        Exit Sub '<--
    ' -------
    
    showison = False

    If frmOptions.lstBdyNk.ListCount > 0 Then
        For i = 0 To frmOptions.lstBdyNk.ListCount - 1
            strTemp = strTemp & frmOptions.lstBdyNk.List(i) & " "
        Next i
        strTemp = Trim$(strTemp)
        For Each ActiveServer In ActiveServers
            If Not ActiveServer Is Nothing Then
                If ActiveServer.WinSockConnection.State = sckConnected Then
                    ActiveServer.preExecute "/ison " & strTemp
                Else
                    For i = 0 To UBound(AllBuddies)
                        AllBuddies(i).isOnline = False
                        For i2 = 0 To AllBuddies(i).ServerCount
                            If AllBuddies(i).Servers(i2) = ActiveServer.HostName Then
                                AllBuddies(i).Servers(i2) = vbNullString
                            ElseIf LenB(AllBuddies(i).Servers(i2)) > 0 Then
                                AllBuddies(i).isOnline = True
                            End If
                        Next i2
                    Next i
                End If
            End If
        Next
    End If
End Sub

Private Sub tmrRefreshSoon_Timer()
    'this timer is used to see if the web page contained
    'in the selected webpage tab was loaded
    'it "clicks" on the selected tab
    'so as the code inside the Click event
    'is executed
    '(that code tests if the site was loaded and
    'shows or hides the loading screen: see that code for more info)
    Dim CurrentTag As String 'stores the TabInfo item for the current tab
    
    'tsTabs_Click CurrentActiveServer.TabsIndex 'code-click on the selected tab, so that the code contained in that event executes
    
    'get the CurrentTag from the TabInfo collection.
    CurrentTag = CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)
    'if the current tab is not a web site...
    If Not CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_WebSite Then
        'we don't need this timer; disable it.
        tmrRefreshSoon.Enabled = False
    Else
        'check if CurrentTag wbStatus browser exists
        On Error Resume Next
        If wbStatus(CurrentTag).Busy And False Then
            'we should never get here
            'unless wbStatus with index CurrentTag
            'does not exist
        Else
            If wbStatus(CurrentTag).Busy Then
                'if the page is still loading, we'll have to check if it's loaded later again
                tmrRefreshSoon.Enabled = False
                tmrRefreshSoon.Enabled = True
            Else
                'if it's loaded it is already displayed(by event tsTabs_Click)
                'so we don't need this timer any more.
                tmrRefreshSoon.Enabled = False
                If wbStatus(CurrentTag).Left < 0 Then
                    tsTabs_Click CurrentActiveServer.Tabs.index
                    Form_Resize
                End If
            End If
        End If
    End If
End Sub
Public Function GetEmptyIdentD() As Integer
    Dim i As Integer
    For i = 0 To wsIdentD.Count
        If wsIdentD Is Nothing Then
            Exit For
        End If
    Next i
    GetEmptyIdentD = i
End Function
Private Sub tmrRegularOperations_Timer()
    Dim i As Integer
    Dim intTemp As Integer
    Dim ActiveServer As clsActiveServer
    
    'Check for NDC typing time-outs
    For i = 1 To UBound(NDCConnections)
        If NDCConnections(i).TypingTime + 2000 < GetTickCount Then
            'allow -t- packs to be received every 2 secs
            NDCConnections(i).Typing = False
            intTemp = CurrentActiveServer.GetTabFromNick(NDCConnections(i).strNicknameA)
            If CurrentActiveServer.Tabs.SelectedItem.index = intTemp Then
                sbar.Panels.Item(1).Text = vbNullString
            End If
        End If
    Next i
    
    For i = 0 To UBound(ActiveServers)
        If Not ActiveServers(i) Is Nothing Then
            Set ActiveServer = ActiveServers(i)
            ActiveServer.RegularOperations
        End If
    Next i
    
    'check for protection time-outs
    ClearProtectionMemory
End Sub
Private Sub tmrScriptingTimer_Timer()
    tmrScriptingTimer.Enabled = False
    RunScript ScriptingTimerNextCall
End Sub
Private Sub tmrShowSendToolTip_Timer()
    Dim strThisText As String, intCurrentPosition As Integer
    Dim strThisWord As String, strCompletedWord As String
    Dim t As Long, FinalWidth As Integer, FinalLeft As Integer
    Dim StartWidth As Integer, StartLeft As Integer
    Dim CurrentPOINT As Integer
    
    If Options.AutoComplete = False Then 'do NOT use (not Options.AutoComplete)
        tmrShowSendToolTip.Enabled = False
        Exit Sub
    End If
    'Text Completer
    'get the current text and the cursor position inside the text
    strThisText = txtSend.Text
    intCurrentPosition = txtSend.SelStart
    'if the cursor is at the beginning of the textbox we don't have anything to complete
    If intCurrentPosition = 0 Then
        'skip textcompleter
        Exit Sub
    End If
    'use GetWord to get the current word
    strThisWord = GetWord(strThisText, intCurrentPosition)
    'if there is no current word...
    If LenB(strThisWord) = 0 Then
        'skip textcompleter
        Exit Sub
    End If
    'use CompleteWord to complete the current word
    If lstSuggestions.Visible = False Then
        lstSuggestions.Clear
        strCompletedWord = CompleteWord(strThisWord, True)
        If lstSuggestions.ListCount > 0 Then
            lstSuggestions.Visible = True
        End If
    End If
    
    If LenB(strCompletedWord) = 0 Then
        Exit Sub
    End If
    tmrShowSendToolTip.Enabled = False
End Sub
Public Sub tsTabs_Click(index As Integer)
    'a tab is being clicked
    'we'll need to update the page displayed bellow
    
    'HERE: tsTabs(Index) Is CurrentActiveServer.Tabs
    '      Index = CurrentActiveServer.TabsIndex
    
    Dim CurrentCaption As String 'what's the caption of the currently selected tab?
    Dim CurrentTag As String 'what is the TabInfo of the currently selected tab?
    Dim i As Integer 'a counter variable for the loops
    Dim wbObject As WebBrowser
    Dim intTemp As Integer
    Dim bTemp As Byte
    Dim ThisTabType As Integer
    Dim ImageIndex As Integer
    Dim TabKey As String
    Dim CurrentTabType As Integer
    
    'if the tab was highlighted and the user clicked on it
    If tsTabs(index).SelectedItem.Tag = "Highlighted" Then
        'un-highlight it
        tsTabs(index).SelectedItem.Tag = "Not Highlighted"
        CurrentTabType = CurrentActiveServer.TabType(tsTabs(index).SelectedItem.index)
        ImageIndex = ImageFromType(CurrentTabType)
        tsTabs(index).SelectedItem.Image = ImageIndex
        TabKey = Switch(CurrentTabType = TabType_Channel, "c_", CurrentTabType = TabType_Private, "p_", CurrentTabType = TabType_Status, "s")
        TabKey = TabKey & GetServerIndexFromActiveServer(CurrentActiveServer)
        If Left$(TabKey, 1) <> "s" Then
            TabKey = TabKey & "_" & tsTabs(index).SelectedItem.Caption
        End If
        For i = 1 To tvConnections.Nodes.Count
            If tvConnections.Nodes.Item(i).Key = TabKey Then
                tvConnections.Nodes.Item(i).Image = ImageIndex
                Exit For
            End If
        Next i
    End If
    'get the current tab caption and store it in CurrentCaption
    CurrentCaption = Left$(tsTabs(index).SelectedItem.Caption, 2)
    
    ThisTabType = CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index)
    'if the current tab is an IRC-tab...
    If ThisTabType = TabType_Channel Or _
       ThisTabType = TabType_Private Or _
       ThisTabType = TabType_Status Or _
       ThisTabType = TabType_DCCFile Then
        'sbar.Panels.Item(1).Text = vbnullstring
        sbar.Panels.Item(2).Text = vbNullString
        Select Case ThisTabType
            Case TabType_Channel
                currentWB = WebBrowserIndex_Chan
                Set webdocCurrentIRCWindow = webdocChanMain
            Case TabType_Private
                CurrentNick = tsTabs(index).SelectedItem.Caption
                currentWB = WebBrowserIndex_Priv
                Set webdocCurrentIRCWindow = webdocPrivates
            Case TabType_DCCFile
                currentWB = WebBrowserIndex_DCC
            Case TabType_Status
                currentWB = WebBrowserIndex_Priv
                Set webdocCurrentIRCWindow = webdocPrivates
        End Select
        LockWindowUpdate Me.hwnd
        ShowBrowser currentWB, True
        '...and hide all other webBrowsers
        For Each wbObject In wbStatus
            If wbObject.index <> currentWB Then
                ShowBrowser wbObject.index, False
            End If
        Next wbObject
        If CurrentActiveServer.intCurrentlySelected <> CurrentActiveServer.Tabs.SelectedItem.index Then
            'and update the webBrowser contents
            buildStatus
        End If
        LockWindowUpdate 0&
        If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
            intTemp = GetNDCFromNickname(CurrentActiveServer.Tabs.SelectedItem.Caption)
            If intTemp <> -1 Then
                If wsNDC.Item(intTemp).State = sckConnected Then
                    bTemp = NDCConnections(intTemp).RemoteStatus
                    'if the NDC's private is focused
                    'show the text for the status
                    If bTemp = 0 Then
                        sbar.Panels(2).Text = vbNullString
                    Else
                        sbar.Panels(2).Text = Replace( _
                                              Replace( _
                                                Language(421), "%1", NDCConnections(intTemp).strNicknameA), _
                                                               "%2", Language(bTemp + 397))
                    End If
                    If NDCConnections(intTemp).Typing Then
                        AddNews Replace(Language(318), "%1", NDCConnections(intTemp).strNicknameA) & "..."
                    End If
                End If
            End If
        End If
    Else
        'sbar.Panels.Item(2).Text = vbnullstring
        'unable to retrieve StatusText
        On Error Resume Next
        sbar.Panels.Item(1).Text = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).StatusText
        'else, if it's a website tab....
        'the TabInfo contains a leading !, remove it
        CurrentTag = CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)
        'if this web site is still busy(loading) display loading screen
        If wbStatus(CurrentTag).Busy And Options.HTMLLoading Then
            'set the webbrowser-to-be-displayed to 1
            CurrentTag = WebBrowserIndex_Loading 'loading
            tmrRefreshSoon.Enabled = True
        End If
        'if it's not loading the webbrowser-to-be-display will be the current web site
        'go trough all webbrowsers and hide them all except of the current one, which we'll show.
        '(this is either the current web site or the loading screen)
        ShowBrowser CurrentTag, True
        Form_Resize
        For Each wbObject In wbStatus
            If wbObject.index <> CurrentTag Then
                ShowBrowser wbObject.index, False
            End If
        Next wbObject
        'set the CurrentWB(select WebBrowser) to be either the loading webbrowser or the website itself
        currentWB = CurrentTag
    End If
    txtSend.Visible = CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_DCCFile And _
                      CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_WebSite
    fraMore.Visible = txtSend.Visible
        
    If CurrentActiveServer.intCurrentlySelected <> CurrentActiveServer.Tabs.SelectedItem.index Then
        'the tab was changed; update controls(espessially webbrowsers) to the new view
        Form_Resize 'tab changed
        'also update the tabs bar(it may should have a different size now)
        UpdateTabsBar
    End If
    CurrentActiveServer.intCurrentlySelected = CurrentActiveServer.Tabs.SelectedItem.index
    
    If strCurrentPanel = "avatar" Then
        On Error Resume Next
        If CurrentActiveServer.TabType(CurrentActiveServer.intCurrentlySelected) <> CurrentActiveServer.TabType(CurrentActiveServer.intPreviousTabIndex) Then
            If CurrentActiveServer.TabType(CurrentActiveServer.intCurrentlySelected) = TabType_Private Or CurrentActiveServer.TabType(CurrentActiveServer.intPreviousTabIndex) = TabType_Private Then
                tmrPanelRefreshSoon.Enabled = False
                tmrPanelRefreshSoon.Enabled = True
            End If
        ElseIf CurrentActiveServer.TabType(CurrentActiveServer.intCurrentlySelected) = TabType_Private And CurrentActiveServer.TabType(CurrentActiveServer.intPreviousTabIndex) = TabType_Private Then
            If CurrentActiveServer.intCurrentlySelected <> CurrentActiveServer.intPreviousTabIndex Then
                tmrPanelRefreshSoon.Enabled = False
                tmrPanelRefreshSoon.Enabled = True
            End If
        End If
    End If
    
    If fraWebTab.Visible Then
        If CurrentActiveServer.TabType(CurrentActiveServer.intCurrentlySelected) <> TabType_WebSite Then
            fraWebTab.Visible = False
        Else
            If txtURL.Text <> wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.intCurrentlySelected)).LocationURL Then
                txtURL.Text = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.intCurrentlySelected)).LocationURL
                txtURL.SelStart = 0
                txtURL.SelLength = Len(txtURL.Text)
                txtURL.SetFocus
            End If
        End If
    Else
        On Error Resume Next
        txtSend.SetFocus
    End If
End Sub
Private Sub tsTabs_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    'execute hot key
    Form_KeyDown KeyCode, Shift
End Sub
Private Sub tsTabs_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim sngX As Single
    
    CurrentActiveServer.intPreviousTabIndex = CurrentActiveServer.Tabs.SelectedItem.index
    
    sngX = X + tsTabs(index).Left
    
    'select the appropriate tab
    For i = 2 To tsTabs(index).Tabs.Count
        If tsTabs(index).Tabs.Item(i).Left > sngX Then
            Set tsTabs(index).SelectedItem = tsTabs(index).Tabs.Item(i - 1)
            GoTo Finished_Selecting
        End If
    Next i
    'no tab was selected, select the last tab
    Set tsTabs(index).SelectedItem = tsTabs(index).Tabs.Item(tsTabs(index).Tabs.Count)
Finished_Selecting:
    If Shift = vbShiftMask Then
        'split
        'TO DO: Split
        'SplitEnabled = Not SplitEnabled
        'SplitIndex = tsTabs(Index).SelectedItem
        'Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(CurrentActiveServer.intPreviousTabIndex)
        'buildStatus
    Else
        'the user mouse-downed a tab
        If Button = 2 Then
            'if he/she right clicked, display the tabs popup menu(with mnuCloseTab the default item)
            'depending on the type of tab, we'll display different menu items
            i = CurrentActiveServer.TabType(tsTabs(index).SelectedItem.index)
            On Error Resume Next
            mnuTabDisconnect.Visible = True 'will be set later again
            mnuTabConnect.Visible = i = TabType_Status
            mnuCloseTab.Visible = i = TabType_WebSite Or i = TabType_Private Or i = TabType_DCCFile
            mnuChanTabLeave.Visible = i = TabType_Channel
            mnuChanTabRejoin.Visible = i = TabType_Channel
            mnuWebTabFav.Visible = i = TabType_WebSite
            mnuWebTabStop.Visible = i = TabType_WebSite
            mnuWebTabBack.Visible = i = TabType_WebSite
            mnuWebTabForward.Visible = i = TabType_WebSite
            mnuWebTabRefresh.Visible = i = TabType_WebSite
            mnuTabDisconnect.Visible = i = TabType_Status
            mnuTabDisconnect.Enabled = ActiveServers(CurrentActiveServer).WinSockConnection.State <> sckClosed And _
                                        ActiveServers(CurrentActiveServer).WinSockConnection.State <> sckClosing And _
                                        ActiveServers(CurrentActiveServer).WinSockConnection.State <> sckError
            Me.PopupMenu mnuTabsPop, , , , IIf(mnuCloseTab.Visible, mnuCloseTab, IIf(mnuTabConnect.Visible, mnuTabConnect, mnuChanTabLeave))
        End If
    End If
End Sub
Private Sub tsTabs_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_WebSite Then
        Exit Sub
    End If
    
    For i = 1 To tsTabs(index).Tabs.Count
        If tsTabs(index).Tabs.Item(i).Left > X + tsTabs.Item(index).Left Then
            If tsTabs(index).SelectedItem.index <> i - 1 Then
                Exit Sub
            Else
                GoTo Got_It
            End If
        End If
    Next i
    If tsTabs(index).SelectedItem.index <> tsTabs(index).Tabs.Count Then
        Exit Sub
    End If
    
Got_It:
    If Not fraWebTab.Visible Then
       ' On Error Resume Next
        txtURL.Text = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).LocationURL
        fraWebTab.Left = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Left
        fraWebTab.Width = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Width
        fraWebTab.Top = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Top
        txtURL.Width = fraWebTab.Width - txtURL.Left * 2
        txtURL.SelStart = 0
        txtURL.SelLength = Len(txtURL.Text)
        
        fraWebTab.ZOrder 0
        
        fraWebTab.Visible = True
        txtURL.SetFocus
    End If
    tmrHideURLBar.Enabled = False
    tmrHideURLBar.Enabled = True
End Sub
Private Sub tsTabs_OLEDragDrop(index As Integer, Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    tsTabs_MouseDown index, Button, Shift, X, Y
    
    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
        On Error GoTo Invalid_File
        If Data.Files.Count >= 1 Then
            CurrentNick = tsTabs(index).SelectedItem.Caption
            strDCCPresetFile = Data.Files.Item(1)
            mnuDCCSend_Click
        End If
    End If
Invalid_File:
End Sub
Private Sub tvConnections_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub
Private Sub tvConnections_Collapse(ByVal Node As ComctlLib.Node)
    If Node.Key = "root" Then
        'the root node can not be collapsed
        Node.Expanded = True
    End If
End Sub
Private Sub tvConnections_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim PointNode As ComctlLib.Node
    Dim intServersCount As Integer
    Dim i As Integer
    
    If Button = 2 Then
        Set PointNode = tvConnections.HitTest(X, Y)
        Set tvConnections.SelectedItem = PointNode
        If Not PointNode Is Nothing Then
            tvConnections_NodeClick PointNode
        End If
        tvConnections.SetFocus
        For i = 0 To UBound(ActiveServers)
            If Not ActiveServers(i) Is Nothing Then
                intServersCount = intServersCount + 1
                If intServersCount > 1 Then
                    Exit For
                End If
            End If
        Next i
        mnuRemoveServer.Enabled = intServersCount > 1 And Not PointNode Is Nothing
        PopupMenu mnuConnectionsPop
    End If
End Sub
Private Sub tvConnections_NodeClick(ByVal Node As ComctlLib.Node)
    Dim intPos As Integer
    Dim strTabCaption As String
    Dim i As Integer
    
    If Node.Key <> "root" Then
        Select Case Left$(Node.Key, 1)
            Case "s"
                ActiveServerMakeCurrent ActiveServers(Right$(Node.Key, Len(Node.Key) - 1))
            Case "c"
                intPos = InStr(3, Node.Key, "_")
                ActiveServerMakeCurrent ActiveServers(Mid$(Node.Key, 3, intPos - 3))
                strTabCaption = Right$(Node.Key, Len(Node.Key) - intPos)
                For i = 1 To CurrentActiveServer.Tabs.Tabs.Count
                    If CurrentActiveServer.Tabs.Tabs.Item(i).Caption = strTabCaption Then
                        Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(i)
                    End If
                Next i
            Case "w"
                If Node.Key <> "wbroot" Then
                    intPos = InStr(3, Node.Key, "_")
                    ActiveServerMakeCurrent ActiveServers(Mid$(Node.Key, 3, intPos - 3))
                    strTabCaption = Right$(Node.Key, Len(Node.Key) - intPos)
                    For i = 1 To CurrentActiveServer.Tabs.Tabs.Count
                        If CurrentActiveServer.TabInfo(i) = strTabCaption Then
                            Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(i)
                        End If
                    Next i
                End If
            Case "d"
                intPos = InStr(3, Node.Key, "_")
                ActiveServerMakeCurrent ActiveServers(Mid$(Node.Key, 3, intPos - 3))
                strTabCaption = Right$(Node.Key, Len(Node.Key) - intPos)
                For i = 1 To CurrentActiveServer.Tabs.Tabs.Count
                    If CurrentActiveServer.Tabs.Tabs.Item(i).Caption & "_chat" & i = strTabCaption Then
                        Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(i)
                    End If
                Next i
            Case "p"
                intPos = InStr(3, Node.Key, "_")
                ActiveServerMakeCurrent ActiveServers(Mid$(Node.Key, 3, intPos - 3))
                strTabCaption = Right$(Node.Key, Len(Node.Key) - intPos)
                For i = 1 To CurrentActiveServer.Tabs.Tabs.Count
                    If CurrentActiveServer.Tabs.Tabs.Item(i).Caption = strTabCaption Then
                        Set CurrentActiveServer.Tabs.SelectedItem = CurrentActiveServer.Tabs.Tabs.Item(i)
                    End If
                Next i
        End Select
    End If
End Sub
Private Sub txtSend_Change()
    Dim intTempPos As Integer
    
    If Not MultilineText Then
        If InStr(1, txtSend.Text, vbNewLine) > 0 Then
            intTempPos = txtSend.SelStart
            txtSend.Text = Replace(txtSend.Text, vbNewLine, vbNullString)
            txtSend.SelStart = intTempPos
            On Error Resume Next 'txtSend invisible, if on website tab
            txtSend.SetFocus
        End If
    End If
End Sub
Private Sub txtSend_GotFocus()
    'tbText.Visible = True
End Sub
Public Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    'the user pressed a key or typed something in the textbox
    Dim strThisWord As String 'the current word being typed
    Dim strThisText As String 'the current text inside the textbox
    Dim strLines() As String 'all the lines of a multiline text
    Dim strCompletedWord As String 'if the user pressed tab this is the completed nickname or word
    Dim intCurrentPosition As Integer 'the current position of the cursor inside the textobx
    Dim strNavigation As String
    Dim intTemp As Integer
    Dim i As Integer
    Static frmWinampControl As frmCustom
    Static intHistoryIndex As Integer 'integer variable holding the position in the history array(if any)
    
    If Left$(txtSend.Text, 2) = vbNewLine Then
        txtSend.Text = Right$(txtSend.Text, Len(txtSend.Text) - 2)
    End If
    
    If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
        intTemp = GetNDCFromNickname(CurrentActiveServer.Tabs.SelectedItem.Caption)
        If intTemp <> -1 Then
            If wsNDC.Item(intTemp).State = sckConnected Then
                If NDCConnections(intTemp).IntroPackSent Then
                    If NDCConnections(intTemp).TypingSent + 1000 < GetTickCount Then
                        'send a -t- every sec
                        wsNDC.Item(intTemp).SendData "t"
                        NDCConnections(intTemp).TypingSent = GetTickCount
                    End If
                End If
            End If
        End If
    End If
    
    If Shift = 0 Or Shift = 1 Then
        'if there is something in the textbox and the user pressed enter
        Select Case KeyCode
            Case vbKeyReturn
                If lstSuggestions.Visible Then
                    If lstSuggestions.ListIndex = -1 Then
                        lstSuggestions.ListIndex = 0
                    Else
                        lstSuggestions_Click
                    End If
                ElseIf Not MultilineText Then
                    CheckGoBack
                    If LenB(txtSend.Text) > 0 Then
                        If boolSaveNextType Then
                            MessageHistory(UBound(MessageHistory)) = txtSend.Text
                            ReDim Preserve MessageHistory(xLet(intHistoryIndex, UBound(MessageHistory) + 1))
                        Else
                            boolSaveNextType = True
                        End If
                        'inform any loaded plugins
                        For i = 0 To NumToPlugIn.Count - 1
                            If Plugins(i).boolLoaded = True And Not Plugins(i).objPlugIn Is Nothing Then
                                Plugins(i).objPlugIn.Sending
                            End If
                        Next i
                        'send the text using preExecute
                        ndScript.ExecuteStatement "DoSend = True"
                        RunScript "Sending"
                        If ndScript.Eval("DoSend") Then
                            If LCase$(Left$(txtSend.Text, 5)) = "/ison" Then showison = True
                            CurrentActiveServer.preExecute txtSend.Text
                        End If
                    End If
                ElseIf Shift = 1 Then
                    CheckGoBack
                    MultilineText = False
                    strLines = Split(txtSend.Text, vbNewLine)
                    For i = 0 To UBound(strLines)
                        txtSend.Text = strLines(i)
                        txtSend_KeyDown vbKeyReturn, 0
                    Next i
                    Form_Resize
                End If
                
            Case vbKeyF10
                MaxMode = Not MaxMode
                If MaxMode Then
                    Set frmFullScreen = New frmCustom
                    LoadDialog App.Path & "/data/dialogs/full.xml", frmFullScreen
                    frmFullScreen.wbCustom(1).Width = Screen.Width + 100
                    frmFullScreen.wbCustom(1).Height = Screen.Height + 100
                    Select Case CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index)
                        Case TabType_Channel
                            strNavigation = App.Path & "\data\html\imports\frameset.html"
                        Case TabType_Private, TabType_Status
                            strNavigation = App.Path & "\temp\priv.html"
                        Case Else
                            MaxMode = False
                            frmFullScreen.Hide
                            Set frmFullScreen = Nothing
                            Exit Sub
                    End Select
                    frmFullScreen.wbCustom(1).Navigate2 strNavigation
                    frmFullScreen.wbCustom(1).Visible = True
                    frmFullScreen.Show
                    frmFullScreen.wbCustom(1).SetFocus
                    Set webdocFullScreen = frmFullScreen.wbCustom(1).Document
                Else
                    frmFullScreen.Hide
                    Set frmFullScreen = Nothing
                    Set webdocFullScreen = Nothing
                    Set webdocFullScreenChanMain = Nothing
                    Set webdocFullScreenChanNicks = Nothing
                End If
                lstSuggestions.Visible = False
                
'            Case vbKeyF6
'                If frmWinampControl Is Nothing Then
'                    Set frmWinampControl = New frmCustom
'                    LoadDialog App.Path & "/data/dialogs/winamp.xml", frmWinampControl
'                    frmWinampControl.Move Me.Left + Me.ScaleWidth - frmWinampControl.Width, Me.Top + (Me.Height - Me.ScaleHeight) + tsTabs.Top
'                    xBasic.xWindowTopMost frmWinampControl.hwnd, True
'                    Me.txtSend.SetFocus
'                Else
'                    frmWinampControl.Hide
'                    Unload frmWinampControl
'                    Set frmWinampControl = Nothing
'                End If
            
            'if the user pressed tab, we'll have to complete his word
            Case vbKeyTab
                If lstSuggestions.Visible Then
                    If lstSuggestions.ListIndex = -1 Then
                        lstSuggestions.ListIndex = 0
                    Else
                        lstSuggestions_Click
                    End If
                ElseIf Options.AutoComplete Then
                    'Text Completer
                    tmrShowSendToolTip_Timer
                    'if there are suggestions
                    If lstSuggestions.ListCount > 0 Then
                        lstSuggestions.ListIndex = 0
                    Else
                        GiveFocus = True
                    End If
                End If
            Case vbKeyUp
                If lstSuggestions.Visible Then
                    intTemp = lstSuggestions.ListIndex - 1
                    If intTemp < -1 Then
                        intTemp = lstSuggestions.ListCount - 1
                    End If
                    BolLstSuggstionsCodeClick = True
                    lstSuggestions.ListIndex = intTemp
                    BolLstSuggstionsCodeClick = False
                ElseIf Not MultilineText Then
                    If intHistoryIndex > 0 Then
                        intHistoryIndex = intHistoryIndex - 1
                    End If
                    txtSend.Text = MessageHistory(intHistoryIndex)
                    txtSend.SelStart = Len(txtSend.Text)
                    lstSuggestions.Visible = False
                End If
            Case vbKeyDown
                If lstSuggestions.Visible Then
                    intTemp = lstSuggestions.ListIndex + 1
                    If intTemp >= lstSuggestions.ListCount Then
                        intTemp = 0
                    End If
                    BolLstSuggstionsCodeClick = True
                    lstSuggestions.ListIndex = intTemp
                    BolLstSuggstionsCodeClick = False
                ElseIf Not MultilineText Then
                    If intHistoryIndex < UBound(MessageHistory) Then
                        intHistoryIndex = intHistoryIndex + 1
                    End If
                    txtSend.Text = MessageHistory(intHistoryIndex)
                    txtSend.SelStart = Len(txtSend.Text)
                    lstSuggestions.Visible = False
                End If
            Case vbKeyEscape
                intHistoryIndex = UBound(MessageHistory)
                txtSend.Text = vbNullString
                lstSuggestions.Visible = False
            Case 65 To 90, 191, 8 'letter, LETTER, number or symbol, /, backspace
                If KeyCode = AscW("/") And txtSend.Text = "/" Then
                    lstSuggestions.Left = txtSend.Left
                    lstSuggestions.Top = txtSend.Top - lstSuggestions.Height
                End If
                If Options.AutoComplete Then
                    tmrShowSendToolTip.Enabled = False
                    tmrShowSendToolTip.Enabled = True
                End If
                lstSuggestions.Visible = False
            Case Else
                GoTo Execute_HotKey
        End Select
    Else
Execute_HotKey:
        Form_KeyDown KeyCode, Shift
    End If
End Sub
Private Sub txtSend_LostFocus()
    'the focus of the textbox was lost
    'was this done by code?
    tbText.Visible = False
    imgMore.Tag = "5"
    If GiveFocus Then
        'yes, and we didn't want it to happen.
        '(the next focus may be lost by the user; set GiveFocus to false)
        GiveFocus = False
        'give the focus back
        txtSend.SetFocus
    End If
End Sub
'Private Sub txtSend_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'the user mouse-downed on the textbox
    'if he/she right-clicked...
    'If Button = 2 Then
        '...disable it and re-enable it so no
        'popup menu is shown.
        'txtSend.Enabled = False
        'txtSend.Enabled = True
        'LoadDialog App.Path & "\data\dialogs\CustomText.xml"
        'show the custom message form.
        'first load it
        'Load frmCustomMessage
        'and set the current text to what is already typed in the textbox
        'frmCustomMessage.rtfMessage.Text = txtSend.Text
        'show the curstommessage dialog
        'frmCustomMessage.Show vbModal
    'End If
'End Sub
Private Sub Form_GotFocus()
    'focus the textbox where messages to be send are written by the user
    'when the user selects the form
    'create an error trap in case another modal form is
    'currently selected, or the main form is minimized
    On Error Resume Next
    'set focus to the textbox
    txtSend.SetFocus
End Sub
Private Sub txtSend_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Stop
End Sub
Private Sub txtSend_DragDrop(Source As Control, X As Single, Y As Single)
    'Stop
End Sub
Private Sub txtURL_Change()
    tmrHideURLBar.Enabled = False
    tmrHideURLBar.Enabled = True
End Sub
Private Sub txtURL_Click()
    tmrHideURLBar.Enabled = False
    tmrHideURLBar.Enabled = True
End Sub
Private Sub txtURL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Navigate2 txtURL.Text
        txtURL.Text = vbNullString
        fraWebTab.Visible = False
    End If
End Sub
Private Sub wbBack_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    wbStatus_BeforeNavigate2 0, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel
End Sub
Private Sub wbBack_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim webdocBack As HTMLDocument
    Set webdocBack = wbBack.Document
    
    On Error Resume Next
    webdocBack.bgColor = ThisSkin.BackgroundColor
    webdocBack.All.Item("skin_pic").src = App.Path & "\data\skins\" & ThisSkin.SkinPic
    webdocBack.All.Item("skin_pic").alt = Language(506)
    webdocBack.All.Item("skin_toolbox_help").src = App.Path & "\data\skins\" & ThisSkin.Icon_Help
    webdocBack.All.Item("skin_toolbox_help").alt = Language(136)
    webdocBack.All.Item("skin_toolbox_options").src = App.Path & "\data\skins\" & ThisSkin.Icon_Options
    webdocBack.All.Item("skin_toolbox_options").alt = Language(4)
    webdocBack.All.Item("skin_toolbox_join").src = App.Path & "\data\skins\" & ThisSkin.Icon_Join
    webdocBack.All.Item("skin_toolbox_join").alt = Language(323)
End Sub
Private Sub wbPanel_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    'this will execute any nodeScripts
    'and also set the value of Cancel to True if required
    '(as it's passed byRef)
    wbStatus_BeforeNavigate2 0, pDisp, URL, Flags, TargetFrameName, PostData, Headers, Cancel
End Sub
Private Sub wbPanel_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim boolAvatarExists As Boolean
    Dim intTemp As Integer
    Dim i As Integer
    Dim strServer As String
    Dim strServerDescription As String
    Dim strServerHostname As String
    Dim strServerPort As String
    Dim strDescriptionQuotes As String
    Dim intFL As Integer
    
    On Error Resume Next
    Set webdocPanel = wbPanel.Document
    With webdocPanel.All
        Select Case strCurrentPanel
            Case "join"
                '# <panel id="join">
                'join panel: past channels
                xNodeTag webdocPanel, "past_chan", GetPastChannels, "xnode_xpanel_"
                
                'join panel: language entries
                xNodeTag webdocPanel, "lang_past_chans", Language(450)
                xNodeTag webdocPanel, "lang_other_chan", Language(454)
                xNodeTag webdocPanel, "lang_join", Language(453)
                xNodeTag webdocPanel, "lang_all_chans", Language(466)
                xNodeTag webdocPanel, "lang_list_chans", Language(467)
                
                webdocPanel.Title = Language(326)
                '# </panel>
            Case "avatar"
                '# <panel id="avatars">
                'avatars panel: lang
                xNodeTag webdocPanel, "lang_users_avatar", Language(458)
                xNodeTag webdocPanel, "lang_my_avatar", Language(459)
                xNodeTag webdocPanel, "lang_change_avatar", Language(460)
                xNodeTag webdocPanel, "lang_remove_avatar", Language(461)
                xNodeTag webdocPanel, "lang_add_avatar", Language(462)
            
                webdocPanel.Title = Language(456)
                
                'avatars xnode xpanel
                xNodeTag webdocPanel, "lang_select_private", Language(553)
                xNodeTag webdocPanel, "nickname", IIf(CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private, CurrentActiveServer.Tabs.SelectedItem.Caption, vbNullString), "xnode_xpanel_"
                
                'TO DO: Add png, tif, (gif) support
                'Pictures|*.jpg;*.jpeg;*.gif;*.tif;*.png|JPEG Images|*.jpg;*.jpeg|GIF Images|*.gif|PNG Images|*.png|TIF Images|*.tif
    
                '"No Avatar" or display avatar
                'xNodeTag webdocPanel, "remote_avatar", Language(IIf(TabType(tsTabs.SelectedItem.Index) = TabType_Private, 463, 465)), "xnode_xpanel_"
                If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
                    'check if he/she has an avatar
                    If FS.FileExists(App.Path & "/conf/avatars/" & CurrentActiveServer.Tabs.SelectedItem.Caption & ".jpg") Then
                        'now show it
                        xNodeTag webdocPanel, "remote_avatar", "<img src=""" & App.Path & "/conf/avatars/" & CurrentActiveServer.Tabs.SelectedItem.Caption & ".jpg"" width=""150px"" height=""150px"">", "xnode_xpanel_"
                    Else
                        intTemp = GetNDCFromNickname(CurrentActiveServer.Tabs.SelectedItem.Caption)
                        If intTemp <> -1 Then
                            'ndc connection is established
                            '"No Avatar"
                            xNodeTag webdocPanel, "remote_avatar", Language(463), "xnode_xpanel_"
                        Else
                            'no NDC connection
                            xNodeTag webdocPanel, "remote_avatar", Language(468), "xnode_xpanel_"
                        End If
                    End If
                Else
                    xNodeTag webdocPanel, "remote_avatar", Language(465), "xnode_xpanel_"
                End If
                
                boolAvatarExists = FS.FileExists(App.Path & "/conf/myavatar.jpg")
                If boolAvatarExists Then
                    xNodeTag webdocPanel, "local_avatar", "<img src=""" & App.Path & "/conf/myavatar.jpg"" width=""150px"" height=""150px"">", "xnode_xpanel_"
                Else
                    xNodeTag webdocPanel, "local_avatar", Language(464), "xnode_xpanel_"
                End If
                
                xNodeTagShow webdocPanel, "change_avatar", boolAvatarExists, "xnode_xpanel_"
                xNodeTagShow webdocPanel, "remove_avatar", boolAvatarExists, "xnode_xpanel_"
                xNodeTagShow webdocPanel, "add_avatar", Not boolAvatarExists, "xnode_xpanel_"
                '# </panel>
            Case "connect"
                '# <panel id="connect">
                xNodeTag webdocPanel, "lang_past_servs", Language(595)
                xNodeTag webdocPanel, "lang_new_serv", Language(470)
                xNodeTag webdocPanel, "lang_displayname", Language(471)
                xNodeTag webdocPanel, "lang_hostname", Language(472)
                xNodeTag webdocPanel, "lang_port", Language(473)
                xNodeTag webdocPanel, "lang_serv_list", Language(474)
                xNodeTag webdocPanel, "lang_list_servs", Language(475)
                xNodeTag webdocPanel, "lang_connect", Language(2)
                xNodeTag webdocPanel, "lang_organize", Language(594)
                xNodeTag webdocPanel, "lang_hostname_spaces", Language(602)
                xNodeTag webdocPanel, "lang_invalid_port", Language(114)
                xNodeTag webdocPanel, "lang_no_hostname", Language(603)
                xNodeTag webdocPanel, "lang_no_port", Language(604)
                
                webdocPanel.All.Item("xnode_lang_organize").Style.display = IIf(ServerListCount() = 0, "none", vbNullString)
                
                xNodeTag webdocPanel, "xpanel_past_serv", CreateServersList()
                
                webdocPanel.Title = Language(2)
                '# </panel>
            Case "favorites"
                '# <panel id="favorites">
                xNodeTag webdocPanel, "lang_newwebname", Language(482)
                xNodeTag webdocPanel, "lang_newwebURL", Language(483)
                xNodeTag webdocPanel, "lang_save_new_web", Language(571)
                xNodeTag webdocPanel, "fav_new_web", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Add & """> " & Language(486)
                xNodeTag webdocPanel, "favwebs", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Web & """> " & Language(481)
                xNodeTag webdocPanel, "chanfavs", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Channel & """> " & Language(495)
                xNodeTag webdocPanel, "serverfavs", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Server & """> " & Language(558)
                
                'favorites panel: favorite websites
                xNodeTag webdocPanel, "favwebs", GetFavWebs, "xnode_xpanel_"
                'favorites panel: favorite servers
                
                xNodeTag webdocPanel, "xpanel_past_servers", CreateServersList()
                
                'favorites panel: favorite channels
                xNodeTag webdocPanel, "past_chan", GetPastChannels, "xnode_xpanel_"
                
                webdocPanel.Title = Language(488)
                '# </panel>
            Case "buddylist"
                '# <panel id="buddy list">
                xNodeTag webdocPanel, "buddylist", Language(310)
                xNodeTag webdocPanel, "skin_add_icon", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Add & """ />"
                xNodeTag webdocPanel, "skin_options_icon", "<img src=""" & App.Path & "/data/skins/" & ThisSkin.Icon_Options & """>"
                xNodeTag webdocPanel, "panel_options", "<a href=""Nodescript:/buddyoptview""> " & Language(503) & " </a>"
                xNodeTag webdocPanel, "panel_addbuddy", "<a href=""Nodescript:/buddyadd""> " & Language(565) & " </a>"
                
                xNodeTag webdocPanel, "showbuddies", GetBuddies, "xnode_xpanel_"

                webdocPanel.Title = Language(310)
                '# </panel>
        End Select
    End With
End Sub
Private Sub wbPanel_TitleChange(ByVal Text As String)
    lblPanelTitle.Caption = Text ' ;)
End Sub
Private Sub wbStatus_BeforeNavigate2(index As Integer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)    'either a channel/private/status webbrowser object is navigating to somewhere
    'or a website tab is navigating.
    Dim sURL As String 'the URL unescaped.
    Dim intTemp As Integer
    
    If boolBuildingPrimary Then
        Exit Sub
    End If
        
    Select Case URL
        Case App.Path & "\temp\main_chan.html", App.Path & "\temp\nicklist.html", App.Path & "\data\html\topic_frame.html"
            boolBuildingPrimary = True
            Cancel = True
            wbStatus(index).Navigate2 App.Path & "\data\html\imports\frameset.html"
            Wait 0.2
            Set webdocChanFrameSet = wbStatus(index).Document
            On Error Resume Next
            Set webdocChanTopic = webdocChanFrameSet.parentWindow.frames(0).Document
            Set webdocChanMain = webdocChanFrameSet.parentWindow.frames(1).Document
            Set webdocChanNicklist = webdocChanFrameSet.parentWindow.frames(2).Document
            boolBuildingPrimary = False
            Wait 0.3
            buildStatus CurrentActiveServer
            Exit Sub
    End Select
    
    'if the new url starts with NodeScript we'll have to execute
    'that script and stop navigation.
    If Strings.Left$(Strings.LCase$(URL), Len("NodeScript:")) = "nodescript:" Then
        'yes, it's a NodeScript
        'UnEscape( = replace HTML notation, for example %20 becomes space) the URL and store the result in sURL
        sURL = UnEscape(URL)
        'execute the command that should be executed
        CurrentActiveServer.preExecute Strings.Right$(sURL, Len(sURL) - Len("NodeScript:")), False
        'and cancel navigation
        Cancel = True
    ElseIf Strings.Left$(Strings.LCase$(URL), Len("JavaScript:")) = "javascript:" Then
        'do nothing; let it navigate <-- commented. don't let it!
    Else
        If index = WebBrowserIndex_Chan Or index = WebBrowserIndex_Priv Then
            If CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) <> TabType_WebSite Then
                CurrentActiveServer.preExecute "/browse """ & URL & """", False
                Cancel = True
            End If
        End If
    End If
End Sub
Private Sub wbStatus_DocumentComplete(index As Integer, ByVal pDisp As Object, URL As Variant)
    Dim webdocT As HTMLDocument
    
    'the current document/website was loaded
    'if it's the currently selected tab...
    If GetTabFromBrowser(index) = CurrentActiveServer.Tabs.SelectedItem.index Then
        '...we will have to hide the loading screen
        tmrRefreshSoon.Enabled = True
    '(if it's not the selected tab, the loading screen is going to be hidden when the user
    ' switches to that tab)
    End If
        
    'xNode tags code ;)
    Select Case index
        Case WebBrowserIndex_Chan, WebBrowserIndex_DCC, WebBrowserIndex_Priv, WebBrowserIndex_Loading, 0
        Case Else
            'web site
            
            If UBound(webdocWebSite) < index Then
                ReDim Preserve webdocWebSite(index)
            End If
            Set webdocWebSite(index) = wbStatus(index).Document
            Set webdocT = webdocWebSite(index)
            On Error GoTo Invalid_DOM_Document
            xNode webdocT
            
            CurrentActiveServer.Tabs.Tabs(GetTabFromBrowser(index)).Image = 4
    End Select
    
Invalid_DOM_Document:
End Sub
Private Sub wbStatus_GotFocus(index As Integer)
    If fraWebTab.Visible Then
        fraWebTab.Visible = False
    End If
End Sub
Private Sub wbStatus_NavigateComplete2(index As Integer, ByVal pDisp As Object, URL As Variant)
    Dim webdocT As HTMLDocument
    'Exit Sub
    Select Case index
        Case WebBrowserIndex_Chan
            Set webdocChanFrameSet = wbStatus(index).Document
            On Error Resume Next 'we'll reload them later
            Set webdocChanMain = webdocChanFrameSet.parentWindow.frames(1).Document
            Set webdocChanNicklist = webdocChanFrameSet.parentWindow.frames(2).Document
            Set webdocChanTopic = webdocChanFrameSet.parentWindow.frames(0).Document
        Case WebBrowserIndex_DCC
            Set webdocDCCs = wbStatus(index).Document
        Case WebBrowserIndex_Priv
            Set webdocPrivates = wbStatus(index).Document
        Case WebBrowserIndex_Loading, 0
            Set webdocT = wbStatus(index).Document
            xNodeTag webdocT, "lang_loading", Language(220)
        Case Else
            'web site
            If UBound(webdocWebSite) < index Then
                ReDim Preserve webdocWebSite(index)
            End If
            Set webdocWebSite(index) = wbStatus(index).Document
            Set webdocT = webdocWebSite(index)
            With webdocT.All
                If Not .Item("xnode_version") Is Nothing Then
                    .Item("xnode_version").innerText = App.Major & "." & App.Minor
                End If
                If Not .Item("xnode_major") Is Nothing Then
                    .Item("xnode_major").innerText = App.Major
                End If
                If Not .Item("xnode_minor") Is Nothing Then
                    .Item("xnode_minor").innerText = App.Minor
                End If
            End With
            SaveSession
    End Select
    If LCase(URL) = LCase(App.Path & "\data\html\imports\error.html") Then
        Set webdocT = webdocWebSite(index)
        xNodeTag webdocT, "lang_error", Language(64)
        xNodeTag webdocT, "lang_err_description", Language(218)
        xNodeTag webdocT, "lang_be_online", Language(219)
        xNodeTag webdocT, "lang_back", Language(276)
    End If
End Sub
Private Sub wbStatus_NavigateError(index As Integer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    Dim webdocError As HTMLDocument
        
    If StatusCode = 200 Then
        Exit Sub
    End If
    If Options.HTMLError = False Then
        'dont use Not here as vbNull is a possible value of Options.HTMLError
        Exit Sub
    End If
    'there was a navigation error
    'create an error trap, as the webBrowser object may fail to navigate
    On Error GoTo refresh_soon
    'navigate to node error page
    wbStatus(index).Navigate2 App.Path & "/data/html/imports/error.html"
    Set webdocError = wbStatus(index).Document
    wbStatus_TitleChange index, Language(535)
    Exit Sub
refresh_soon:
    'the webbrowser object couldn't navigate; we'll need to  do it soon.
    tmrRefreshSoon.Enabled = True
End Sub
Private Sub wbStatus_NewWindow2(index As Integer, ppDisp As Object, Cancel As Boolean)
    'instead of opening in new window, open in new tab ;-)
    CurrentActiveServer.preExecute "/browse about:blank"
    Set ppDisp = wbStatus(CurrentActiveServer.TabInfo(CurrentActiveServer.Tabs.SelectedItem.index)).Object
End Sub
Private Sub wbStatus_ProgressChange(index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
    Dim intTabArrayIndex As Integer
    Dim intServerIndex As Integer
    Dim sgnProgress As Single
    
    'the page being loaded has loaded some procent
    'if it's the selected one...
    If GetTabFromBrowser(index) = CurrentActiveServer.Tabs.SelectedItem.index Then
        '...we may need to remove the loading screen soon
        tmrRefreshSoon.Enabled = False
        tmrRefreshSoon.Enabled = True
    End If
    
    intTabArrayIndex = GetTabFromBrowser(index)
    intServerIndex = GetServerFromBrowser(index)
    
    If intTabArrayIndex = -1 Then
        Exit Sub
    Else
        If ActiveServers(intServerIndex).TabType(intTabArrayIndex) = TabType_WebSite Then
            If ProgressMax = 0 Then
                ActiveServers(intServerIndex).Tabs.Tabs(intTabArrayIndex).Image = 4
            Else
                sgnProgress = (Progress / ProgressMax)
                If sgnProgress > 1 Then
                    ActiveServers(intServerIndex).Tabs.Tabs(intTabArrayIndex).Image = 4
                Else
                    ActiveServers(intServerIndex).Tabs.Tabs(intTabArrayIndex).Image = sgnProgress * 13 + 7
                End If
           End If
        End If
    End If
End Sub
Private Sub wbStatus_SetSecureLockIcon(index As Integer, ByVal SecureLockIcon As Long)
    Dim strStatusText As String
    
    Select Case SecureLockIcon
        Case secureLockIconUnsecure
            strStatusText = vbNullString
        Case secureLockIconMixed
            strStatusText = Language(729)
        Case secureLockIconUnknownBits
            strStatusText = Language(730)
        Case secureLockIcon40Bit
            strStatusText = Language(725)
        Case secureLockIcon56Bit
            strStatusText = Language(726)
        Case secureLockIconFortezza
            strStatusText = Language(728)
        Case secureLockIcon128Bit
            strStatusText = Language(727)
    End Select
    sbar.Panels(2).Text = strStatusText
End Sub
Private Sub wbStatus_StatusTextChange(index As Integer, ByVal Text As String)
    Dim intServerIndex As Integer
    Dim TheServer As clsActiveServer
    
    Select Case index
        Case WebBrowserIndex_Loading
        Case WebBrowserIndex_Chan
        Case WebBrowserIndex_DCC
        Case WebBrowserIndex_Priv
        Case 0
        Case Else
            intServerIndex = GetServerFromBrowser(index)
            Set TheServer = ActiveServers(intServerIndex)
            
            If boolBuildingPrimary Then
                Exit Sub
            End If
            
            On Error Resume Next
            If TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) = index And TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_WebSite Then
                On Error Resume Next
                AddNews Replace(Text, "&", "&")
            End If
    End Select
End Sub
Private Sub wbStatus_TitleChange(index As Integer, ByVal Text As String)
    'the title of a website changes.
    'update the caption of its tab
    Dim intTabIndex As Integer
    Dim i As Integer
    Dim nodetype As String
    Dim strwebnum As Integer
    Dim intServerIndex As Integer
    
    Select Case index
        Case WebBrowserIndex_Loading
        Case WebBrowserIndex_Chan
        Case WebBrowserIndex_DCC
        Case WebBrowserIndex_Priv
        Case 0
        Case Else
            If index = WebBrowserIndex_Loading Or index < WebBrowserIndex_Loading + 1 Then
                Exit Sub
            End If
            
            intServerIndex = GetServerFromBrowser(index)
            For i = 1 To tvConnections.Nodes.Count
                If tvConnections.Nodes.Item(i).Key = "w_" & intServerIndex & "_" & index Then
                    tvConnections.Nodes.Item(i).Text = Text
                    Exit For
                End If
            Next i
            
            If Not ActiveServers(intServerIndex) Is CurrentActiveServer Then
                ActiveServers(intServerIndex).Tabs.Tabs.Item(GetTabFromBrowser(index)).Caption = Text
                Exit Sub
            End If
            
            intTabIndex = GetTabFromBrowser(index)
            If CurrentActiveServer.Tabs.Tabs.Item(intTabIndex).Caption <> Text Then
                If LCase$(wbStatus(index).LocationURL) <> "file:///" & LCase$(Replace(App.Path & "/temp/error.html", "\", "/")) Then
                    'get the tab index, set the caption,
                    CurrentActiveServer.Tabs.Tabs.Item(intTabIndex).Caption = Text
                    'and update.
                    UpdateTabsBar
                End If
            End If
    End Select
End Sub
Private Sub webdocCurrentIRCWindow_onmousedown()
    Dim bCurrentTabType As Byte
    Dim strNet As String
    Dim strFile As String
    
    If webdocCurrentIRCWindow.parentWindow.Event.Button = 2 Then
        'right-click
        bCurrentTabType = CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index)
        If bCurrentTabType = TabType_Channel Then
            If FS.FileExists(GetLogFile(CurrentActiveServer.Tabs.SelectedItem.Caption, CurrentActiveServer)) Then
                'logs exist; display menu item
                mnuViewLogs.Enabled = True
            Else
                'logs for the selected
                'IRC window do not exist
                mnuViewLogs.Enabled = False
            End If
        ElseIf bCurrentTabType = TabType_Private Then
            CurrentNick = CurrentActiveServer.Tabs.SelectedItem.Caption
            
            mnuKick.Enabled = False
            mnuMode.Enabled = False
            mnuWhisper.Enabled = False
            mnuNickClear.Enabled = Len(webdocCurrentIRCWindow.All("mainText").innerText)
            mnuNickViewLogs.Enabled = FS.FileExists(GetLogFile(CurrentNick, CurrentActiveServer))
            
            'display popup menu
            Me.PopupMenu mnuNickPop
        
            Exit Sub
        Else
            'not a channel or private window
            mnuViewLogs.Enabled = False
        End If

        mnuClear.Enabled = Len(webdocCurrentIRCWindow.All.tags("div").Item(0).innerText)
        mnuChanProperties.Enabled = bCurrentTabType = TabType_Channel
        mnuChanModes.Enabled = bCurrentTabType = TabType_Channel
        PopupMenu mnuIRCPop
    End If
End Sub
Private Sub webdocCurrentIRCWindow_onmouseup()
    If Options.SelectionCopy Then
        Clipboard.Clear
        On Error GoTo Failed_To_Create_Range
        Clipboard.SetText webdocCurrentIRCWindow.selection.createRange.Text
        If Options.SelectionClear Then
            webdocCurrentIRCWindow.selection.empty
        End If
    End If
Failed_To_Create_Range:
End Sub
Private Function webdocFullScreen_ondblclick() As Boolean
    MaxMode = False
    frmFullScreen.Hide
    Set frmFullScreen = Nothing
End Function
Private Sub webdocFullScreen_onkeydown()
    If e_KeyCode = 0 Then
        e_KeyCode = webdocFullScreen.parentWindow.Event.KeyCode
    End If
    If e_KeyCode = vbKeyEscape Or e_KeyCode = vbKeyF10 Then
        MaxMode = False
        frmFullScreen.Hide
        Set frmFullScreen = Nothing
    End If
    e_KeyCode = 0
End Sub
Private Sub webdocFullScreenChanMain_onkeydown()
    e_KeyCode = webdocFullScreenChanMain.parentWindow.Event.KeyCode
    webdocFullScreen_onkeydown
End Sub
Private Sub webdocFullScreenChanNicks_onkeydown()
    e_KeyCode = webdocFullScreenChanNicks.parentWindow.Event.KeyCode
    webdocFullScreen_onkeydown
End Sub
Private Sub wsDCC_ConnectionClosed(ByVal index As Integer) 'RCV: YES
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, True))
    
    'closing the receiving winsock
    wsDCC.Item(index).Close
    
    TheServer.DCCFile_Progress(TheServer.GetDCCFileIndexFromWsIndex(index, True), True) = 100
    
    If TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_DCCFile And TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) = 0 Then
        buildStatus
    End If
End Sub
Private Sub wsDCC_Connect(ByVal index As Integer) 'RCV: YES
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, True))
    
    'winsock for receiving a file is connected, save the "time"
    TheServer.DCCFile_StartTime(TheServer.GetDCCFileIndexFromWsIndex(index, True), True) = GetTickCount
End Sub
Private Sub wsDCC_DataArrival(ByVal index As Integer, ByVal bytesTotal As Long) 'RCV: YES
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, True))
    
    TheServer.DccRcv_DataReceived index
End Sub
Private Sub wsDCC_Error(ByVal index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean) 'RCV: YES
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, True))
    TheServer.DccRcv_Error index
    
    AddStatus Language(892) & vbNewLine & WinSockErrorIDToLangString(Number) & vbNewLine, TheServer
End Sub

Private Sub wsDCCChat_ConnectionClosed(ByVal index As Integer)
    'the DCC Chat winsock connection has been closed
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromDCCChatWsIndex(index))
    
    TheServer.DCCChats_Disconnect TheServer.Tabs.Tabs.Item(TheServer.GetDCCChatsTabIndexFromWsIndex(index)).Caption
    
    'TheServer.DCCChats_Disconnect (TheServer.DCCChats_UserName(TheServer.GetDCCChatsIndexFromWsIndex(Index)))
End Sub

Private Sub wsDCCChat_ConnectionRequest(ByVal index As Integer, ByVal requestID As Long)
    'the DCC Chat winsock that was waiting for a reply got one
    'all the connection and display it
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromDCCChatWsIndex(index))
    
    TheServer.DccChat_ConnectionRequest index, requestID
End Sub
Private Sub wsDCCChat_DataArrival(ByVal index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromDCCChatWsIndex(index))
    
    'get the info that was received over the winsock
    wsDCCChat.Item(index).GetData strData
    
    TheServer.DccChat_DataReceived index, strData
    
End Sub
Private Sub wsDCCSend_Connect(ByVal index As Integer) 'RCV: NO
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, False))
    
    TheServer.DccSend_Connected index
End Sub
Private Sub wsDCCSend_ConnectionRequest(ByVal index As Integer, ByVal requestID As Long) 'RCV: NO
    Dim TheServer As clsActiveServer
    
    wsDCCSend.Item(index).Close
    wsDCCSend.Item(index).accept requestID 'allow the connection
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, False))
    TheServer.DccSend_ConnectionRequest index, requestID
End Sub
Public Sub wsDCCSend_DataArrival(ByVal index As Integer, ByVal bytesTotal As Long) 'RCV: NO
    Dim strData As String
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(GetServerFromWsDCCIndex(index, False))
    
    On Error Resume Next
    wsDCCSend.Item(index).GetData strData
    
    If Err Then
        DB.X "Warning: DCCSend failed to receive check bytes from remote host."
    End If
    
    TheServer.DccSend_DataReceived index, strData
End Sub
Private Sub wsIdentD_ConnectionRequest(ByVal index As Integer, ByVal requestID As Long)
    'This must be IdentD #0
   
    Dim strRemoteHost As String
    
    With wsIdentD.Item(wsIdentD.LoadNew)
        .Close
        .accept requestID
        wsIdentD.Item(0).Close
        On Error GoTo CouldNotInitIDENT
        wsIdentD.Item(0).Listen
        strRemoteHost = .RemoteHost
        If Len(strRemoteHost) <= 0 Then
            strRemoteHost = .RemoteHostIP
        End If
    End With
    AddStatus EVENT_PREFIX & Replace(Language(396), "%1", strRemoteHost) & EVENT_SUFFIX & vbNewLine, CurrentActiveServer
    AddNews Language(686)
CouldNotInitIDENT:
End Sub
Private Sub wsIdentD_DataArrival(ByVal index As Integer, ByVal bytesTotal As Long)
    Dim strTemp As String
    Dim strData() As String
    Static bData As String
    
    On Error GoTo Invalid_Data_RCV
    wsIdentD.Item(index).GetData strTemp, , bytesTotal
    bData = bData & strTemp
ReInterprent:
    If InStr(1, bData, vbNewLine) > 0 Then
        strTemp = Left$(bData, InStr(1, bData, vbNewLine) - 1)
        bData = Right$(bData, Len(bData) - Len(strTemp) - 2)
        strData = Split(strTemp, " , ")
        wsIdentD.Item(index).SendData strData(0) & ", " & strData(1) & " : USERID : UNIX : " & frmOptions.txtNickname.Text & vbNewLine
    End If
    If LenB(bData) > 0 Then
        GoTo ReInterprent
    End If
Invalid_Data_RCV:
End Sub
Private Sub wsIdentD_Error(ByVal index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wsIdentD.Item(index).Close
    If index = 0 Then
        wsIdentD.Item(index).Listen
    Else
        'Unload wsIdentD.Item(Index)
    End If
End Sub
Private Sub wsNDC_Connect(ByVal index As Integer)
    ReDim Preserve NDCConnections(index)
    
    NDCSendIntro index
End Sub
Private Sub wsNDC_ConnectionRequest(ByVal index As Integer, ByVal requestID As Long)
    Dim intIndex As Integer
    
    With wsNDC.Item(xLet(intIndex, wsNDC.LoadNew))
        ReDim Preserve NDCConnections(intIndex)
        .Close
        .accept requestID
        wsNDC.Item(0).Close
        wsNDC.Item(0).Listen
    End With
End Sub
Private Sub wsNDC_DataArrival(ByVal index As Integer, ByVal bytesTotal As Long)
    'we get some NDC data
    Dim iData As String
    Dim bData As String
    Dim intDot1Pos As Integer
    Dim intDot2Pos As Integer
    Dim intTemp As Integer
    Dim bTemp As Byte
    Dim strTemp As String
    Dim intPrevDataLen As Integer
    Dim intBadCount As Integer
    
    'we may not be connected; error handle this
    On Local Error GoTo Invalid_Data
    'get the incoming data
    wsNDC.Item(index).GetData bData, , bytesTotal
    'add it to the current data; this is done in order
    'not to avoid fragmented data
    NDCConnections(index).aData = NDCConnections(index).aData & bData
ReInterprent:
    'if we haven't recieved an intro pack yet
    If Not NDCConnections(index).IntroPackRecieved Then
        'intro pack not recieved yet
        'this must be an intro pack
        'as we can't recieve anything
        'else before this
        'get the position of the first seperator character :
        intDot1Pos = InStr(1, NDCConnections(index).aData, ":")
        'Position of 2nd :
        intDot2Pos = InStr(intDot1Pos + 1, NDCConnections(index).aData, ":")
        If intDot2Pos > 0 Then
                            '    :  + TimeZone + TCP
            'if this is true, we have recieved the hole intro pack
            'if it's false, we should wait for more data
            If Len(NDCConnections(index).aData) >= intDot2Pos + 1 + 2 Then
                'full intro pack recieved
                'send an i-pack back
                NDCIntro Left$(NDCConnections(index).aData, intDot2Pos + 1 + 2), index
                'remove the i-pack data we just recieved & analyzed
                NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - intDot2Pos - 5)
            End If
        End If
        
    'we have recieved an intro pack. This must be a normal NDC message
    Else
        'if the data we recieved is "t" it is a typing pack, one byte
        If Left$(NDCConnections(index).aData, 1) = "t" Then
            'typing msg
            'remove the data we recieved
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 1)
            'the remote user is typing
            NDCConnections(index).Typing = True
            'he/she started typing now
            NDCConnections(index).TypingTime = GetTickCount
            'get the index of the private tab
            intTemp = CurrentActiveServer.GetTabFromNick(NDCConnections(index).strNicknameA)
            'if it's the selected one
            If CurrentActiveServer.Tabs.SelectedItem.index = intTemp And CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
                'display
                'Nickname is typing a message...
                'at the status bar
                AddNews Replace(Language(318), "%1", NDCConnections(index).strNicknameA) & "..."
            'if it's not, it will be displayed when we change tabs
            End If
        
        'status pack, two bytes
        ElseIf Left$(NDCConnections(index).aData, 1) = "s" And Len(NDCConnections(index).aData) >= 2 Then
            'status pack
            'get the status ID from the data we recieved
            bTemp = AscW(Mid$(NDCConnections(index).aData, 2, 1))
            'remove the data we recieved
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 2)
            'check if it's a valid status ID
            If bTemp <= 10 Then
                'valid status
                'store it
                NDCConnections(index).RemoteStatus = bTemp
                'get the index of the private tab
                intTemp = CurrentActiveServer.GetTabFromNick(NDCConnections(index).strNicknameA)
                'if it's the selected one
                If CurrentActiveServer.Tabs.SelectedItem.index = intTemp And CurrentActiveServer.TabType(CurrentActiveServer.Tabs.SelectedItem.index) = TabType_Private Then
                    'if the NDC's private is focused
                    'show the text for the status
                    'online status?
                    If bTemp = Status_Online Then
                        'don't show anything
                        sbar.Panels(2).Text = vbNullString
                    'other status
                    Else
                        'show the status
                        sbar.Panels(2).Text = Replace( _
                                              Replace( _
                                                Language(421), "%1", NDCConnections(index).strNicknameA), _
                                                               "%2", Language(bTemp + 397))
                    End If
                End If
            End If
        
        'M SYN
        ElseIf Left$(NDCConnections(index).aData, 2) = "M?" And Len(NDCConnections(index).aData) >= 4 Then
            'program session start request
            'get program ID
            Select Case Mid$(NDCConnections(index).aData, 3, 2)
                'reserved
                Case ChrW$(0) & ChrW$(0)
                
                'Microsoft NetMeeting
                Case ChrW$(1) & ChrW$(9)
                    If MsgBox(Replace(Replace(Language(787), "%1", NDCConnections(index).strNicknameA), "%2", Language(783)), vbQuestion Or vbYesNo, Language(788)) = vbYes Then
                        'send M SYN/ACK
                        wsNDC.Item(index).SendData "M+" & Mid$(NDCConnections(index).aData, 3, 2)
                        NDCConnections(index).MMNetMeetingStatus = 2 'accepted and waiting for ACK
                        'StartNetMeetingSession False, wsNDC.Item(Index).RemoteHostIP
                    Else
                        'send M RST
                        wsNDC.Item(index).SendData "M-" & Mid$(NDCConnections(index).aData, 3, 2)
                        NDCConnections(index).MMNetMeetingStatus = 0 'refused
                    End If
                    
                Case Else
                    AddStatus Language(786), NDCConnections(index).ActiveServer
            End Select
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 4)
        
        'M RST
        ElseIf Left$(NDCConnections(index).aData, 2) = "M-" And Len(NDCConnections(index).aData) >= 4 Then
            'program session start request
            'get program ID
            Select Case Mid$(NDCConnections(index).aData, 3, 2)
                'reserved
                Case ChrW$(0) & ChrW$(0)
                
                'Microsoft NetMeeting
                Case ChrW$(1) & ChrW$(9)
                    'check if we actually DID a request before we got that reply
                    If NDCConnections(index).MMNetMeetingStatus = 1 Then
                        'yep, display "refused" message
                        AddStatus Replace(Replace(Language(790), "%1", NDCConnections(index).strNicknameA), "%2", Language(783)) & vbNewLine, NDCConnections(index).ActiveServer, NDCConnections(index).ActiveServer.GetChanID(NDCConnections(index).strNicknameA)
                        NDCConnections(index).MMNetMeetingStatus = 0 'restore
                    Else
                        'we didn't do a request, but we
                        'got an RST message
                        'don't do anything
                    End If
                    
                Case Else
                    AddStatus Language(786), NDCConnections(index).ActiveServer
            End Select
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 4)
        
        'M SYN/ACK
        ElseIf Left$(NDCConnections(index).aData, 2) = "M+" And Len(NDCConnections(index).aData) >= 4 Then
            'program session accept pack
            'get program ID
            Select Case Mid$(NDCConnections(index).aData, 3, 2)
                'reserved
                Case ChrW$(0) & ChrW$(0)
                
                'Microsoft NetMeeting
                Case ChrW$(1) & ChrW$(9)
                    'check if we actually DID a request before we got that reply
                    If NDCConnections(index).MMNetMeetingStatus = 1 Then
                        'yep, display "accepted" message
                        AddStatus Replace(Replace(Language(791), "%1", NDCConnections(index).strNicknameA), "%2", Language(783)) & vbNewLine, NDCConnections(index).ActiveServer, NDCConnections(index).ActiveServer.GetChanID(NDCConnections(index).strNicknameA)
                        NDCConnections(index).MMNetMeetingStatus = 3 'ready
                        AddStatus Replace(Language(792), "%2", Language(783)) & vbNewLine, NDCConnections(index).ActiveServer, NDCConnections(index).ActiveServer.GetChanID(NDCConnections(index).strNicknameA)
                        'start NetMeeting and wait for connection
                        StartNetMeetingSession True
                    
                        'and send ACK
                        wsNDC.Item(index).SendData "M*" & Mid$(NDCConnections(index).aData, 3, 2)
                    Else
                        'we didn't do a request, but we
                        'got a SYN/ACK message
                        'don't do anything
                    End If
                    
                Case Else
                    AddStatus Language(786), NDCConnections(index).ActiveServer
            End Select
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 4)
        
        'M ACK
        ElseIf Left$(NDCConnections(index).aData, 2) = "M*" And Len(NDCConnections(index).aData) >= 4 Then
            'program session ACK pack
            'get program ID
            Select Case Mid$(NDCConnections(index).aData, 3, 2)
                'reserved
                Case ChrW$(0) & ChrW$(0)
                
                'Microsoft NetMeeting
                Case ChrW$(1) & ChrW$(9)
                    'check if we actually accepted that request
                    If NDCConnections(index).MMNetMeetingStatus = 2 Then
                        'yep, we did
                        NDCConnections(index).MMNetMeetingStatus = 3 'ready
                        AddStatus Language(792) & vbNewLine, NDCConnections(index).ActiveServer, NDCConnections(index).ActiveServer.GetChanID(NDCConnections(index).strNicknameA)
                        'start NetMeeting and connect to remote user
                        StartNetMeetingSession False, wsNDC.Item(index).RemoteHostIP
                    
                        'and send ACK
                        wsNDC.Item(index).SendData "M*" & Mid$(NDCConnections(index).aData, 3, 2)
                    Else
                        'we didn't do a request, but we
                        'got a SYN/ACK message
                        'don't do anything
                    End If
                    
                Case Else
                    AddStatus Language(786), NDCConnections(index).ActiveServer
            End Select
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 4)
        
        'audio conversation request
        ElseIf Left$(NDCConnections(index).aData, 1) = "a" And Len(NDCConnections(index).aData) >= 3 Then
            'addStatus EVENT_PREFIX & _
                        Replace(Replace(Replace( _
                                        Language(427), "%1", NDCConnections(Index).strNicknameA), _
                                                       "%2", "<a href=""NodeScript:/audio-accept"">" & Language(313) & "</a>"), _
                                                       "%3", "<a href=""NodeScript:/audio-decline"">" & Language(428) & "</a>") & _
                                        EVENT_SUFFIX & vbnewline, GetChanID(NDCConnections(Index).strNicknameA)
            'ask the user if he/she wants to accept the conversation
            If MsgBox(Replace(Language(431), "%1", NDCConnections(index).strNicknameA), vbQuestion Or vbYesNo, Language(432)) = vbYes Then
                'accepted.
                'get the TCP port for audio conversations
                'NDCConnections(Index).AudioTCP = UpperLowerToInt(Mid$(NDCConnections(Index).aData, 2, 1), Mid$(NDCConnections(Index).aData, 3, 1))
                'we are now connected
                NDCConnections(index).AudioRequested = False
                NDCConnections(index).AudioConnected = True
                'let the remote user know that we accepted
                wsNDC.Item(index).SendData "AA"
                sbar.Panels(2).Text = Replace(Language(429), "%1", NDCConnections(index).strNicknameA)
            Else
                'refused
                'we are not connected
                NDCConnections(index).AudioRequested = False
                NDCConnections(index).AudioConnected = False
                'et the remote user know that we rejected the request
                wsNDC.Item(index).SendData "AD"
            End If
            'remove the data we recieved
            NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 1)
        
        'audio/accepted
'        ElseIf Left$(NDCConnections(Index).aData, 2) = "AA" Then
'            If NDCConnections(Index).AudioRequested Then
'                AddStatus EVENT_PREFIX & Replace(Language(429), "%1", NDCConnections(Index).strNicknameA) & _
'                                         EVENT_SUFFIX & vbnewline, GetChanID(NDCConnections(Index).strNicknameA)
'                sbar.Panels(2).Text = Replace(Language(429), "%1", NDCConnections(Index).strNicknameA)
'                NDCConnections(Index).AudioConnected = True
'                NDCConnections(Index).AudioRequested = False
'            End If
'            'remove the data we just recieved
'            NDCConnections(Index).aData = Right$(NDCConnections(Index).aData, Len(NDCConnections(Index).aData) - 2)
        
        'audio/declined
'        ElseIf Strings.Left$(NDCConnections(Index).aData, 2) = "AD" Then
'            If NDCConnections(Index).AudioRequested Then
'                AddStatus EVENT_PREFIX & Replace(Language(430), "%1", NDCConnections(Index).strNicknameA) & _
'                                         EVENT_SUFFIX & vbnewline, GetChanID(NDCConnections(Index).strNicknameA)
'                NDCConnections(Index).AudioConnected = False
'                NDCConnections(Index).AudioRequested = False
'            End If
'            'remove the data we just recieved
'            NDCConnections(Index).aData = Right$(NDCConnections(Index).aData, Len(NDCConnections(Index).aData) - 2)
        
        'new nick pack
        ElseIf Strings.Left$(NDCConnections(index).aData, 1) = "n" And InStr(1, NDCConnections(index).aData, ":") > 1 Then
            strTemp = Mid$(NDCConnections(index).aData, 2, InStr(1, NDCConnections(index).aData, ":") - 1)
            Set FS = New FileSystemObject
            If FS.FileExists(App.Path & "/conf/avatars/" & NDCConnections(index).strNicknameA & ".jpg") Then
                FS.MoveFile App.Path & "/conf/avatars/" & NDCConnections(index).strNicknameA & ".jpg", App.Path & "/conf/avatars/" & strTemp & ".jpg"
            End If
            NDCConnections(index).strNicknameA = strTemp
            
        'audio send request(audata) - new audio data available
        ElseIf Strings.Left$(NDCConnections(index).aData, 1) = "[" And Len(NDCConnections(index).aData) >= 5 Then
                intTemp = UBound(AllowedHidden) + 1
                ReDim Preserve AllowedHidden(intTemp)
                AllowedHidden(intTemp).AllowedFileName = Strings.Mid$(NDCConnections(index).aData, 2, 4)
                AllowedHidden(intTemp).AllowedIP = wsNDC.Item(index).RemoteHostIP
                AllowedHidden(intTemp).AllowedNickname = NDCConnections(index).strNicknameA
                AllowedHidden(intTemp).NDCConnectNUM = index
                AllowedHidden(intTemp).WriteFileName = NDCConnections(index).strNicknameA '& Strings.Mid$(NDCConnections(Index).aData, 2, 4)
                '2 - audio connections
                AllowedHidden(intTemp).EventID = 2
                'remove the data we just recieved
                NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 5)
        
        'avatar send request(raverix)
        ElseIf Strings.Left$(NDCConnections(index).aData, 1) = "v" And Len(NDCConnections(index).aData) > 1 Then
            If Mid$(NDCConnections(index).aData, 2, 1) = "x" Then
                'no avatar
                'delete user's avatar
                NDCConnections(index).UserAvatar = vbNullString
            
                'remove the data we just recieved
                NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 2)
                If FS.FileExists(App.Path & "/conf/avatars" & NDCConnections(index).strNicknameA & ".jpg") Then
                    Kill App.Path & "/conf/avatars" & NDCConnections(index).strNicknameA & ".jpg"
                End If
            Else
                'check to see if we have recieved the full file name
                'or some bytes are still pending
                If Len(NDCConnections(index).aData) >= 5 Then
                    intTemp = UBound(AllowedHidden) + 1
                    ReDim Preserve AllowedHidden(intTemp)
                    AllowedHidden(intTemp).AllowedFileName = Strings.Mid$(NDCConnections(index).aData, 2, 4)
                    AllowedHidden(intTemp).AllowedIP = wsNDC.Item(index).RemoteHostIP
                    AllowedHidden(intTemp).AllowedNickname = NDCConnections(index).strNicknameA
                    AllowedHidden(intTemp).NDCConnectNUM = index
                    AllowedHidden(intTemp).WriteFileName = NDCConnections(index).strNicknameA '& Strings.Mid$(NDCConnections(Index).aData, 2, 4)
                    '1 - avatar
                    AllowedHidden(intTemp).EventID = 1
                    'remove the data we just recieved
                    NDCConnections(index).aData = Right$(NDCConnections(index).aData, Len(NDCConnections(index).aData) - 5)
                Else
                    'we need some more bytes to be recieved
                    '(do not reinterprent)
                    Exit Sub
                End If
            End If
        Else
            'we need some more bytes to be recieved
            'or the pack is invalid
            'do not reinterprent
            '(so that we don't get into an endless loop)
            Exit Sub
        End If
    End If
    'if there's more data to analyze...
    If Len(NDCConnections(index).aData) > 0 Then
        'go for it
        intPrevDataLen = Len(NDCConnections(index).aData)
        If intPrevDataLen <= Len(NDCConnections(index).aData) Then
            intBadCount = intBadCount + 1
            If intBadCount >= 10 Then
                'invalid NDC data
                AddStatus Language(476) & vbNewLine, NDCConnections(index).ActiveServer
                wsNDC.Item(index).Close
                Exit Sub
            End If
        End If
        GoTo ReInterprent
    End If
    Exit Sub
Invalid_Data:
End Sub
Private Sub NDCIntro(ByVal strPack As String, ByVal index As Integer)
    Dim intDot1Pos As Integer
    Dim intDot2Pos As Integer
    Dim introVersion As Integer '2 bytes
    Dim introVersionUpper As Byte
    Dim introVersionLower As Byte
    Dim strVersion As String * 2
    Dim introNicknameA As String
    Dim introNicknameB As String
    Dim introTimeZone As Byte
    Dim introTCPUpper As Byte
    Dim introTCPLower As Byte
    Dim introTCP As Integer
    Dim intTemp As Integer
    Dim i As Integer
    
    'we haven't recieved an intro pack yet
    intDot1Pos = Strings.InStr(1, strPack, ":")
    If intDot1Pos < 3 Or intDot1Pos > 260 Then
        'invalid position of the first :
        'nickname A is longer than 256 bytes
        'or shorter than 1 byte
        GoTo Invalid_Intro_Pack
    End If
    'Position of 2nd :
    intDot2Pos = Strings.InStr(intDot1Pos + 1, strPack, ":")
    If intDot2Pos < intDot1Pos + 1 Or intDot2Pos > intDot1Pos + 256 Then
        'invalid position of the second :
        'nickname B is longer than 256 bytes
        'or shorter than 1 byte
        GoTo Invalid_Intro_Pack
    End If
    strVersion = Strings.LeftB(strPack, 2)
    introVersionUpper = AscW(Strings.Left$(strVersion, 1))
    introVersionLower = AscW(Strings.Right$(strVersion, 1))
    introVersion = xBasic.UpperLowerToInt(introVersionUpper, introVersionLower) + 32767
    introNicknameA = Strings.Mid$(strPack, 3, intDot1Pos - 3)
    
    'now we have to look inside the current pending NDC connections
    'and find this one. We will only be able to find it if the connection
    'was *not* requested by us, but by the remote user
    'However, if the connection was requested by us, we
    'already know the ActiveServer of the user, and therefore
    'we don't need to check that array
    If NDCConnections(index).ActiveServer Is Nothing Then
        For i = 0 To UBound(PendingNDCConnectionRequests)
            If PendingNDCConnectionRequests(i).Nickname = introNicknameA Then
                Set NDCConnections(index).ActiveServer = PendingNDCConnectionRequests(i).ActiveServer
                'clear this pending connection; we won't need it any more
                PendingNDCConnectionRequests(i).Nickname = vbNullString
                Set PendingNDCConnectionRequests(i).ActiveServer = Nothing
                Exit For
            End If
        Next i
    End If
    
    'now we know the current ActiveServer :)
    
    introNicknameB = Strings.Mid$(strPack, intDot1Pos + 1, intDot2Pos - intDot1Pos - 1)
    introTimeZone = AscW(Mid$(strPack, intDot2Pos + 1, 1))
    introTCPUpper = AscW(Strings.Mid$(strPack, intDot2Pos + 3, 1))
    introTCPLower = AscW(Strings.Mid$(strPack, intDot2Pos + 2, 1))
    introTCP = UpperLowerToInt(introTCPUpper, introTCPLower) '+ 32768 in order to connect
    'check if timezone is OK.
    If introTimeZone > 23 Then
        GoTo Invalid_Intro_Pack
    End If
    'protocols compatibility check
    If introVersion > App.Major * 100 + App.Minor Then
        'the protocol version of the remote client is newer
        'than our.
        'Let the remote client decide whether a connection
        'is possible.
    ElseIf introVersion > 30 Then
        'versions greater than 30
        'are compatible with the current protocol
    Else
        'the protocols are not compatible
        GoTo Incompatible_Intro_Pack
    End If
    'everything OK, save info
    NDCConnections(index).bTimeZone = introTimeZone
    NDCConnections(index).intTCP = introTCP 'note: before trying to connect add 32768 to the port number
    NDCConnections(index).intVersion = introVersion
    NDCConnections(index).strNicknameA = introNicknameA
    NDCConnections(index).strNicknameB = introNicknameB
    NDCConnections(index).IntroPackRecieved = True
    If Not NDCConnections(index).IntroPackSent Then
        NDCSendIntro index
    End If
    'addStatus "RECIEVED NDC INFORMATION:" & vbnewline & _
              "Nickname: " & introNicknameA & vbnewline & _
              "(NicknameB): " & introNicknameB & vbnewline & _
              "Version: " & introVersion & " (compatible)" & vbnewline & _
              "TimeZone: " & introTimeZone & " (i.e. GMT " & IIf(Sgn(introTimeZone) - 11 > 0, "+", "") & introTimeZone - 11 & ":00)" & vbnewline
    intTemp = NDCConnections(index).ActiveServer.GetChanID(introNicknameA)
    If intTemp = 0 Then
        intTemp = NDCConnections(index).ActiveServer.GetStatusID
    End If
    AddStatus EVENT_PREFIX & "[" & Replace(Language(443), "%1", introNicknameA) & "]" & vbNewLine & EVENT_SUFFIX, NDCConnections(index).ActiveServer, intTemp
    sbar.Panels(2).Text = Replace(Language(443), "%1", introNicknameA)
    
    If NDCConnections(index).IntroPackSent Then
        'after we recieve the i-pack
        'we can send a raverix request
        NDCRaverixSendRequest index
    End If
    
    Exit Sub
Invalid_Intro_Pack:
Incompatible_Intro_Pack:
    wsNDC.Item(index).Close
End Sub
Private Sub NDCSendIntro(index As Integer)
    Dim strIntro As String
    Dim introVersionUpper As Byte
    Dim introVersionLower As Byte
    Dim introVersion As Integer
    Dim introNicknames As String
    Dim introTCP As String
    If NDCConnections(index).IntroPackSent = True And NDCConnections(index).IntroPackRecieved = True Then
        Exit Sub
    End If
    IntToUpperLower App.Major * 100 + App.Minor - 32768, introVersionUpper, introVersionLower
    If LenB(NDCConnections(index).ActiveServer.myNick) = 0 Then
        NDCConnections(index).ActiveServer.myNick = frmOptions.txtNickname
    End If
    introNicknames = NDCConnections(index).ActiveServer.myNick & ":x:"
    strIntro = ChrW$(introVersionUpper) & ChrW$(introVersionLower) & _
               introNicknames & ChrW$(GetTimeZoneBIAS) & _
               ChrW$(0) & ChrW$(0)
    wsNDC.Item(index).SendData strIntro
    NDCConnections(index).IntroPackSent = True
    
    NDCSendStatus index

    If NDCConnections(index).IntroPackRecieved Then
        'after we recieve the i-pack
        'we can send a raverix request
        NDCRaverixSendRequest index
    End If

End Sub
Public Sub NDCSendStatus(index As Integer)
    If wsNDC.Item(index).State = sckConnected Then
        wsNDC.Item(index).SendData "s" & ChrW$(NDCConnections(index).ActiveServer.MyStatus)
    End If
End Sub
Private Sub NDCRaverixSendRequest(index As Integer)
    If wsNDC.Item(index).State = sckConnected Then
        NDCRandomCurrent = NDCRandomCurrent + 1
        'send NDC avatar send request pack
        If FS.FileExists(App.Path & "/conf/myavatar.jpg") Then
            wsNDC.Item(index).SendData "v" & FixLeadingZero(NDCRandomCurrent, 4)
            NDCConnections(index).AvatarToSend = FixLeadingZero(NDCRandomCurrent, 4)
            NDCGlobalEventID = NDCGlobalEventID + 1
            ReDim AllowedHidden(UBound(AllowedHidden) + 1)
            AllowedHidden(UBound(AllowedHidden)).AllowedFileName = FixLeadingZero(NDCRandomCurrent, 4)
            AllowedHidden(UBound(AllowedHidden)).AllowedIP = wsNDC.Item(index).RemoteHostIP
            AllowedHidden(UBound(AllowedHidden)).AllowedNickname = NDCConnections(index).strNicknameA
            AllowedHidden(UBound(AllowedHidden)).EventID = 1
            FS.CopyFile App.Path & "\conf\myavatar.jpg", App.Path & "\temp\myavatar_" & FixLeadingZero(NDCRandomCurrent, 4) & ".jpg"
            CurrentNick = NDCConnections(index).strNicknameA
            
            'TO DO:
            'Direct Raverix Send, not DCC request via Server.
            
            NDCConnections(index).ActiveServer.DCCTransfer_SendFile CurrentNick, App.Path & "\temp\myavatar_" & FixLeadingZero(NDCRandomCurrent, 4) & ".jpg", True
        Else
            'no avatar
            wsNDC.Item(index).SendData "vx"
        End If
    End If
End Sub
Private Function NDCGenerateRandomID() As Integer
    '0...9999
    NDCGenerateRandomID = Int(Rnd * 10000)
End Function
'Private Function NDCAudioAreActiveConnections() As Boolean
'    Dim i As Integer
'    For i = 0 To UBound(NDCConnections)
'        If wsNDC.Item(i).State = sckConnected Then
'            If NDCConnections(i).AudioConnected Then
'                NDCAudioAreActiveConnections = True
'                Exit Function
'            End If
'        End If
'    Next i
'    NDCAudioAreActiveConnections = False
'End Function
Private Sub NDCAudioRecord()
    'NDC audio start recording
    Dim lpszReturnString As String * 256
    'open device; create a new file to save in
    mciSendString "open new type waveaudio alias ndcaudio", lpszReturnString, _
                    Len(lpszReturnString), 0
    'start recording to that file
    mciSendString "record ndcaudio", lpszReturnString, _
                    Len(lpszReturnString), 0
End Sub
Private Sub NDCAudioStopRecording()
    Dim lpszReturnString As String * 256
    'stop recording
    mciSendString "stop ndcaudio", lpszReturnString, _
                    Len(lpszReturnString), 0
    'save data
    mciSendString "save ndcaudio ndcaudio.wav", lpszReturnString, _
                    Len(lpszReturnString), 0
    'close audio file
    mciSendString "close ndcaudio", lpszReturnString, _
                    Len(lpszReturnString), 0
End Sub
Private Sub wsNDC_Error(ByVal index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Stop
End Sub

'Private Sub NDCAudioRegularOperations()
'    Dim boolActiveAudio As Boolean
'    Dim i As Integer
'    Static boolNeedToStop As Boolean
'    'check to see if there are any active audio conversations
'    'and store the result in boolActiveAudio
'    boolActiveAudio = NDCAudioAreActiveConnections
'    'if there are some active *recordings* (not necessary conversations)
'    If boolNeedToStop Then
'        'stop them
'        NDCAudioStopRecording
'        'we have stopped and saved some data
'        'if the conversations which caused the
'        'recording are still active
'        If boolActiveAudio Then
'            'we need
'            'to send the recorded data
'            'check to see which conversations are still active
'            For i = 0 To UBound(NDCConnections)
'                If wsNDC.Item(i).State = sckConnected Then
'                    If NDCConnections(i).AudioConnected Then
'                        NDCAudioSendData i
'                    End If
'                End If
'            Next i
'        End If
'    End If
'    'if there are active NDC audio conversations
'    If boolActiveAudio Then
'        'start recording
'        NDCAudioRecord
'        'we will need to stop this recording later
'        boolNeedToStop = True
'    'there are no active NDC audio conversations
'    Else
'        'we won't need to stop later
'        boolNeedToStop = False
'    End If
'End Sub
'Private Sub NDCAudioSendData(ByVal ConnectionID As Integer)
'    'TO DO:
'    'Implement Stream Audio Communication
'
'    If wsNDC.Item(ConnectionID).State = sckConnected Then
'        NDCRandomCurrent = NDCRandomCurrent + 1
'        'send NDC audata send request pack(new data available)
'        wsNDC.Item(ConnectionID).SendData "[" & FixLeadingZero(NDCRandomCurrent, 4)
'        NDCConnections(ConnectionID).AudioSentTime = GetTickCount
'        ReDim AllowedHidden(UBound(AllowedHidden) + 1)
'        AllowedHidden(UBound(AllowedHidden)).AllowedFileName = FixLeadingZero(NDCRandomCurrent, 4)
'        AllowedHidden(UBound(AllowedHidden)).AllowedIP = wsNDC.Item(ConnectionID).RemoteHostIP
'        AllowedHidden(UBound(AllowedHidden)).AllowedNickname = NDCConnections(ConnectionID).strNicknameA
'        AllowedHidden(UBound(AllowedHidden)).EventID = 2
'        FS.CopyFile App.Path & "\conf\myavatar.jpg", App.Path & "\conf\" & FixLeadingZero(NDCRandomCurrent, 4) & ".jpg"
'        CurrentNick = NDCConnections(ConnectionID).strNicknameA
'        cdfile.FileName = App.Path & "\temp\" & FixLeadingZero(NDCRandomCurrent, 4) & ".jpg"
'        wsNDC.Item(ConnectionID).SendData "[" & FixLeadingZero(NDCRandomCurrent, 4)
'        ActiveServers(wsNDC.Item(ConnectionID).Tag).DCCTransfer_SendFile CurrentNick, cdfile.FileName, True
'    End If
'End Sub
Private Sub xpBalloon_BalloonClicked()
    Select Case ThisTip
        Case WelcomeToNode
            If frmOptions.Visible Then
                frmOptions.Hide
            End If
            CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=changes"
        Case Connected
            If frmOptions.Visible Then
                frmOptions.Hide
            End If
            CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=help"
        Case Kicked
            CurrentActiveServer.preExecute "/join " & strLastBalloonInfo
        Case Invitation
            CurrentActiveServer.preExecute "/join " & strLastBalloonInfo
        Case NickInUse
            Set frmOptions.lvList.SelectedItem = frmOptions.lvList.ListItems.Item(1) 'user
            frmOptions.Show
        Case LangChange
            If frmOptions.Visible Then
                frmOptions.Hide
            End If
            CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=translate"
        Case SkinChange
            If frmOptions.Visible Then
                frmOptions.Hide
            End If
            CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=skins"
        Case SkinChange
            Set frmOptions.lvList.SelectedItem = frmOptions.lvList.ListItems.Item(11) 'sessions
            frmOptions.Show
        Case BuddySignOn
            nmnuView_MenuClick 2
        Case BuddySignOff
            nmnuView_MenuClick 2
    End Select
End Sub
Public Sub RealDataArrival(ByRef TheServer As clsActiveServer, ByVal bytesTotal As Long)
    Dim iData As String
    Dim intFL As Integer
    Dim intActiveServerIndex As Integer
    Static aData() As String
    Dim bData As String
    
    intActiveServerIndex = GetServerIndexFromActiveServer(TheServer)
    
    'use this error handler in case the array
    'aData() hasn't been initialized
    'In that case, the If line will cause
    'a Subscript out of Range exception
    'causing the next line to be executed
    'So, the ReDim-ention will occur
    'either if the data we are receiving
    'is the first data of the first server
    'or if a new server is receiving its
    'first data
    On Local Error Resume Next
    If intActiveServerIndex > UBound(aData) Then
        ReDim Preserve aData(intActiveServerIndex)
    End If
    
    On Local Error GoTo Invalid_Data
    TheServer.WinSockConnection.GetData bData, , bytesTotal
    
    'Log RAW Data
    If GetSetting("Node", "Options", "Debug\Raw", False) Then
        intFL = FreeFile
        Open App.Path & "\logs\raw.dat" For Append As #intFL
        Print #intFL, bData
        Close #intFL
    End If
    
    aData(intActiveServerIndex) = aData(intActiveServerIndex) & bData
       
    'On Error GoTo DataArrival_BugInCode
    Do
        iData = Strings.Left$(aData(intActiveServerIndex), InStr(1, aData(intActiveServerIndex), vbLf))
        DB.X "RAW-in:: " & Left$(iData, Len(iData) - 2)
        DataArrival iData, intActiveServerIndex
        aData(intActiveServerIndex) = Strings.Right$(aData(intActiveServerIndex), Len(aData(intActiveServerIndex)) - Len(iData))
        If InStr(1, aData(intActiveServerIndex), vbLf) <= 0 Then
            Exit Do
        End If
        DoEvents
    Loop
    'aData = Right$(aData, (Len(aData) - InStr(1, aData, vbLf)))
    Exit Sub
DataArrival_BugInCode:
'    addStatus SpecialSmiley("Arrow") & EVENT_PREFIX & " " & Language(64) & EVENT_SUFFIX & ": " & Language(65) & " DataArrival " & _
'              Language(66) & " :" & REASON_PREFIX & _
'              Err.Description & REASON_SUFFIX & ". " & Language(67) & " " & HTML_OPEN & "a href=""http://sourceforge.net/tracker/?func=add&group_id=94591&atid=608388""" & HTML_CLOSE & Language(68) & HTML_OPEN & "/a" & HTML_CLOSE & vbnewline
    ReportBug "RealDataArrival Error!"
Invalid_Data:
    'Resume
End Sub
Public Function executeCommand(ByVal strFullStatement As String, Optional ByVal TheServerIndex As Integer = -1) As Boolean
    'This function is used to execute Client commands.
    'it returns True if the command was executed
    Dim b_ExecuteCommand As Boolean 'a temporary variable storing the return value
    Dim strFileName As String 'a variable used to store the filename for file-related commands(for example /append)
    Dim CurFile As Integer 'an integer variable used to store the current open file index
    Dim strStatement As String 'the statement of the command(the command without its parameters)
    Dim ChannelId As Integer 'a variable used to store the current ChannelID(the channel-textbox index)
    Dim Nick1 As String 'a variable used to temporarily store a nickname
    Dim i As Integer 'our usual counter variable
    Dim CurFile1 As Integer 'index of a second open file
    Dim CurFile2 As Integer 'and for a third one
    Dim intX As Integer, intY As Integer 'variables storing the current mouse position
    Dim strTemp As String, strTemp2 As String 'two temporary variables used to store temporary information
    Dim intTemp As Integer 'another temporary variable used to store an integer
    Dim boolTemp As Boolean
    Dim thiswb As WebBrowser 'An object variable used to temporarily store the current WebBrowser
                             '(this is different from currentWB which stores the INDEX of the current WebBrowser)
    Dim tmpServer As clsActiveServer
    Dim newStatement As String
    Dim strChan As String
    Dim frmProfile As frmCustom 'used to load the buddy profile dialog
    Dim mnuCustomMenu As Menu
    Dim strMenuCaption As String
    Dim strMenuCommand As String
    Dim strPastServers() As String
    Dim strDescription As String
    Dim strDescriptionQuotes As String
    Dim intServerIndex As Integer
    Dim objNickList As NodeNickList
    Static frmEditServer As frmCustom
    Dim TheServer As clsActiveServer
    Dim nodetype As String, nodetype2 As String
    Dim strwebnum As Integer
    
    If TheServerIndex = -1 Then
        Set TheServer = CurrentActiveServer
    Else
        Set TheServer = ActiveServers(TheServerIndex)
    End If
    
    'let windows update everything
    DoEvents
    newStatement = strFullStatement
    'get the statement from the command using GetStatement function; use LCase as the case of the statement doesn't matter
    '(while the case of the parameters DOES matter; that's why we don't lCase the whole command)
    strStatement = Strings.LCase$(GetStatement(strFullStatement))
    'execute the command
    'assume the command was executed
    b_ExecuteCommand = True
    'only certain statements are client-statements. Only if one of these was typed execute it
    Select Case strStatement
        Case "names"
                namesboolean = True
                TheServer.SendData strFullStatement & vbNewLine
        Case "nick"
                newStatement = Replace(strFullStatement, strStatement & " ", vbNullString)
                frmOptions.txtNickname = newStatement
                frmOptions.cmdApply.value = True
                If LenB(TheServer.myNick) = 0 Then
                    TheServer.myNick = newStatement
                End If
                TheServer.SendData strFullStatement & vbNewLine
                GoTo Final_Mountain
        Case "topic"
                'If InStr(1, newstatement, "/") > 0 Then newstatement = Replace(newstatement, "/", vbnullstring)
                newStatement = Replace(strFullStatement, "topic #", "topic#")
                If InStr(1, newStatement, " ") > 0 Then
                    newStatement = Replace(newStatement, " ", " :", 1, 1)
                    newStatement = Replace(newStatement, "topic#", "topic #")
                    newStatement = Strings.Mid$(newStatement, 1, Len(newStatement) - (Len(newStatement) - InStr(1, newStatement, ":")))
                    strFullStatement = newStatement & Strings.Mid$(strFullStatement, Len(newStatement))
                Else
                    newStatement = Replace(newStatement, "topic#", "topic #")
                End If
                TheServer.SendData strFullStatement & vbNewLine
                GoTo Final_Mountain
        Case "msg"
                newStatement = Replace(Strings.LCase$(strFullStatement), "msg ", "privmsg ", 1, 1)
                newStatement = Replace(Strings.LCase$(newStatement), " ", ".", 1, 1)
                newStatement = Replace(newStatement, " ", " :", 1, 1)
                newStatement = Replace(newStatement, ".", " ", 1, 1)
                newStatement = Strings.Mid$(newStatement, 1, Len(newStatement) - (Len(newStatement) - InStr(1, newStatement, ":")))
                strFullStatement = newStatement & Strings.Mid$(strFullStatement, Len(newStatement) - 4)
                TheServer.SendData strFullStatement & vbNewLine
                GoTo Final_Mountain
        'Case "privmsg"
        '        newstatement = Replace(Strings.LCase$(strFullStatement), " ", ".", 1, 1)
        '        newstatement = Replace(newstatement, " ", " :", 1, 1)
        '        newstatement = Replace(newstatement, ".", " ", 1, 1)
        '        newstatement = Strings.Mid$(newstatement, 1, Len(newstatement) - (Len(newstatement) - InStr(1, newstatement, ":")))
        '        strFullStatement = newstatement & Strings.Mid$(strFullStatement, Len(newstatement))
        '        wsIRC.SendData strFullStatement & vbnewline
        '        GoTo Final_Mountain
        Case "xlist"
                TheServerIndex = GetParameter(strFullStatement, 2)
                If TheServerIndex = -1 Then
                    DB.X "Warning: ServerIndex -1 while trying to Sort Channel List!"
                Else
                    ActiveServers(TheServerIndex).bCLSorting = GetParameter(strFullStatement)
                    IRCAction ndChannelList, , , , "end of list"
                End If
        Case "me"
                newStatement = "privmsg "
                strFullStatement = newStatement & Strings.Mid$(strFullStatement, 4)
                strChan = TheServer.Tabs.SelectedItem.Caption
                If TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_Channel Or TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_Private Then
                    strFullStatement = AddPhrase(strFullStatement, strChan & " :" & "ACTION ", 8)
                    strFullStatement = strFullStatement & ""
                    IRCAction ndPrivMsg, TheServer.myNick, strChan, , Strings.Mid$(strFullStatement, InStr(1, strFullStatement, ""))
                    If InStr(1, TheServer.Tabs.SelectedItem.Caption, "@") > 0 And Strings.Right$(TheServer.Tabs.SelectedItem.Caption, 4) = CStr(Val(Strings.Right$(TheServer.Tabs.SelectedItem.Caption, 4))) Then
                        For i = 1 To wsDCCChat.Count - 1
                            If wsDCCChat.Item(i).LocalPort = Strings.Right$(TheServer.Tabs.SelectedItem.Caption, 4) Then
                                strFullStatement = Strings.Mid$(strFullStatement, InStr(1, strFullStatement, ""), Len(strFullStatement) - InStr(1, strFullStatement, ""))
                                wsDCCChat.Item(i).SendData strFullStatement & vbNewLine
                                GoTo Final_Mountain
                            End If
                        Next i
                    End If
                    TheServer.SendData strFullStatement & vbNewLine
                    GoTo Final_Mountain
                Else
                    'this is not a channel or query
                    AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(34) & EVENT_SUFFIX & vbNewLine, TheServer 'You are not on a channel
                    'exit sub; don't remove the text the user has typed. He or she may wanted to send it to a channel, but forgot to select it.
                    GoTo Final_Mountain
                End If
        
        'Statement connect.
        'Used to connect to a server.
        'Syntax:
        '/connect server port "server description"
        Case "connect"
            'command /connect server port
            'close any previous connection
            'we'll do that later
            'TheServer.WinSockConnection.Close
            'close all the channel windows
            TheServer.ClearAllNicklists
            'let Windows update
            DoEvents
            '36 = "Connect to this server"
            '37 = "Connecting to"
            'show "Connecting to" text on the status window
            boolTemp = IsParameter(strFullStatement, 3)
            If boolTemp Then
                strTemp = GetParameter(strFullStatement, 3)
            Else
                strTemp = GetParameter(strFullStatement, 1)
            End If
            AddStatus Language(37) & " " & _
                HTML_OPEN & "acronym title=""" & Language(36) & """" & HTML_CLOSE & _
                HTML_OPEN & "a href='NodeScript:/connect " & GetParameter(strFullStatement, 1) & " " & Val(GetParameter(strFullStatement, 2)) & " """ & strTemp & """ '" & HTML_CLOSE & _
                    strTemp & _
                HTML_OPEN & "/a" & HTML_CLOSE & HTML_OPEN & "/acronym" & HTML_CLOSE & "..." & vbNewLine, TheServer
            'let windows update(again!)
            DoEvents
            
            TheServer.ConnectRetry = False
            
            'establish the actual connection
            TheServer.WinSockConnection.Close
            
            'TO DO: Local port range
            TheServer.WinSockConnection.LocalPort = 0
            TheServer.HostName = GetParameter(strFullStatement, 1)
            TheServer.Port = Val(GetParameter(strFullStatement, 2))
            If IsParameter(strFullStatement, 3) Then
                strDescription = GetParameter(strFullStatement, 3)
            End If
            
            If LenB(strDescription) = 0 Then
                strDescription = CurrentActiveServer.HostName
            End If
            
            If Options.UseProxy Then
                TheServer.WinSockConnection.Connect Options.ProxyIP, Options.ProxyPort
            Else
                TheServer.WinSockConnection.Connect TheServer.HostName, TheServer.Port
            End If
            
            'Wait some time so that WinSock can start connecting
            Wait 0.1
            'select the status tab
            TheServer.Tabs.Tabs.Item(TheServer.GetTab(TheServer.GetStatusID)).Selected = True
            'save the current session
            SaveSession
            ReDim strPastServers(0)
            'assume the server does not exist
            intServerIndex = -1
            If FS.FileExists(App.Path & "\conf\servers.lst") Then
                CurFile1 = FreeFile
                Open App.Path & "\conf\servers.lst" For Input Access Read Shared As #CurFile1
                Do Until EOF(CurFile1)
                    i = i + 1
                    Line Input #CurFile1, strPastServers(UBound(strPastServers))
                    If LCase$(GetStatement(strPastServers(UBound(strPastServers)))) = LCase$(TheServer.HostName) Then
                        'server already exists in the past servers
                        'check to see if the port and description are the
                        'same
                        If GetParameter(strPastServers(UBound(strPastServers))) = TheServer.Port And _
                           GetParameter(strPastServers(UBound(strPastServers)), 2) = TheServer.Description Then
                            'they are
                            'we don't need to update anything
                            Close #CurFile1
                            GoTo Final_Mountain
                        Else
                            'the description or port
                            'aren't the same
                            'we will need to update them
                            'the server already exists
                            'save the index
                            intServerIndex = i
                        End If
                    End If
                    ReDim Preserve strPastServers(UBound(strPastServers) + 1)
                Loop
                Close #CurFile1
            End If
            
            'we have to update the server list
            CurFile1 = FreeFile
            Open App.Path & "\conf\servers.tmp" For Output Access Write Lock Write As #CurFile1
            For i = 0 To UBound(strPastServers)
                If i + 1 = intServerIndex Then
                    strDescriptionQuotes = IIf(InStr(1, strDescription, " ") > 0, """", vbNullString)
                    Print #CurFile1, TheServer.HostName & " " & TheServer.Port & " " & strDescriptionQuotes & strDescription & strDescriptionQuotes
                ElseIf LenB(strPastServers(i)) > 0 Then
                    'copy the server from the server list to the temporary file
                    Print #CurFile1, strPastServers(i)
                End If
            Next i
            If intServerIndex = -1 Then
                strDescriptionQuotes = IIf(InStr(1, strDescription, " ") > 0, """", vbNullString)
                Print #CurFile1, TheServer.HostName & " " & TheServer.Port & " " & strDescriptionQuotes & strDescription & strDescriptionQuotes
            End If
            Close #CurFile1
            
            If FS.FileExists(App.Path & "\conf\servers.lst") Then
                FS.DeleteFile App.Path & "\conf\servers.lst", True
            End If
            FS.MoveFile App.Path & "\conf\servers.tmp", App.Path & "\conf\servers.lst"
            LoadPanel "connect"
            tmrPanelRefreshSoon.Enabled = False
            tmrPanelRefreshSoon.Enabled = True
            
        Case "lag"
            'get the current lag to the server
            Load tmrLag(tmrLag.Count)
            tmrLag(tmrLag.Count - 1).Tag = 0
            tmrLag(tmrLag.Count - 1).Enabled = True
            TheServer.preExecute "/notice " & TheServer.myNick & " @" & tmrLag.Count - 1
                        
        'Statement nickmenu.
        'Used to display the popup menu at the nickmenu.
        '(used almost only by code: inside NickHTML code)
        'Syntax:
        '/nickmenu nickname
        Case "nickmenu"
            'get the parameter and store it to the CurrentNick variable
            'this variable is used later by the code in the menu events
            '(for example when `Query' item in the menu is clicked it
            ' uses this variable to determine to which one we are going
            ' to talk to)
            CurrentNick = vbNullString
            On Error Resume Next
            CurrentNick = GetParameter(strFullStatement)
            If LenB(CurrentNick) = 0 Then
                'private?
                'the code for the right-click menu
                'is located at webdocCurrentIRCWindow_onmousedown
            Else
                mnuKick.Enabled = True
                mnuMode.Enabled = True
                mnuWhisper.Enabled = True
                mnuNickClear.Enabled = False
                mnuNickViewLogs.Enabled = False
                For i = 0 To frmOptions.lstBdyNk.ListCount - 1
                    If frmOptions.lstBdyNk.List(i) = CurrentNick Then
                        mnuAddBuddy.Caption = Language(796)
                        i = -1
                        Exit For
                    End If
                Next
                If i <> -1 Then
                    mnuAddBuddy.Caption = Language(565)
                End If
                For i = 0 To frmOptions.lstIgnore(0).ListCount - 1
                    If frmOptions.lstIgnore(0).List(i) = CurrentNick Then
                        mnuNickIgnore.Caption = Language(841)
                        i = -1
                        Exit For
                    End If
                Next
                If i <> -1 Then
                    mnuNickIgnore.Caption = Language(77)
                End If
                mnuGiveHalfOp.Enabled = CurrentActiveServer.SupportsHalfop
                mnuTakeHalfOp.Enabled = mnuGiveHalfOp.Enabled
                
                mnuGiveOp.Enabled = CurrentActiveServer.SupportsOp
                mnuTakeOp.Enabled = mnuGiveOp.Enabled
                
                mnuGiveVoice.Enabled = CurrentActiveServer.SupportsVoice
                mnuTakeVoice.Enabled = mnuGiveVoice.Enabled
                
                'display popup menu
                Me.PopupMenu mnuNickPop
            End If
        
            
        'Statement savesize.
        'Used to save the current nicklist size.
        '(used almost only by code: inside FrameHTML code)
        'Syntax:
        '/savesize SizeToSave
        Case "savesize"
            'save current nicklist size; store it to the tag of the nicks listbox object
            TheServer.NickList_Size(TheServer.TabInfo(TheServer.Tabs.SelectedItem.index)) = GetParameter(strFullStatement)
            'just a debugging line; may be necessary if a bug regarding this code is found later
            'addStatus "New NickList size is " & GetParameter(strFullStatement)
        
        Case "savesetting"
            SaveSetting App.EXEName, "Options", GetParameter(strFullStatement), GetParameter(strFullStatement, 2)
            
        'Statement chanproperties.
        'Used to display the channel properties dialog of the currently selected channel.
        '(used almost only by code: inside MainChanHTML)
        'Syntax:
        '/chanproperties
        Case "chanproperties"
            'LoadDialog App.Path & "\data\dialogs\channel.xml"
            If FS.FileExists(App.Path & "/temp/banlist.html") Then
                FS.DeleteFile App.Path & "/temp/banlist.html", True
            End If
            TheServer.preExecute "/mode " & TheServer.Tabs.SelectedItem.Caption & " +b"
        
        Case "chanmodes"
            TheServer.preExecute "/mode " & TheServer.Tabs.SelectedItem.Caption
            Set frmChanModes = New frmCustom
            frmChanModes.DialogData = TheServer.NickList_Modes(TheServer.GetChanID(TheServer.Tabs.SelectedItem.Caption))
            frmChanModes.DialogData4 = TheServer.Tabs.SelectedItem.Caption
            LoadDialog App.Path & "/data/dialogs/chanmodes.xml", frmChanModes
            
        'Statement dialog.
        'Used to load and display an XML dialog file.
        '(used almost only by code: called by the program[for kick, disconnect, etc],
        ' but can also be used by a developer to see if certain dialog
        ' file actually works)
        'Syntax:
        '/dialog DialogFile
        Case "dialog"
            'use LoadDialog routine to load the dialog file passed as a parameter
            LoadDialog GetParameter(strFullStatement, 1)
        
        'Statement hop.
        'Used to rejoin a channel.
        'Syntax:
        '/hop
        Case "hop"
            If TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_Channel Then
                strTemp = TheServer.Tabs.SelectedItem.Caption
                TheServer.preExecute "/part " & strTemp
                'necessary
                'because you have to have time to leave
                'first before you send the JOIN command
                'again.
                
                'TO DO:
                'This won't work for LAG > 0.5s
                'We should JOIN only after
                'IRCAction(ndPart) has been fired
                Wait 0.5
                TheServer.preExecute "/join " & strTemp
            Else
                'you're not on a channel
                AddStatus EVENT_PREFIX & Language(34) & EVENT_SUFFIX & vbNewLine, TheServer
            End If
            
        Case "ctcp"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("ctcp "))
            strTemp = Replace(strTemp, " ", " :" & MIRC_CTCP, 1, 1)
            If LCase$(Right$(strTemp, Len("ping"))) = "ping" Then
                strTemp = strTemp & " " & ToTimeStamp 'Add a UNIX Timestamp
            End If
            strTemp = strTemp & MIRC_CTCP
            CurrentActiveServer.SendData "PRIVMSG " & strTemp & vbNewLine
            
        'Statement amsg.
        'Used to send a message to all the channels you are in.
        'Syntax:
        '/amsg Message
        Case "amsg"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("amsg "))
            executeCommand "message-to-all c " & strTemp
            
        Case "aamsg"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("aamsg "))
            executeCommand "message-to-all * " & strTemp
        
        Case "pamsg"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("pamsg "))
            executeCommand "message-to-all p " & strTemp
            
        Case "ame"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("ame "))
            executeCommand "message-to-all c /me " & strTemp
        
        Case "aame"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("aame "))
            executeCommand "message-to-all * /me " & strTemp
        
        Case "pame"
            strTemp = Right$(strFullStatement, Len(strFullStatement) - Len("pame "))
            executeCommand "message-to-all p /me " & strTemp
        
        Case "message-to-all"
            strTemp2 = GetParameter(strFullStatement)
            strTemp = Right$(strFullStatement, Len(strFullStatement) - InStr(InStr(1, strFullStatement, " ") + 1, strFullStatement, " "))
            intTemp = TheServer.Tabs.SelectedItem.index
            Screen.MousePointer = vbHourglass
            LockWindowUpdate TheServer.Tabs.hwnd
            'LockWindowUpdate wbStatus(currentWB).hwnd
            'LockWindowUpdate Me.hwnd
            'wbStatus(currentWB).Visible = False
            TheServer.Tabs.Visible = False
            BSCodeCall = True
            For i = 1 To TheServer.Tabs.Tabs.Count
                If (TheServer.TabType(i) = TabType_Channel And (strTemp2 = "c" Or strTemp2 = "*")) Or (TheServer.TabType(i) = TabType_Private And (strTemp2 = "p" Or strTemp2 = "*")) Then
                    TheServer.Tabs.Tabs(i).Selected = True
                    TheServer.preExecute strTemp
                End If
            Next i
            BSCodeCall = False
            TheServer.Tabs.Visible = True
            'wbStatus(currentWB).Visible = True
            LockWindowUpdate 0&
            TheServer.Tabs.Tabs(intTemp).Selected = True
            Screen.MousePointer = vbDefault
            
        
        'Statement splay.
        'Used to play a sound file using the MediaPlayer object.
        '(used almost only by code: inside scripts to play event-sounds)
        'Syntax:
        '/splay SoundFile
        'Case "splay"
            'mp1.Stop
            'set the filename of the media file that will be played
            'MP.FileName = GetParameter(strFullStatement)
            'use the play command(in the case the user has disabled auto-play option)
            'MP.Play
        
        'Statement sstop.
        'Used to send the stop command to the MediaPlayer object in order to stop a media file which is current playing.
        '(can be used be users annoyed by stupid event-sounds ;)
        'Syntax:
        '/sstop
        '(takes no parameters)
        'Case "sstop"
            'stop what's currently playing
            'MP.Stop
        
        'Statement runscript.
        'Used to run a specific routine inside a scriptfile.
        '(used almost only by code: inside scripts to play event-sounds)
        'Syntax:
        '/runscript routine
        Case "runscript"
            'use RunScript method to execute the script passed as a parameter
            'if there wasn't anything passed, pass nothing to the method( vbnullstring )
            RunScript GetParameter(strFullStatement) ', IIf(IsParameter(strFullStatement, 2), GetParameter(strFullStatement, 2), vbnullstring)
            
        Case "tempscript"
            
        'Statement browse.
        'Used to open a new website tab and browse to a specific site.
        '(used almost only by code: by dialogs, or inside websites, when a website(http://) link is clicked, etc.)
        'Syntax:
        '/browse URL
        Case "browse"
            If IsParameter(strFullStatement) Then
                strTemp = GetParameter(strFullStatement)
                strTemp = Replace(strTemp, "$$app$$", App.Path)
            Else
                strTemp = Options.StartPageURL
            End If
            
            If Options.BrowseInternalBrowser Then
                'Create a new website tab; set its caption to "WebSite Tab"
                'set its image to the WebSite tab image
                TheServer.Tabs.Tabs.Add , , Language(38), TabImage_WebSite
                TheServer.TabInfo.Add wbStatus.Count
                TheServer.TabType.Add TabType_WebSite
                Load wbStatus(wbStatus.Count)
                Set thiswb = wbStatus(wbStatus.Count - 1)
                thiswb.Visible = True
                thiswb.Navigate2 strTemp
                thiswb.ZOrder 0
                TheServer.preExecute "/focus last"
                CreateWSRoot
                For i = 1 To tvConnections.Nodes.Count
                    If tvConnections.Nodes.Item(i).Key = "wbroot" Then
                        tvConnections.Nodes.Add tvConnections.Nodes.Item(i), tvwChild, "w_" & GetServerIndexFromActiveServer(TheServer) & "_" & wbStatus.Count - 1, Language(38), TabImage_WebSite
                    End If
                Next i
                MenusToFront
                SaveSession
            Else
                xShell strTemp, 0
            End If
            
        'Statement focus.
        'Used to focus a specific tab.
        '(used almost only by code: /browse uses it to focus the new tab, dialogs may use it as well)
        'Syntax:
        '/focus ( last | index | id | caption ) [ tab_identity ]
        'tab_identify parameter type depends on the previous parameter
        'if last was used it is ignored.
        Case "focus"
            'focus tab
            strTemp = GetParameter(strFullStatement)
            If strTemp = "last" Then
                'focus the last tab
                TheServer.Tabs.Tabs.Item(TheServer.Tabs.Tabs.Count).Selected = True
                GoTo Final_Mountain
            End If
            strTemp2 = GetParameter(strFullStatement, 2)
            Select Case strTemp
                Case "index"
                    'focus by index
                    TheServer.Tabs.Tabs.Item(strTemp2).Selected = True
                Case "id"
                    'focus by id
                    TheServer.Tabs.Tabs.Item(TheServer.GetTab(strTemp2)).Selected = True
                Case "caption"
                    'focus by caption
                    TheServer.Tabs.Tabs.Item(TheServer.GetTab(TheServer.GetChanID(strTemp2))).Selected = True
            End Select
            
        'Statement append.
        'Used to add something to a file.
        '(used almost only by code: reserved for later use)
        'Syntax:
        '/append
        Case "append"
            'Add a line into a file
            strFileName = GetParameter(strFullStatement)
            If FS.FileExists(strFileName) Then
                CurFile = FreeFile
                Open strFileName For Append Access Write Lock Write As #CurFile
                Print #CurFile, GetParameter(strFullStatement, 2)
                Close #CurFile
            End If
            
        'Statement echo.
        'Display something in the current window.
        '(used almost only by code: by dialogs to display certain text, etc.)
        'Syntax:
        '/echo TextToDisplay
        Case "echo"
            AddStatus GetParameter(strFullStatement) & vbNewLine, TheServer
        
        Case "notice"
            On Error Resume Next
            Nick1 = GetParameter(strFullStatement, 2)
            If Err Then
            Else
                GoTo Final_Mountain
            End If
            Err.Clear
            Nick1 = GetParameter(strFullStatement)
            If Err Then
            Else
                executeCommand "query " & GetParameter(strFullStatement)
            End If
            
        'Statement query.
        'Open a new private tab.
        '(used either by code(when the user selects `Query' from the nick popup menu) or by the user)
        'Syntax:
        '/query Nickname
        Case "query"
            If MaxMode Then
                txtSend_KeyDown vbKeyF10, 0
            End If
            For i = 1 To TheServer.Tabs.Tabs.Count
                If Strings.LCase$(TheServer.Tabs.Tabs(i).Caption) = Strings.LCase$(xLet(Nick1, GetParameter(strFullStatement))) Then
                    GoTo Focus_Query
                End If
            Next i
            TheServer.Tabs.Tabs.Add , , Nick1, TabImage_Private
            TheServer.TabInfo.Add xLet(ChannelId, TheServer.GetEmptyChannelID)
            TheServer.TabType.Add TabType_Private
            tvConnections.Nodes.Add TheServer.ServerNode, tvwChild, "p_" & GetServerIndexFromActiveServer(TheServer) & "_" & Nick1, Nick1, TabImage_Private
            UpdateTabsBar
            If ChannelId > TheServer.IRCData_Count - 1 Then
                TheServer.IRCData_ReDim ChannelId
            End If
            TheServer.IRCData(ChannelId) = vbNullString
Focus_Query:
            TheServer.Tabs.Tabs(TheServer.Tabs.Tabs.Count).Selected = True
            'request NDC if we're auto-requesting
            If Options.AutoNDC Then
                mnuNDCConnect_Click
            End If
            showison = False
            For i = 1 To TheServer.Tabs.Tabs.Count
                If TheServer.TabType(i) = TabType_Private Then
                    If InStr(1, TheServer.Tabs.Tabs.Item(i).Caption, "(") > 0 Then
                        Nick1 = Nick1 & " " & Strings.Left$(TheServer.Tabs.Tabs.Item(i).Caption, InStr(1, TheServer.Tabs.Tabs.Item(i).Caption, "(") - 1)
                    Else
                        Nick1 = Nick1 & " " & TheServer.Tabs.Tabs.Item(i).Caption
                    End If
                End If
            Next i
            TheServer.preExecute "/ison " & Nick1
            
        'Statement close.
        'Close the selected website or private tab.
        '(used almost only by code: when the user selectes `Close' from the tabs popup menu)
        'Syntax:
        '/close
        '(takes no parameters)
        Case "close"
            intTemp = TheServer.TabType(TheServer.Tabs.SelectedItem.index)
            'Close Query Tab
            Select Case intTemp
                Case TabType_Private
                    
                    TheServer.DCCChats_Disconnect TheServer.Tabs.SelectedItem.Caption
                
                    For i = 1 To tvConnections.Nodes.Count
                        nodetype = Strings.Left$(tvConnections.Nodes.Item(i).Key, 1)
                        If nodetype = "d" Then
                            If tvConnections.Nodes.Item(i).Key = "d_" & GetServerIndexFromActiveServer(TheServer) & "_" & TheServer.Tabs.SelectedItem.Caption & "_chat" & TheServer.Tabs.SelectedItem.index Then
                                tvConnections.Nodes.Remove i
                                Exit For
                            End If
                        ElseIf nodetype = "p" Then
                            If tvConnections.Nodes.Item(i).Key = "p_" & GetServerIndexFromActiveServer(TheServer) & "_" & TheServer.Tabs.SelectedItem.Caption Then
                                tvConnections.Nodes.Remove i
                                Exit For
                            End If
                        End If
                    Next i
                    
                    'Unload txtStatus(TabInfo(tsTabs.SelectedItem.index))
                Case TabType_WebSite
                    On Error Resume Next 'DOM document not assosiated
                    Set webdocWebSite(TheServer.TabInfo(TheServer.Tabs.SelectedItem.index)) = Nothing
                    For i = 1 To tvConnections.Nodes.Count
                        nodetype = Strings.Left$(tvConnections.Nodes.Item(i).Key, 1)
                        If nodetype = "w" Then
                            nodetype = Right$(tvConnections.Nodes.Item(i).Key, Len(tvConnections.Nodes.Item(i).Key) - InStr(3, tvConnections.Nodes.Item(i).Key, "_"))
                            If nodetype = TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) Then
                                tvConnections.Nodes.Remove i
                                RemoveWSRoot
                                Exit For
                            End If
                        End If
                    Next i
                    Unload wbStatus(TheServer.TabInfo(TheServer.Tabs.SelectedItem.index))
                Case TabType_DCCFile
                    On Error Resume Next 'the wsDCC.Item(send) object could have been unloaded
                    If TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) = 1 Then 'sending
                        TheServer.DCCFile_CancelTransfer TheServer.GetDCCFromTab(TheServer.Tabs.SelectedItem.index, False), False
                    Else '= 0; recieving
                        TheServer.DCCFile_CancelTransfer TheServer.GetDCCFromTab(TheServer.Tabs.SelectedItem.index, False), True
                    End If
                Case TabType_Channel
                    If TheServer.NickList_List(TheServer.GetChanID(TheServer.Tabs.SelectedItem.Caption)).Count > 0 Then
                        TheServer.SendData "PART " & TheServer.Tabs.SelectedItem.Caption & vbNewLine
                        GoTo Final_Mountain 'do not remove tab; it will be removed when we part
                    Else 'simply remove tab
                        'remove tvConnections node
                        For i = 1 To tvConnections.Nodes.Count
                            If tvConnections.Nodes.Item(i).Key = "c_" & GetServerIndexFromActiveServer(TheServer) & "_" & TheServer.Tabs.SelectedItem.Caption Then
                                tvConnections.Nodes.Remove i
                                Exit For
                            End If
                        Next i
                        'Unload txtStatus(TabInfo(tsTabs.SelectedItem.index))
                        TheServer.NickList_List_SetToNothing TheServer.TabInfo(TheServer.Tabs.SelectedItem.index)
                    End If
                Case Else
                    '"Sorry, but you can not close this tab!"
                    AddStatus Language(39) & vbNewLine, TheServer
                    GoTo Final_Mountain
            End Select
            
            TheServer.DCCTabRemove TheServer.Tabs.SelectedItem.index
            
            TheServer.TabInfo.Remove TheServer.Tabs.SelectedItem.index
            TheServer.TabType.Remove TheServer.Tabs.SelectedItem.index
            TheServer.Tabs.Tabs.Remove TheServer.Tabs.SelectedItem.index
            SaveSession
            
            Set TheServer.Tabs.SelectedItem = TheServer.Tabs.Tabs.Item(TheServer.GetTab(TheServer.GetStatusID))
            tmrMakeItQuicker.Enabled = False
            tmrMakeItQuicker.Enabled = True
            If strCurrentPanel = "avatar" Then
                wbPanel_DocumentComplete Nothing, vbNullString
            End If
            UpdateTabsBar
            Form_Resize
            
        'Statement clear.
        'Clear the current irc-textbox.
        'Syntax:
        '/clear
        '(takes no parameters)
        Case "clear"
            'if the selected tab is a channel or a private message or the status tab
            '(if it's not a website tab it must be one of these)
            If Not TheServer.TabType(TheServer.Tabs.SelectedItem.index) = TabType_WebSite Then
                'clear the text
                TheServer.IRCData(TheServer.GetChanID(TheServer.Tabs.SelectedItem.Caption)) = vbNullString
                'txtStatus(GetChanID(tsTabs.SelectedItem.Caption)).Text = vbnullstring
                'and update view
                buildStatus
            End If
           
        'Statement buddy.
        'Check to see if any buddies are online
        'Syntax:
        '/buddy
        '(takes no parameters)
        Case "buddy"
            If frmOptions.lstBdyNk.ListCount > 0 Then
                For i = 0 To UBound(AllBuddies)
                    For intTemp = 0 To AllBuddies(i).ServerCount
                        strTemp = AllBuddies(i).Servers(intTemp) & ", "
                    Next intTemp
                    strTemp = Left$(strTemp, Len(strTemp) - 2)
                    If AllBuddies(i).isOnline = True Then
                        MsgBox AllBuddies(i).Name & " is online at " & strTemp
                    Else
                        MsgBox AllBuddies(i).Name & " is not online"
                    End If
                Next i
            Else
                MsgBox "You dont have any buddies"
            End If
            
        Case "debug"
            DB.ShowDebugWindow
            
        'this is the old code
        'Case "buddy"
        '    Dim bdyname(1 To 100)
        '    Dim bdynamelist As Integer
        '    Dim nicknum As Integer
        '    Dim roomnum As Integer
        '    Dim goodname As Integer
        '    Dim foundone As Integer
        '    bdynamelist = 1
        '    showwhois = False
        '    Do Until bdynamelist = frmOptions.lstBdyNk.ListCount + 1
        '        bdyname(bdynamelist) = frmOptions.lstBdyNk.List(bdynamelist - 1)
        '        bdynamelist = bdynamelist + 1
        '    Loop
        '    If Not (fs.FileExists(App.Path & "/conf/buddyonline.lst")) Then
        '        inttemp = FreeFile
        '        Open App.Path & "/conf/buddyonline.lst" For Output As #inttemp
        '        Close #inttemp
        '        For i = 1 To bdynamelist - 1
        '            If wsIRC.State = 7 Then wsIRC.SendData "WHOIS " & bdyname(i) & vbnewline
        '        Next i
        '    Else
        '        If buddieschecked < 1 Or showwhois = True Then
        '            Kill App.Path & "/conf/buddyonline.lst"
        '            executeCommand "buddy"
        '            GoTo Final_Mountain:
        '        End If
        '        If buddieschecked = Val(bdynamelist) - 1 Then
        '            intX = FreeFile
        '            Open App.Path & "/conf/buddyonline2.lst" For Output As #intX
        '            For nicknum = 1 To bdynamelist - 1
        '                intY = 0
        '                inttemp = FreeFile
        '                Open App.Path & "/conf/buddyonline.lst" For Input As #inttemp
        '                Do Until EOF(inttemp)
        '                    Line Input #inttemp, strtemp
        '                    If Strings.Left$(strtemp, InStr(1, strtemp, " ") - 1) = bdyname(nicknum) Then
        '                        Print #intX, strtemp
        '                        intY = 1
        '                    End If
        '                Loop
        '                Close #inttemp
        '                If intY = 0 Then
        '                    Print #intX, bdyname(nicknum) & " offline"
        '                    Set fs = New FileSystemObject
        '                    If Not (fs.FileExists(App.Path & "/conf/buddies/" & bdyname(nicknum) & ".info")) Then
        '                        inttemp = FreeFile
        '                        Open App.Path & "/conf/buddies/" & bdyname(nicknum) & ".info" For Output As #inttemp
        '                            Print #inttemp, "Real Name: " & Language(261)
        '                            Print #inttemp, "Version: " & Language(261)
        '                            Print #inttemp, "Last Seen: " & Language(519)
        '                            Print #inttemp, "Additional Info: "
        '                        Close #inttemp
        '                    End If
        '                End If
        '            Next nicknum
        '            Close #intX
        '            Kill App.Path & "/conf/buddyonline.lst"
        '            fs.MoveFile App.Path & "/conf/buddyonline2.lst", App.Path & "/conf/buddyonline.lst"
        '            buddieschecked = 0
        '            tmrPanelRefreshSoon.Enabled = False
        '            tmrPanelRefreshSoon.Enabled = True
        '            showwhois = True
        '        End If
        '    End If
            
        Case "cancel_connect_retry"
            Set tmpServer = ActiveServers(GetParameter(strFullStatement))
            If tmpServer.ConnectRetry Then
                tmpServer.ConnectRetry = False
                AddStatus Language(781) & vbNewLine, tmpServer
                'wbStatus(tmpServer.GetStatusID).Document.All.Item("cancel_connection_link_" & GetParameter(strFullStatement, 2)).href = vbnullstring
            End If
            
        Case "servers-organize"
            'Used to display the `Organize My Servers' dialog
            'Ran by the panel Connect
            'build the servers list
            
            'display dialog
            Set frmOrganize = New frmCustom
            frmOrganize.DialogData = CreateServersList(True)
            LoadDialog App.Path & "/data/dialogs/servers.xml", frmOrganize
        
        Case "alter-servers-save-edit"
            intServerIndex = GetParameter(strFullStatement)
            AlterServer intServerIndex, False, GetParameter(strFullStatement, 4), GetParameter(strFullStatement, 2), GetParameter(strFullStatement, 3)
            If LCase$(strCurrentPanel) = "connect" Then
                wbPanel_DocumentComplete Nothing, vbNullString
            End If
            frmOrganize.Hide
            Set frmOrganize = Nothing
            executeCommand "servers-organize"
            
        Case "alter-servers-edit"
            Set frmEditServer = New frmCustom
            
            intServerIndex = GetParameter(strFullStatement)
            frmEditServer.DialogData = intServerIndex
            LoadDialog App.Path & "\data\dialogs\edit_server.xml", frmEditServer
            
        Case "alter-servers-move"
            intServerIndex = GetParameter(strFullStatement)
            'if the 2nd parameter is 1, then we're moving the server up
            'if it's not, we're moving it down
            boolTemp = GetParameter(strFullStatement, 2) = "1"
            MoveServer intServerIndex, boolTemp
            
        Case "alter-servers-delete"
            intServerIndex = GetParameter(strFullStatement)
            AlterServer intServerIndex, True
            
        Case "alter-servers-sort"
            SortServers
            
        Case "hot-keys-add"
            HotKeysAdd GetParameter(strFullStatement), GetParameter(strFullStatement, 2)
            If frmOptions.Visible Then
                frmOptions.LoadAll
            End If
            
        Case "run"
            On Error Resume Next
            'ShellExecute Me.hWnd, "open", Replace(GetParameter(strFullStatement), "/", "\"), vbnullstring, vbnullstring, 0
            'Shell "explorer """ & Replace(GetParameter(strFullStatement), "/", "\") & """", vbNormalFocus
            strFullStatement = Replace(strFullStatement, "$$app$$", App.Path)
            'xShell """" & strFullStatement & """ """"", 0
            ShellExecute 0, "open", """" & Replace(GetParameter(strFullStatement), "/", "\") & """", """""", vbNullString, SW_SHOW
            If Err Then
                Shell "explorer.exe """ & Replace(GetParameter(strFullStatement), "/", "\") & """", vbMaximizedFocus
            End If
            
        'Statement </script>.
        'It's added here so it the last line of every script is not sent to the server: it just does nothing.
        '(used only by code: the last line of scripts)
        'Syntax:
        '/</script>
        '(takes no parameters)
        Case "</script>"
            'Don't send the last line of any script
        
        '/my-avatar-change
        'called by the avatars panel
        Case "my-avatar-change"
            cdPickAvatar.DialogTitle = Language(457)
            On Error Resume Next
            'let the user pick an avatar
            cdPickAvatar.ShowOpen
            If Err.Number = 0 Then
                'only if the user didn't cancel
                'copy the avatar to the configuration folder
                FileCopy cdPickAvatar.FileName, App.Path & "\conf\myavatar.jpg"
                
            End If
            If strCurrentPanel = "avatar" Then
                tmrPanelRefreshSoon.Enabled = False
                tmrPanelRefreshSoon.Enabled = True
            End If
            For i = 1 To UBound(NDCConnections)
                If wsNDC.Item(i).State = sckConnected Then
                    NDCRaverixSendRequest i
                End If
            Next i
            
        '/my-avatar-remove
        'called by the avatars panel
        Case "my-avatar-remove"
            'remove the avatar from the configuration folder
            Kill App.Path & "\conf\myavatar.jpg"
            If strCurrentPanel = "avatar" Then
                tmrPanelRefreshSoon.Enabled = False
                tmrPanelRefreshSoon.Enabled = True
            End If
            For i = 1 To UBound(NDCConnections)
                If wsNDC.Item(i).State = sckConnected Then
                    NDCRaverixSendRequest i
                End If
            Next i
        
        '/favweb-add
        'called by favorites panel
        Case "favweb-add"
            intTemp = FreeFile
            strTemp = webdocPanel.getElementById("web_name").getAttribute("value")
            strTemp2 = webdocPanel.getElementById("web_URL").getAttribute("value")
            If LenB(strTemp) > 0 And LenB(strTemp2) > 0 Then
                Open App.Path & "/conf/favwebs.lst" For Append As #intTemp
                    Print #intTemp, webdocPanel.getElementById("web_name").getAttribute("value")
                    Print #intTemp, webdocPanel.getElementById("web_URL").getAttribute("value")
                Close #intTemp
                webdocPanel.getElementById("web_name").setAttribute "value", vbNullString
                webdocPanel.getElementById("web_URL").setAttribute "value", vbNullString
                nmnuView_MenuClick 3
            Else
                xNodeTag webdocPanel, "lang_savefavresult", Language(485)
            End If
        '/remove-fav
        'called by favorites panel
        Case "remove-fav"
            intY = 0
            strFullStatement = LCase$(Replace(strFullStatement, "remove-fav ", vbNullString))
            intTemp = FreeFile
            Open App.Path & "/conf/favwebs.lst" For Input As #intTemp
            intX = FreeFile
            Open App.Path & "/conf/favwebs2.lst" For Output As #intX
                Do Until EOF(intTemp)
                    Line Input #intTemp, strTemp
                    If strTemp <> strFullStatement Then
                            Print #intX, strTemp
                            Line Input #intTemp, strTemp
                            Print #intX, strTemp
                    Else
                        If intY = 0 Then
                            Line Input #intTemp, strTemp
                            intY = 1
                        Else
                            Print #intX, strTemp
                            Line Input #intTemp, strTemp
                            Print #intX, strTemp
                        End If
                    End If
                Loop
            Close #intTemp
            Close #intX
            Kill App.Path & "/conf/favwebs.lst"
            Set FS = New FileSystemObject
            FS.MoveFile App.Path & "/conf/favwebs2.lst", App.Path & "/conf/favwebs.lst"
            nmnuView_MenuClick 3
            
        '/buddyoptview
        'called by buddy list panel
        Case "buddyoptview"
            For i = 1 To frmOptions.tvOptions.Nodes.Count
                If frmOptions.tvOptions.Nodes.Item(i).Key = "k06" Then
                    Set frmOptions.tvOptions.SelectedItem = frmOptions.tvOptions.Nodes.Item(i)
                    frmOptions.tvOptions_NodeClick frmOptions.tvOptions.SelectedItem
                    frmOptions.Show
                    frmOptions.tvOptions.SetFocus
                    Exit For
                End If
            Next i
            
        Case "buddyadd"
            frmOptions.cmdAddnick_Click
        
        Case "buddychangeview"
            strTemp = GetParameter(strFullStatement)
            Set webdocProfile = wbPanel.Document
            If webdocProfile.All(strTemp).Style.display = "inline" Then
                webdocProfile.All("tree").src = App.Path & "/data/skins/" & ThisSkin.TreeCollapsedImage
                webdocProfile.All(strTemp).Style.display = "none"
            Else
                webdocProfile.All("tree").src = App.Path & "/data/skins/" & ThisSkin.TreeExpanedImage
                webdocProfile.All(strTemp).Style.display = "inline"
            End If
            
            
        '/buddyprofile
        'called by buddy list panel
        Case "buddyprofile"
            strTemp = GetParameter(strFullStatement)
            Set frmProfile = New frmCustom
            LoadDialog App.Path & "\data\dialogs\profile.xml", frmProfile
            'here you need to set all the properties
            'of the dialog window
            'Note: take a look at profile.xml
            'We load an .html
            'file that is going to be displayed
            'at the <web> component on the dialog
            'here we need to get the DOM object of the Document
            'that has been loaded
            Set webdocProfile = frmProfile.wbCustom(1).Document
            Wait 0.1
            'TO DO: ^ may be inaccurate; move code to frmCustom.wbCustom_DocumentComplete; do check using ActiveDialog
            With webdocProfile.All
                xNodeTag webdocProfile, "user_name_info", Language(500)
                
                Set FS = New FileSystemObject
                If FS.FileExists(App.Path & "/conf/buddies/" & strTemp & ".info") Then
                    intTemp = FreeFile
                    Open App.Path & "/conf/buddies/" & strTemp & ".info" For Input As #intTemp
                        Line Input #intTemp, strTemp2
                        xNodeTag webdocProfile, "lang_realname", "&nbsp;&nbsp;&nbsp;" & Language(128) & ": " & Strings.Right$(strTemp2, Len(strTemp2) - InStr(1, strTemp2, ":"))
                        Line Input #intTemp, strTemp2
                        xNodeTag webdocProfile, "lang_version", "&nbsp;&nbsp;&nbsp;" & "<a href=""JavaScript:getversion(" & "'" & strTemp & "'" & ")"""">" & Language(342) & "</a>: " & Strings.Right$(strTemp2, Len(strTemp2) - InStr(1, strTemp2, " "))
                        Line Input #intTemp, strTemp2
                        xNodeTag webdocProfile, "lang_lastseen", "&nbsp;&nbsp;&nbsp;" & Language(501) & ": " & Strings.Right$(strTemp2, Len(strTemp2) - InStr(1, strTemp2, ":"))
                        Line Input #intTemp, strTemp2
                        xNodeTag webdocProfile, "lang_additionalinfo", "&nbsp;&nbsp;&nbsp;" & Language(502) & ":"
                        xNodeTag webdocProfile, "infotext", Strings.Right$(strTemp2, Len(strTemp2) - InStr(1, strTemp2, ":"))
                        xNodeTag webdocProfile, "lang_edittext", "<a href=""JavaScript:editinfoJ(" & "'" & strTemp & "'" & ")"""">Edit</a>"
                        For i = 0 To UBound(NDCConnections)
                            If NDCConnections(i).strNicknameA = strTemp Then
                                xNodeTag webdocProfile, "lang_timezone", "&nbsp;&nbsp;&nbsp;" & Language(518) & ": " & NDCConnections(i).bTimeZone
                                i = -1
                                Exit For
                            End If
                        Next i
                        If i <> -1 Then xNodeTag webdocProfile, "lang_timezone", "&nbsp;&nbsp;&nbsp;" & Language(518) & ": " & Language(261)
                        
                    Close #intTemp
                Else
                    'TO DO:
                    'show "unknown"
                    'and create the file!!
                End If
                xNodeTag webdocProfile, "lang_cancel", Language(515)
                Set FS = New FileSystemObject
                If FS.FileExists(App.Path & "/conf/avatars/" & strTemp & ".jpg") Then
                    xNodeTag webdocProfile, "img_avatar", "<img src=""" & App.Path & "/conf/avatars/" & strTemp & ".jpg"" width=""147"" height=""140"" align=""left"" border=""2"">", "xnode_xpanel_"
                Else
                    xNodeTag webdocProfile, "img_avatar", "<img src=""" & App.Path & "\data\graphics\splash.jpg"" width=""147"" height=""140"" align=""left"" border=""2"">", "xnode_xpanel_"
                End If
            End With
            'and then set everything
            
            'set title of window
            frmProfile.Caption = Replace(Language(499), "%1", strTemp)
        Case "editinfo"
            strTemp = GetParameter(strFullStatement)
            Set FS = New FileSystemObject
            If FS.FileExists(App.Path & "/conf/buddies/" & strTemp & ".info") Then
                intTemp = FreeFile
                Open App.Path & "/conf/buddies/" & strTemp & ".info" For Input As #intTemp
                intX = FreeFile
                Open App.Path & "/conf/buddies/" & strTemp & "2.info" For Output As #intX
                    Line Input #intTemp, strTemp2
                    Print #intX, strTemp2
                    Line Input #intTemp, strTemp2
                    Print #intX, strTemp2
                    Line Input #intTemp, strTemp2
                    Print #intX, strTemp2
                    strTemp2 = InputBox("Enter additional info for buddy", "Edit Buddy Profile")
                    If LenB(strTemp2) = 0 Then
                        Line Input #intTemp, strTemp2
                        strTemp2 = Strings.Right$(strTemp2, Len(strTemp2) - InStr(1, strTemp2, ":"))
                    End If
                    Print #intX, "Additional Info: " & strTemp2
                Close #intX
                Close #intTemp
                Set FS = New FileSystemObject
                FS.DeleteFile App.Path & "/conf/buddies/" & strTemp & ".info"
                FS.MoveFile App.Path & "/conf/buddies/" & strTemp & "2.info", App.Path & "/conf/buddies/" & strTemp & ".info"
            End If
            TheServer.preExecute "/buddyprofile " & strTemp
        
        Case "buddygetversion"
            strTemp = GetParameter(strFullStatement)
            TheServer.SendData "PRIVMSG " & strTemp & " " & MIRC_CTCP & "VERSION" & MIRC_CTCP & vbNewLine
            Wait 1
            TheServer.preExecute "/buddyprofile " & strTemp
            
        Case "nicklist-hide"
            If webdocChanFrameSet.All.tags("frameset").Item(1).cols = "*,0" Then
                webdocChanFrameSet.All.tags("frameset").Item(1).cols = "*,180"
            Else
                webdocChanFrameSet.All.tags("frameset").Item(1).cols = "*,0"
            End If
            
        Case "list"
            TheServer.bCLSorting = 255
            b_ExecuteCommand = False
            
        'another statement
        Case Else
            'we didn't execute it
            b_ExecuteCommand = False
    End Select
Final_Mountain:
    'this is the point the actual return value is returned
    'please do not use `Exit Function' in commands above
    'to avoid execution of following parts of code.
    'Instead use `GoTo Fianl_Mountain' in so this
    'code is executed.
    executeCommand = b_ExecuteCommand
End Function
Public Sub IRCAction(ByRef Action As NodeAction, Optional ByVal Nick1 As String, Optional ByVal Channel As String, Optional ByVal Nick2 As String, Optional ByVal strText As String, Optional ByVal TheServerIndex As Integer = -1)
    'This is the actual sub where all actions are taken
    'depending on what was recieved from the server.
    'It is called from the sub DataArrival which
    'parses the data and passes it here.
    Dim lTemp As Long, intTemp As Integer 'two temporary variables, one for long types and one for integers
    Dim dTemp As Double
    Dim strTemp As String, strTemp2 As String, strTemp3 As String 'three more for strings.
    Dim bolTemp As Boolean 'another temporary variable for booleans
    Dim i As Long, i2 As Long 'two counter variables for loops
                              '(mozillagodzilla): Longs are faster for loops.
    Dim CurFile1 As Integer, CurFile2 As Integer 'two variables used to store file indexes
    Dim ChannelId As Long 'temporary variable used to store the Channel's ID.
    Dim objTabData As TextBox 'another one which stores an irc-textbox object for quicker access
    Dim names() As String 'array used to temporarily store the nicknames inside a channel
    Dim CurrentNickList As Collection 'object variable used for quick access to a specific nicklist box
    Dim AllModes() As String 'array used to store the modes which just changed on a channel(they can be several)
    Dim CurrentMode As Boolean 'tells us whether the current mode is set or unset
    Dim strNick As String, strNick2 As String 'two variables used to temporarily store two nicknames
    Dim CheckBud As Integer
    Dim intNoArgsModesCount As Integer
    Dim isHiddenTransfer As Boolean
    Dim Pos As Integer
    Dim CTCPIgnoring As Boolean
    Dim cChannels As Collection
    Dim cTopics As Collection
    Dim cUsers As Collection
    Dim cCLSorting As Collection
    Static RetrievingChanList As Boolean
    Dim TheServer As clsActiveServer
    Dim frmChannels As frmCustom, frmBuddy As frmCustom
    Dim strResult As String
    
    If TheServerIndex = -1 Then
        Set TheServer = CurrentActiveServer
    Else
        Set TheServer = ActiveServers(TheServerIndex)
    End If
    
    On Error Resume Next
    'inform any loaded plugins
    For i = 0 To NumToPlugIn.Count - 1
        If Plugins(i).boolLoaded Then
            Plugins(i).objPlugIn.IRCAction Action
        End If
    Next i
    
    'if a channel was passed
    If LenB(Channel) > 0 Then
        'get its ID and store it in ChannelID variable
        ChannelId = TheServer.GetChanID(Channel)
    End If
    'Set current sphere variables
    ScriptNick = Nick1
    ScriptNick2 = Nick2
    ScriptChan = Channel
    Select Case Action
        Case ndIson
            ReDim names(0)
            names = Split(strText, " ")
            If frmOptions.lstBdyNk.ListCount > 0 And UBound(names) > 0 Then
            For i2 = 0 To frmOptions.lstBdyNk.ListCount - 1
                For i = 0 To UBound(names) - 1
                    
                        If LCase$(frmOptions.lstBdyNk.List(i2)) = LCase$(names(i)) Then
                                For intTemp = 0 To UBound(AllBuddies)
                                    If LCase$(AllBuddies(intTemp).Name) = LCase$(names(i)) Then
                                        AllBuddies(intTemp).isOnline = True
                                        For Pos = 0 To AllBuddies(intTemp).ServerCount
                                            If AllBuddies(intTemp).Servers(Pos) = TheServer.HostName Then
                                                Pos = -1
                                                Exit For
                                            End If
                                        Next Pos
                                        If Pos <> -1 Then
                                            AllBuddies(intTemp).AddServer TheServer.HostName
                                        End If
                                        i = -1
                                        Exit For
                                    End If
                                Next intTemp
                            End If
                            If i = -1 Then Exit For
                        Next i
                        If i <> -1 Then
                            For intTemp = 0 To UBound(AllBuddies)
                                If LCase$(AllBuddies(intTemp).Name) = LCase$(frmOptions.lstBdyNk.List(i2)) Then
                                    For Pos = 0 To AllBuddies(intTemp).ServerCount
                                        If AllBuddies(intTemp).Servers(Pos) = TheServer.HostName Then
                                            AllBuddies(intTemp).Servers(Pos) = vbNullString
                                        End If
                                    Next Pos
                                End If
                            Next intTemp
                        End If
                    Next i2
            End If
            If showison = True Then
                If UBound(names) <> -1 Then
                    For i = 0 To UBound(names) - 1
                        AddStatus EVENT_PREFIX & Language(48) & names(i) & " " & Language(551) & EVENT_SUFFIX & vbNewLine, TheServer
                    Next i
                Else
                    AddStatus IMPORTANT_PREFIX & Language(552) & IMPORTANT_SUFFIX & vbNewLine, TheServer
                End If
            Else
                For i = 1 To TheServer.Tabs.Tabs.Count
                    If TheServer.TabType(i) = TabType_Private Then
                        intTemp = 0
                        For i2 = 0 To UBound(names)
                            If TheServer.Tabs.Tabs(i).Caption = names(i2) Then
                                intTemp = 1
                            End If
                        Next i2
                        If intTemp = 0 Then
                            AddStatus IMPORTANT_PREFIX & Language(48) & TheServer.Tabs.Tabs(i).Caption & " " & Language(549) & IMPORTANT_SUFFIX & vbNewLine, TheServer, TheServer.GetChanID(TheServer.Tabs.Tabs(i).Caption)
                        End If
                    End If
                Next i

            End If
            If strCurrentPanel = "buddylist" Then
                wbPanel_DocumentComplete Nothing, vbNullString
            End If
        Case ndNamesSpecial
            AddStatus strText, TheServer
        Case ndBuddyName
            If Nick2 = "name" Then
                Set FS = New FileSystemObject
                If FS.FileExists(App.Path & "\conf\buddies\" & Nick1 & ".info") Then
                    intTemp = FreeFile
                    Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Input As #intTemp
                    i = FreeFile
                    Open App.Path & "/conf/buddies/" & Nick1 & "2.info" For Output As #i
                        Line Input #intTemp, strTemp
                        Print #i, "Real Name: " & strText
                        Line Input #intTemp, strTemp
                        Print #i, strTemp
                        Line Input #intTemp, strTemp
                        Print #i, strTemp
                        Line Input #intTemp, strTemp
                        Print #i, strTemp
                    Close #intTemp
                    Close #i
                    Kill App.Path & "/conf/buddies/" & Nick1 & ".info"
                    Set FS = New FileSystemObject
                    FS.MoveFile App.Path & "/conf/buddies/" & Nick1 & "2.info", App.Path & "/conf/buddies/" & Nick1 & ".info"
                Else
                    intTemp = FreeFile
                    Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Output As #intTemp
                        Print #intTemp, "Real Name: " & strText
                        Print #intTemp, "Version: x"
                        Print #intTemp, "Last Seen: x"
                        Print #intTemp, "Additional Info: x"
                    Close #intTemp
                End If
            End If
        
        'listing nicknames inside a channel
        Case ndNames
            If ChannelId = 0 Then
                IRCAction ndJoin, TheServer.myNick, Channel, , , TheServerIndex
                ChannelId = TheServer.GetChanID(Channel)
            End If
            'clear the names array
            ReDim names(0)
            'use Split to get the several names from the strText string
            names = Split(strText, " ")
            'add all names to the nicklist of the specified channel
            For i = 0 To UBound(names)
                If GetNickID(names(i), ChannelId, TheServer) = -1 Then
                    If names(i) <> vbNewLine Then
                        'only add him/her to the list if he/she is not already there
                        '(so we don't have double entries)
                        TheServer.NickList_List(ChannelId).Add names(i)
                    End If
                End If
            Next i
            SortCollection2 TheServer.NickList_List(ChannelId)
            'if that channel is selected...
            If ChannelId = TheServer.GetChanID(TheServer.Tabs.SelectedItem.Caption) Then
                '...update view
                buildStatus
            End If
            'don't move further
            Exit Sub
        Case ndNotice
            If Strings.Left$(strText, 1) = MIRC_CTCP Or Strings.Left$(strText, Len("dcc send")) = "DCC SEND" Then
                If Strings.Left$(strText, 4) = MIRC_CTCP & "DCC" Or Strings.Left$(strText, Len("dcc send")) = "DCC SEND" Then
                    TheServer.IPSender = Strings.Mid$(strText, InStrRev(strText, "(") + 1, Len(strText) - InStrRev(strText, "(") - 1)
                    'addStatus IMPORTANT_PREFIX & Nick1 & Language(155) & IMPORTANT_SUFFIX & vbnewline
                ElseIf Strings.Left$(strText, 8) = MIRC_CTCP & "VERSION" Then
                    AddStatus IMPORTANT_PREFIX & "> " & Replace(strText, MIRC_CTCP, vbNullString) & IMPORTANT_SUFFIX & vbNewLine, TheServer
                    strText = Replace(strText, MIRC_CTCP & "VERSION ", vbNullString)
                    strText = Replace(strText, MIRC_CTCP, vbNullString)
                    Set FS = New FileSystemObject
                    If FS.FileExists(App.Path & "/conf/buddies/" & Nick1 & ".info") Then
                        intTemp = FreeFile
                        Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Input As #intTemp
                        i = FreeFile
                        Open App.Path & "/conf/buddies/" & Nick1 & "2.info" For Output As #i
                            Do Until EOF(intTemp)
                                Line Input #intTemp, strTemp
                                If Strings.Left$(strTemp, 8) = "Version:" Then
                                    Print #i, "Version: " & strText
                                Else
                                    Print #i, strTemp
                                End If
                            Loop
                        Close #intTemp
                        Close #i
                        Kill App.Path & "/conf/buddies/" & Nick1 & ".info"
                        Set FS = New FileSystemObject
                        FS.MoveFile App.Path & "/conf/buddies/" & Nick1 & "2.info", App.Path & "/conf/buddies/" & Nick1 & ".info"
                    End If
                    Exit Sub
                ElseIf LCase$(Left$(strText, Len(MIRC_CTCP & "ping "))) = LCase$(MIRC_CTCP & "ping ") Then
                    strTemp = Right$(strText, Len(strText) - Len(MIRC_CTCP & "PING "))
                    i = ToTimeStamp - Left$(strTemp, Len(strTemp) - Len(MIRC_CTCP))
                    strTemp = Replace(Language(922), "%1", Nick1)
                    strTemp = Replace(strTemp, "%2", i & " " & Language(IIf(i = 1, 252, 251)))
                    AddStatus IMPORTANT_PREFIX & "> " & strTemp & IMPORTANT_SUFFIX & vbNewLine, TheServer

                Else
                    AddStatus IMPORTANT_PREFIX & "> " & Replace(strText, MIRC_CTCP, vbNullString) & IMPORTANT_SUFFIX & vbNewLine, TheServer
                End If
            ElseIf Strings.Left$(strText, 1) = "@" And Nick1 = TheServer.myNick Then
                'lag answer
                tmrLag(Strings.Right$(strText, Len(strText) - 1)).Enabled = False
                AddStatus IMPORTANT_PREFIX & "> Your current lag is " & tmrLag(Strings.Right$(strText, Len(strText) - 1)).Tag & " milliseconds" & IMPORTANT_SUFFIX & vbNewLine, TheServer
                
            Else
                If Options.ParseMemoServ Then
                    If LCase$(Nick1) = "memoserv" Then
                        If Channel = vbNullString Then
                            If ParseMemo(strText, TheServer) Then
                                'don't display the incoming notice
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                AddStatus IMPORTANT_PREFIX & strText & IMPORTANT_SUFFIX & vbNewLine, TheServer, IIf(LenB(Channel) > 0, ChannelId, -1)
                If Strings.Left$(strText, 8) = "DCC Chat" Then
                    TheServer.IPSender = Strings.Mid$(strText, InStr(1, strText, "(") + 1, Len(strText) - InStr(1, strText, "(") - 1)
                End If
            End If
            Exit Sub
        Case ndInvite
            'use Node special coloring codes for display
            If LenB(Nick2) = 0 Then
                AddStatus EVENT_PREFIX & Replace(Replace(Language(688), "%1", REASON_PREFIX & Nick1 & REASON_SUFFIX), "%2", HTML_OPEN & "a href=""NodeScript:/join " & Channel & """" & HTML_CLOSE & Channel & HTML_OPEN & "/a" & HTML_CLOSE) & EVENT_SUFFIX & vbNewLine, TheServer, , False
                If GetSetting(App.EXEName, "InfoTips", "Invite", "0") = "0" Then
                    strLastBalloonInfo = strText
                    ShowInfoTip Invitation
                    SaveSetting App.EXEName, "InfoTips", "Invite", "1"
                End If
                If Options.JoinOnInvite Then
                    TheServer.preExecute "/join " & Channel, False
                End If
                AddNews Replace(Replace(Language(688), "%1", Nick1), "%2", Channel)
            Else
                AddStatus EVENT_PREFIX & Replace(Replace(Replace(Language(881), "%1", REASON_PREFIX & Nick1 & REASON_SUFFIX), "%2", REASON_PREFIX & Nick2 & REASON_SUFFIX), "%3", REASON_PREFIX & Channel & REASON_SUFFIX) & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            End If
            RunScript "Invite"
            ThisSoundSchemePlaySound "invitation"
            Exit Sub
        Case ndChannelList
            'check to see if a list message was the one who called
            If strText <> "end of list" Then
                AddNews Language(479) & IIf(ChannelsCount = 0, "", " (" & ChannelsCount & ")")
                'empty the Names container
                ReDim names(0)
                strText = Strings.Mid$(strText, 2, Len(strText) - 1)
                names = Split(strText, " ", 3)
                On Error Resume Next
                names(2) = Strings.Right$(names(2), Len(names(2)) - 1)
                If Err.Number <> 0 Then
                    names(2) = Strings.Right$(names(2), Len(names(2)) - 1)
                End If
                If Not RetrievingChanList Then
                    If FS.FileExists(App.Path & "/temp/channels.html") Then
                        FS.DeleteFile App.Path & "/temp/channels.html"
                    End If
                    intTemp = 0
                    RetrievingChanList = True
                Else
                    If FS.FileExists(App.Path & "/temp/channels.html") Then
                        intTemp = 1
                    Else
                        intTemp = 2
                    End If
                End If
                If intTemp = 0 Then
                    TheServer.ChannelList_ReDim 0, False
                    ChannelsCount = 0
                End If
                intTemp = 1
                For i2 = 0 To UBound(names)
                    i = TheServer.ChannelList_UBound
                    If intTemp = 1 Then
                        TheServer.ChannelList_Channel(i) = names(i2)
                    ElseIf intTemp = 2 Then
                        TheServer.ChannelList_Users(i) = names(i2)
                    ElseIf intTemp = 3 Then
                        TheServer.ChannelList_Topic(i) = names(i2)
                    End If
                                   
                    If intTemp = 3 Then
                        intTemp = 1
                        ChannelsCount = ChannelsCount + 1
                        TheServer.ChannelList_ReDim i + 1
                    Else
                        intTemp = intTemp + 1
                    End If
                Next i2
                
                'ChannelsCount = ChannelsCount + 1
                'ReDim Preserve CurrentActiveServer.ChannelList(i + 1)
                Exit Sub
            Else
                ChannelsCount = 0
                AddNews Language(480)
                intTemp = 0
                On Error GoTo Invalid_ChanLst
                If CurrentActiveServer.ChannelList_UBound > 0 Then
                    Screen.MousePointer = vbHourglass
                    
                    'transfer data from the ChannelList Array to Collections
                    Set cChannels = New Collection
                    Set cTopics = New Collection
                    Set cUsers = New Collection
                    Set cCLSorting = New Collection
                    For i2 = 0 To CurrentActiveServer.ChannelList_UBound - 1
                        cChannels.Add CurrentActiveServer.ChannelList_Channel(i2)
                        cTopics.Add CurrentActiveServer.ChannelList_Topic(i2)
                        cUsers.Add CurrentActiveServer.ChannelList_Users(i2)
                    Next i2
                    If TheServer.bCLSorting <> 255 Then
                        Select Case TheServer.bCLSorting
                            Case 0
                                SortCollection3 cChannels, cTopics, cUsers
                            Case 1
                                SortCollection3 cTopics, cChannels, cUsers
                            Case 2
                                SortCollection3 cUsers, cChannels, cTopics
                        End Select
                    End If
                    
                    strResult = "<table>" & _
                                "<tr>" & _
                                "<td class='bhead' id='lang_chan'><a href='JavaScript:sortby(0);'>Channel</a></td>" & _
                                "<td class='bhead' id='lang_ppl'><a href='JavaScript:sortby(2);'>People</a></td>" & _
                                "<td class='bhead' id='lang_topic'><a href='JavaScript:sortby(1);'>Topic</a></td>" & _
                                "</tr>"
                    
                    Dim N As Long
                    Dim P As Long
                    N = GetTickCount
                    For i2 = 1 To cChannels.Count
                        strTemp = CreateMainText(HTML_OPEN & "tr" & HTML_CLOSE & _
                                HTML_OPEN & "td class=""blist""" & HTML_CLOSE & _
                                HTML_OPEN & _
                                "acronym title=""" & Replace(Language(702), "%1", cChannels(i2)) & """" & HTML_CLOSE & HTML_OPEN & "a href=""NodeScript:/join " & cChannels(i2) & """" & HTML_CLOSE & _
                                cChannels(i2) & HTML_OPEN & "/a" & HTML_CLOSE & HTML_OPEN & "/acronym" & HTML_CLOSE & _
                                HTML_OPEN & "/td" & HTML_CLOSE & HTML_OPEN & "td class=""blist""" & HTML_CLOSE, TheServer)
                        strTemp = strTemp & CreateMainText(SPECIAL_PREFIX & _
                                        cUsers(i2) & SPECIAL_SUFFIX & _
                                        HTML_OPEN & "/td" & HTML_CLOSE & HTML_OPEN & "td class=""blist""" & HTML_CLOSE, TheServer)
                        strTemp = strTemp & CreateMainText(REASON_PREFIX & cTopics(i2) & "&nbsp;" & REASON_SUFFIX & HTML_OPEN & "/td" & HTML_CLOSE & HTML_OPEN & "/tr" & HTML_CLOSE, TheServer)
                        If InStr(1, strTemp, "#") > 0 Then
                            strResult = strResult & strTemp
                        End If
                    Next i2
                    P = GetTickCount
                    
                    'MsgBox P - N
                    
                    strResult = strResult & "</table>"
                    
                    Screen.MousePointer = vbDefault
                    
                    'If FS.FileExists(App.Path & "/temp/channels2.html") Then
                    '    FS.DeleteFile App.Path & "/temp/channels2.html", True
                    'End If
                    'FS.MoveFile App.Path & "/temp/channels.html", App.Path & "/temp/channels2.html"
                    
                    Set frmChannels = New frmCustom
                    frmChannels.DialogData = strResult
                    frmChannels.DialogData2 = TheServerIndex
                    LoadDialog App.Path & "/data/dialogs/channels.xml", frmChannels
                    RetrievingChanList = False
                Else
Invalid_ChanLst:
                    'no such channels available
                    AddStatus Language(152) & vbNewLine, TheServer, , True
                    Resume
                End If
            End If
        Case ndBanList
            CurFile1 = FreeFile
            bolTemp = FS.FileExists(App.Path & "/temp/banlist.html")
            Open App.Path & "/temp/banlist.html" For Append Access Write Lock Write As #CurFile1
            If Not bolTemp Then
                CurFile2 = FreeFile
                Open App.Path & "/data/html/imports/chanproperties_intro.html" For Input Lock Write As #CurFile2
                PartCopyR CurFile1, CurFile2
                Close #CurFile2
            End If
                        
            If strText <> "end of list" Then
                BanListIndex = BanListIndex + 1
                Print #CurFile1, _
                    "<tr><td class=""blist"" id=""banmode" & BanListIndex & """>" & GetParameterQuick(strText, 0) & "</td>" & _
                        "<td class=""blist"">" & GetParameterQuick(strText, 1) & "</td>" & _
                        "<td class=""blist"">" & LocalTime(GetParameterQuick(strText, 2)) & "</td>" & _
                        "<td class=""blist""><input type=""button"" value=""Remove"" " & _
                        "onClick=""removeban(banmode" & BanListIndex & ");""></tr>"
                Close #CurFile1
                'ChanProps.wbCustom(1).Refresh2
            Else
                Print #CurFile1, "<div style=""display:none"" id=""chan"">" & Channel & "</div>"
                Print #CurFile1, "<script language=""JavaScript"">"
                Print #CurFile1, "channeltitle.innerText = chan.innerText;"
                Print #CurFile1, "newtopic.value = '" & Replace(Replace(TheServer.NickList_Topic(ChannelId), "'", "\'"), "\", "\\") & "';"
                Print #CurFile1, "</script>"
                CurFile2 = FreeFile
                Open App.Path & "/data/html/imports/chanproperties_end.html" For Input Lock Write As #CurFile2
                PartCopy CurFile1, CurFile2
                Close #CurFile2
                Close #CurFile1
                'ChanProps.wbCustom(1).Refresh2
            End If
            
            If Not ChanProps Is Nothing Then
                Unload ChanProps
            End If
            Set ChanProps = New frmCustom
            LoadDialog App.Path & "/data/dialogs/channel.xml", ChanProps
        
        Case ndTopic
            If Nick2 = "original" Then
                AddStatus EVENT_PREFIX & Language(140) & " " & REASON_PREFIX & strText & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            Else
                ThisSoundSchemePlaySound "topicchange"
                AddStatus EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(154) & " " & REASON_PREFIX & strText & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            End If
            TheServer.NickList_Topic(ChannelId) = strText
            
            Exit Sub
        Case ndTopicTime
            ReDim names(0)
            names = Split(strText, " ")
            strNick = names(2)
            strTemp2 = names(3)
            Channel = names(1)
            lTemp = Val(strTemp2)
            strTemp2 = LocalTime(lTemp)
            ChannelId = TheServer.GetChanID(Channel)
            AddStatus EVENT_PREFIX & Language(150) & REASON_PREFIX & strNick & REASON_SUFFIX & " " & Language(151) & " " & REASON_PREFIX & strTemp2 & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            Exit Sub
    End Select
    'if we did that action the program should act differently...
    If Nick1 = TheServer.myNick Then
        '...depending on the action we did
        Select Case Action
            Case ndJoin
                'We enter a new channel
                'only if there isn't such a tab, create one
                'go through the current tabs
                For i = 1 To TheServer.Tabs.Tabs.Count
                    'if this tab is the channel which we joined...
                    If Strings.LCase$(TheServer.Tabs.Tabs(i).Caption) = Strings.LCase$(Channel) Then
                        '...just show that we joined
                        GoTo Show_Join
                    End If
                    'go to the next existing tab
                Next i
                'There isn't such a tab.
                'We'll have to create one
                intTemp = TheServer.Tabs.Tabs.Add(, , Channel, TabImage_Channel).index
                'When we create a tab we should also create two items at the TabInfo and TabType collections
                'use ChannelID variable to store the new ChannelID
                'we will have to create a new textbox which index
                'will be an unused ChannelID.
                'so ChannelID will be txtStatus.Count
                TheServer.TabInfo.Add xLet(ChannelId, TheServer.GetEmptyChannelID)
                'the type of the tab is Channel
                TheServer.TabType.Add TabType_Channel
                'new textbox. refer from the one to the other using TabInfo(i.e. TabInfo(tsTabs.Tab.Item(index).Index) = txtStatus.Index)
                'After adding this tab we'll have to update the tabs bar.
                If TheServer Is CurrentActiveServer Then
                    UpdateTabsBar
                End If
                'Here a new irc-textbox is created and so is a NickList collection
                '(the one used to store the contents of the channel we join,
                ' the other to store the nicknames of the people who are in
                ' this channel)
                If ChannelId > TheServer.IRCData_Count - 1 Then
                    TheServer.IRCData_ReDim ChannelId
                End If
                TheServer.NickList_ReDim ChannelId
                TheServer.NickList_List_LoadNewCollection ChannelId
                'the tag of the nicklist(used to store the size of the nicklist)
                'will be set to its default value
                TheServer.NickList_Size(ChannelId) = NickList_DefaultSize
                'whereas the text/contents of the channel will
                'be nothing
                TheServer.IRCData(ChannelId) = vbNullString
                RememberChannel Channel
                If strCurrentPanel = "join" Then
                    tmrPanelRefreshSoon.Enabled = False
                    tmrPanelRefreshSoon.Enabled = True
                ElseIf strCurrentPanel = "favorites" Then
                    tmrPanelRefreshSoon.Enabled = False
                    tmrPanelRefreshSoon.Enabled = True
                End If
                SaveSession
                TheServer.Tabs.Tabs.Item(intTemp).Selected = True
Show_Join:
                If Options.ModesOnJoin Then
                    TheServer.preExecute "/mode " & Channel
                End If
                tvConnections.Nodes.Add TheServer.ServerNode, tvwChild, "c_" & TheServerIndex & "_" & Channel, Channel, TabImage_Channel
                '40 = "You have joined"
                'Display `You have joined that channel' message
                AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Replace(Language(689), "%1", Channel) & _
                      EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
                AddLog "-- " & Language(198) & " " & Channel & " " & Language(151) & " " & DateTime.Time & ", " & DateTime.Date & " --", ChannelId, GetServerIndexFromActiveServer(TheServer)
                'and run the join routine of the main script
                RunScript "Join"
                ThisSoundSchemePlaySound "join"
                If GetSetting(App.EXEName, "InfoTips", "Join", "0") = "0" Then
                    ShowInfoTip Joined
                    SaveSetting App.EXEName, "InfoTips", "Join", "1"
                End If
                AddNews Replace(Language(689), "%1", Channel)
                Exit Sub
            Case ndPart
                AddLog "-- " & Language(199) & " " & Channel & " " & Language(151) & " " & DateTime.Time & ", " & DateTime.Date & " --", ChannelId, GetServerIndexFromActiveServer(TheServer)
                'we leave a channel
                '41 = "You have left"
                'clear the nicknames listbox of the channel we left
                ClearList TheServer.NickList_List(ChannelId)
                'display `You have left that channel' message
                AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Replace(Language(690), "%1", Channel) & _
                      EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
                'remove the tab
                'Unload txtStatus(TabInfo(tsTabs.SelectedItem.index))
                TheServer.NickList_List_SetToNothing TheServer.TabInfo(TheServer.Tabs.SelectedItem.index)
                intTemp = TheServer.GetTab(ChannelId)
                TheServer.TabInfo.Remove intTemp
                TheServer.TabType.Remove intTemp
                TheServer.Tabs.Tabs.Remove intTemp
                TheServer.DCCTabRemove intTemp
                For i = 1 To tvConnections.Nodes.Count
                    If tvConnections.Nodes.Item(i).Key = "c_" & TheServerIndex & "_" & Channel Then
                        tvConnections.Nodes.Remove i
                        Exit For
                    End If
                Next i
                Set TheServer.Tabs.SelectedItem = TheServer.Tabs.Tabs.Item(TheServer.GetTab(TheServer.GetStatusID))
                'and run the appropriate script routine
                RunScript "Part"
                ThisSoundSchemePlaySound "part"
                AddNews Replace(Language(690), "%1", Channel)
                Exit Sub
            Case ndQuit
                AddLog "-- " & Language(124) & " " & Language(151) & " " & DateTime.Time & ", " & DateTime.Date & " --", ChannelId, GetServerIndexFromActiveServer(TheServer)
                'disconnected from IRC
                'completely close any soft-closed connection
                TheServer.WinSockConnection.Close
                'and run the quit routine of the main scripts
                RunScript "Quit"
                ThisSoundSchemePlaySound "quit"
                AddNews "-- " & Language(124) & " --"
                Exit Sub
            Case ndError
                'an error occured
                'Error
                '42 = "Error"
                'display it
                AddStatus EVENT_PREFIX & Language(42) & IIf(Not IsMissing(strText), "(" & REASON_PREFIX & strText & REASON_SUFFIX & ")", vbNullString) & EVENT_SUFFIX & vbNewLine, TheServer, False
                'run the correct routine
                RunScript "Error"
                ThisSoundSchemePlaySound "error"
                AddNews Language(691)
                Exit Sub
            Case ndKick
                'We have been kicked out of a channel
                '43 = "You have been kicked out of"
                '44 = "by"
                'clear the nicklist of that channel
                'as we would if we had part a channel
                ClearList TheServer.NickList_List(ChannelId)
                'show the kick message
                AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(43) & _
                      " " & Channel & " " & Language(44) & " " & Nick2 & "(" & REASON_PREFIX & strText & REASON_SUFFIX & ")" & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
                AddLog Language(43) & _
                      " " & Channel & " " & Language(44) & " " & Nick2 & "(" & strText & ")" & vbNewLine, ChannelId
                'and run the kick script
                RunScript "Kick"
                ThisSoundSchemePlaySound "kick"
                If GetSetting(App.EXEName, "InfoTips", "Kicked", "0") = "0" Then
                    strLastBalloonInfo = Channel
                    ShowInfoTip Kicked
                    SaveSetting App.EXEName, "InfoTips", "Kicked", "1"
                End If
                If Options.JoinOnKick Then
                    TheServer.preExecute "/join " & Channel
                End If
                AddNews Replace(Language(692), "%1", Channel)
                Exit Sub
            Case ndNick
                'our nick was changed(either by the server or by us using /nick)
                'store the new nick in myNick variable
                TheServer.myNick = Nick2
                'set the sphere variables in order to update myNick variable.
                SetSphereVariables
                'Save the New Nick to the config file
                'We will use it next time too
                'get a free file index
                CurFile1 = FreeFile
                'open the current userinfo data file
                Open App.Path & "\conf\info.dat" For Input As #CurFile1
                'get another free file index
                CurFile2 = FreeFile
                'create a temporary file; we will create the new info
                'there and then we will rename this to our current userinfo file
                'and delete the old one.
                Open App.Path & "\conf\info.tmp" For Output As #CurFile2
                'skip one line from the current info file
                xLineInput CurFile1
                'write the current nick in the temporary file
                Print #CurFile2, TheServer.myNick
                'copy all other lines from the original file to the temporary one.
                Do Until EOF(CurFile1)
                    Print #CurFile2, xLineInput(CurFile1)
                Loop
                'save the temporary file to the disk
                Close #CurFile2
                'close the original file
                Close #CurFile1
                'delete the original file
                Kill App.Path & "\conf\info.dat"
                'rename the temporary file to the actual user info one
                Name App.Path & "\conf\info.tmp" As App.Path & "\conf\info.dat"
                'send the necessary messages over NDC
                For i = 1 To UBound(NDCConnections)
                    If wsNDC.Item(i).State = sckConnected Then
                        If NDCConnections(i).IntroPackSent Then
                            wsNDC.Item(i).SendData "n" & TheServer.myNick & ":"
                        End If
                    End If
                Next i
                '"You are now known as "
                'Display `You are now know as someone' message
                AddStatus Replace(Language(693), "%1", Nick2) & vbNewLine, TheServer, , False
                AddNews Replace(Language(693), "%1", Nick2)
        End Select
    End If
    'general actions for irc events(if these actions where taken by us or others,
    'unless we used Exit Sub above(which we did almost for every event))
    
    'depending on the irc-action taken, we will take the appropriate actions(display messages, show/hide tabs, modify the nicklists, etc.)
    Select Case Action
        Case ndJoin
            'someone joins a channel
            'add his/her nickname to the list
            
            For i = 1 To TheServer.NickList_List(ChannelId).Count
                If TheServer.NickList_List(ChannelId).Item(i) = Nick1 Then
                    GoTo Nick_Already
                End If
            Next i
            
            CollectionAppend TheServer.NickList_List(ChannelId), Nick1
Nick_Already:
            'and display the `Someone joins a channel' message
            AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(46) & _
                  " " & Channel & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            AddLog Language(48) & Nick1 & " " & Language(46) & " " & Channel, ChannelId
            If frmOptions.lstBdyNk.ListCount > 0 Then
                For CheckBud = 0 To frmOptions.lstBdyNk.ListCount - 1
                    If LCase$(Nick1) = LCase$(frmOptions.lstBdyNk.List(CheckBud)) Then
                        If LenB(frmOptions.lstBdyWT.List(CheckBud)) > 0 Then
                            If Strings.Left$(frmOptions.lstBdyWT.List(CheckBud), 1) <> "/" Then
                                TheServer.SendData "privmsg " & Channel & " :" & frmOptions.lstBdyWT.List(CheckBud) & vbNewLine
                                AddStatus "&lt;" & TheServer.myNick & "&gt; " & frmOptions.lstBdyWT.List(CheckBud) & vbNewLine, TheServer, ChannelId
                            Else
                                TheServer.preExecute frmOptions.lstBdyWT.List(CheckBud)
                            End If
                            If Options.BuddyEnterMSG Then
                                AddStatus SPECIAL_PREFIX & Language(48) & Nick1 & Language(134) & Channel & _
                                          SPECIAL_SUFFIX & vbNewLine, TheServer
                                'preExecute ("/focus caption status")
                            End If
                        End If
                        If Options.BuddyEnterWIN Then
                            If Not (Me.WindowState = vbMinimized) Then
                                Set frmBuddy = New frmCustom
                                frmBuddy.DialogData = Nick1
                                LoadDialog App.Path & "/data/dialogs/buddysignon.xml", frmBuddy
                            End If
                        End If
                        If FS.FileExists(App.Path & "/conf/buddies/" & Nick1 & ".info") Then
                            intTemp = FreeFile
                            Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Input As #intTemp
                            i = FreeFile
                            Open App.Path & "/conf/buddies/" & Nick1 & "2.info" For Output As #i
                                Line Input #intTemp, strTemp
                                Print #i, strTemp
                                Line Input #intTemp, strTemp
                                Print #i, strTemp
                                Line Input #intTemp, strTemp
                                Print #i, "Last Seen: " & DateTime.Now
                                Line Input #intTemp, strTemp
                                Print #i, strTemp
                            Close #intTemp
                            Close #i
                            Kill App.Path & "/conf/buddies/" & Nick1 & ".info"
                            FS.MoveFile App.Path & "/conf/buddies/" & Nick1 & "2.info", App.Path & "/conf/buddies/" & Nick1 & ".info"
                        Else
                            intTemp = FreeFile
                            Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Output As #intTemp
                                Print #intTemp, "Real Name: " & Language(261)
                                Print #intTemp, "Version: " & Language(261)
                                Print #intTemp, "Last Seen: " & DateTime.Now
                                Print #intTemp, "Additional Info: "
                            Close #intTemp
                        End If
                        If Me.WindowState = vbMinimized Then
                            ShowInfoTip BuddySignOn
                        End If
                    End If
                Next CheckBud
                
                If LenB(frmOptions.lstBdyWT.List(CheckBud)) > 0 Then
                End If
            End If
            SortCollection2 TheServer.NickList_List(ChannelId)
            'inform any loaded plugins
            For i = 0 To NumToPlugIn.Count - 1
                If Plugins(i).boolLoaded And Not Plugins(i).objPlugIn Is Nothing Then
                    'error handler: plugins may include code with bugs
                    On Error Resume Next
                    Plugins(i).objPlugIn.SomeoneJoined Nick1, Channel
                End If
            Next i
            'don't forget to run the join script...
            RunScript "Join"
            ThisSoundSchemePlaySound "join"
        Case ndPart
            'someone leaves a channel
            'remove his/her nickname from the list
            TheServer.NickList_List(ChannelId).Remove GetNickID(Nick1, ChannelId, TheServer)
            'and display the `Someone leaves a channel' message
            AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(47) & _
                  " " & Channel & IIf(CBool(Len(strText)), "(" & REASON_PREFIX & strText & REASON_SUFFIX & ")", vbNullString) & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            AddLog Language(48) & Nick1 & " " & Language(47) & " " & Channel, ChannelId
            'run the appropriate script once more.
            RunScript "Part"
            ThisSoundSchemePlaySound "part"
        Case ndKick
            'some is being kicked either by someone else, or by us.
            'remove his/her nickname from the nicklist as if he left.
            On Error Resume Next 'no such nick(?)
            TheServer.NickList_List(ChannelId).Remove GetNickID(Nick1, ChannelId, TheServer)
            'display `Someone was kicked by someone_else' message
            AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(50) & _
                  " " & Channel & " " & Language(44) & " " & REASON_PREFIX & Nick2 & REASON_SUFFIX & "(" & REASON_PREFIX & strText & REASON_SUFFIX & ")" & EVENT_SUFFIX & vbNewLine, TheServer, ChannelId, False
            AddLog Language(48) & Nick1 & Language(50) & " " & Channel & " " & Language(44), ChannelId
            'and run the correct script one more time.
            RunScript "Kick"
            ThisSoundSchemePlaySound "kick"
        Case ndQuit
            'someone quits IRC.
            'that's the text which will be shown in every channel that person is in.
            strTemp = SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(49) & _
                      "(" & REASON_PREFIX & strText & REASON_SUFFIX & ")" & EVENT_SUFFIX & vbNewLine
            'this time we'll have to remove his/her nickname from all nickname lists.
            For i = 1 To TheServer.Tabs.Tabs.Count
                'it must be a channel
                If TheServer.TabType(i) = TabType_Channel Then
                    'we use GetNickID to get the index of a specific nickname
                    'inside the nicklist of the specified channel.
                    lTemp = GetNickID(Nick1, TheServer.TabInfo(i), TheServer)
                    'if the nickname exists on the channel
                    If lTemp <> -1 Then
                        'remove him/her.
                        TheServer.NickList_List(TheServer.TabInfo(i)).Remove lTemp
                        'display `That_person has quit IRC'
                        'on every channel he/she is in.
                        AddStatus strTemp, TheServer, TheServer.GetChanID(TheServer.Tabs.Tabs.Item(i).Caption), False
                        AddLog Language(48) & Nick1 & " " & Language(49) & "(" & strText & ")", TheServer.GetChanID(TheServer.Tabs.Tabs.Item(i).Caption), GetServerIndexFromActiveServer(TheServer)
                        Channel = TheServer.Tabs.Tabs.Item(i).Caption
                    End If
                End If
            Next i
            If frmOptions.lstBdyNk.ListCount > 0 Then
                For CheckBud = 0 To frmOptions.lstBdyNk.ListCount - 1
                    If Nick1 = frmOptions.lstBdyNk.List(CheckBud) Then
                        If Options.BuddyLeaveWIN Then
                            If Not (Me.WindowState = vbMinimized) Then
                                Set frmBuddy = New frmCustom
                                frmBuddy.DialogData = Nick1
                                LoadDialog App.Path & "/data/dialogs/buddysignoff.xml", frmBuddy
                            End If
                        End If
                        If Options.BuddyLeaveMSG Then
                            'preExecute ("/focus caption status")
                            AddStatus Language(48) & Nick1 & " " & Language(49) & vbNewLine, TheServer
                        End If
                        For i = 1 To TheServer.Tabs.Tabs.Count
                            If TheServer.TabType(i) = TabType_Private Then
                                If TheServer.Tabs.Tabs(i).Caption = Nick1 Then
                                    AddStatus IMPORTANT_PREFIX & Language(48) & TheServer.Tabs.Tabs(i).Caption & " " & Language(550) & IMPORTANT_SUFFIX & vbNewLine, TheServer, TheServer.GetChanID(TheServer.Tabs.Tabs(i).Caption)
                                End If
                            End If
                        Next i
                        Set FS = New FileSystemObject
                        If FS.FileExists(App.Path & "/conf/buddies/" & Nick1 & ".info") Then
                            intTemp = FreeFile
                            Open App.Path & "/conf/buddies/" & Nick1 & ".info" For Input As #intTemp
                            i = FreeFile
                            Open App.Path & "/conf/buddies/" & Nick1 & "2.info" For Output As #i
                                Do Until EOF(intTemp)
                                    Line Input #intTemp, strTemp
                                    If Strings.Left$(strTemp, 10) = "Last Seen:" Then
                                        Print #i, "Last Seen: " & DateTime.Now
                                    Else
                                        Print #i, strTemp
                                    End If
                                Loop
                            Close #intTemp
                            Close #i
                            Kill App.Path & "/conf/buddies/" & Nick1 & ".info"
                            Set FS = New FileSystemObject
                            FS.MoveFile App.Path & "/conf/buddies/" & Nick1 & "2.info", App.Path & "/conf/buddies/" & Nick1 & ".info"
                        End If
                        If Me.WindowState = vbMinimized Then
                            ShowInfoTip BuddySignOff
                        End If
                    End If
                Next CheckBud
            End If
            'and another script routine runs...
            RunScript "Quit"
            ThisSoundSchemePlaySound "quit"
        Case ndInitModes
            If strText <> "+" Then
                AddStatus SpecialSmiley("Arrow") & " " & Replace(Replace(Language(821), "%1", REASON_PREFIX & Channel & REASON_SUFFIX), "%2", REASON_PREFIX & strText & REASON_SUFFIX) & vbNewLine, TheServer, ChannelId, False
                AddLog Replace(Replace(Language(821), "%1", Channel), "%2", strText) & vbNewLine, ChannelId, TheServerIndex, Options.TimeStampChannels And Options.TimeStampLogs
                
                'let's see if these modes are stored
                'start from 2 -- ommitting the + symbol
                
                TheServer.NickList_Modes(ChannelId) = vbNullString
                
                strTemp = TheServer.NickList_Modes(ChannelId)
                For i = 2 To Len(strText)
                    strTemp2 = Mid$(strText, i, 1)
                    'check to see if the mode is already there
                    If InStr(1, strTemp, strTemp2) <= 0 Then
                        'it's not
                        'store it
                        TheServer.NickList_Modes(ChannelId) = TheServer.NickList_Modes(ChannelId) & strTemp2
                    End If
                Next i
            Else
                TheServer.NickList_Modes(ChannelId) = vbNullString
            End If
            If Not frmChanModes Is Nothing Then
                If frmChanModes.Visible Then
                    frmChanModes.DialogData = TheServer.NickList_Modes(ChannelId)
                    LoadDialog App.Path & "/data/dialogs/chanmodes.xml", frmChanModes
                End If
            End If
        Case ndMode
            'someone has changed the modes in a channel(headache :P)
            'first we display the modes set not worrying about WHAT was actually changed
            AddStatus SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(51) & " " & Channel & " " & Language(52) & " " & REASON_PREFIX & strText & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer, TheServer.GetChanID(Channel), False
            AddLog Language(48) & Nick1 & " " & Language(51) & " " & Channel & " " & Language(52) & " " & strText
            
            'Analysize Mode Changes
            strTemp2 = Strings.Left$(strText, InStr(1, strText, " ") - 1)
            'strTemp2 contains all modes(complete) e.g. +o-v
            strTemp = Replace(Replace(strTemp2, "+", vbNullString), "-", vbNullString)
            'strTemp contains all mode types e.g. ov
            'Len(strTemp) is the modes changes count
            ReDim AllModes(2, Len(strTemp)) 'resize and clear array
            'This array is used to store all mode-changes info.
            'Here's what we store in each place
            '(x is the mode-which-changed index)
            'AllModes(0, x) = ( + | - )
            'AllModes(1, x) = ( o | v | h | b )  ||  ( n | t | p | s | i | m | k | l | ( h | c | u | w | r ) )
            'AllModes(2, x) = [ Nickname | Hostname | Key | Limit ]
            'the count of the modes
            intTemp = Len(strTemp2) + 1
            'initialize the two counters...
            i = 0
            i2 = 0
            'now we go through the modes to determine which
            'ones are being set and which ones are being unset.
            'the resule is stored inside strTemp2
            Do
                'we move to the NEXT mode.
                i = i + 1
                i2 = i2 + 1
                'if there is no next mode we don't have to go through it ;)
                If i > intTemp Then Exit Do
                'if there's a + symbol we know that all further modes
                'are SET (unless we find another - symbol later)
                If Strings.Mid$(strTemp2, i2, 1) = "+" Then
                    'so... the current mode is SET, which means plus; true.
                    CurrentMode = True
                    'we don't need that symbol. The string strTemp2 will
                    'contain one symbol for each mode only.
                    'So we remove this leading symbol
                    strTemp2 = Strings.Left$(strTemp2, i - 1) & Strings.Right$(strTemp2, Len(strTemp2) - i)
                    'we move to the next character
                    i = i + 1
                    'but we didn't have to move to the next mode; we remove the 1 we added before.
                    i2 = i2 - 1
                'if there's a - symbol we know all further modes are UNSET
                'unless we find another + later...
                ElseIf Strings.Mid$(strTemp2, i, 1) = "-" Then
                    'so, the current mode is False this time.
                    CurrentMode = False
                    'remove the symbol again
                    strTemp2 = Strings.Left$(strTemp2, i - 1) & Strings.Right$(strTemp2, Len(strTemp2) - i)
                    i = i + 1
                    'and to the previous mode(the one we skipped)
                    i2 = i2 - 1
                Else 'mode letter
                    'that's a letter. This is the character
                    'that will be actually replaced by + or -
                    'Note: this Mid is NOT the same as Strings.Mid!
                    Mid$(strTemp2, i2, 1) = IIf(CurrentMode, "+", "-")
                End If
            Loop
            'now strTemp2 contains info about whether modes are set or not, for example +-
            'Len(strTemp2) must be Len(strTemp)
            'because strTemp2 contains if the modes are set or unset
            'whereas strTemp contains the mode letters.
            'the current nicklist is the nicklist of the current channel
            'it will be used to modify user oper and voice statuses depending
            'on the modes.
            
            'we need to save the set modes and remove the unset modes
            'from the channel's NickList_Modes
            
            Set CurrentNickList = TheServer.NickList_List(ChannelId)
            'Parameter can be invalid
            On Error Resume Next
            For i = 1 To Len(strTemp) 'go through all modes
                'and store everything inside AllModes.
                AllModes(0, i - 1) = Strings.Mid$(strTemp2, i, 1) '+ | -
                AllModes(1, i - 1) = xLet(strTemp3, Strings.Mid$(strTemp, i, 1)) 'o | v | h | b
                'if the current mode is one of these: ntsmipr
                'then it won't take any arguments
                If strTemp3 = "p" Or strTemp3 = "s" Or _
                   strTemp3 = "i" Or strTemp3 = "t" Or _
                   strTemp3 = "n" Or strTemp3 = "m" Or _
                   strTemp3 = "c" Or strTemp3 = "u" Or _
                   strTemp3 = "w" Or strTemp3 = "r" Then
                    'these modes don't take any parameters
                    intNoArgsModesCount = intNoArgsModesCount + 1
                    
                    '
                    'we need to store the new modes that are SET
                    'and remove the modes that are UNSET
                    'in .NickList_Modes
                    '
                    'using strNick as a temporary variable
                    strNick = TheServer.NickList_Modes(ChannelId)
                    If strTemp2 = "+" Then
                        'mode is being SET
                        'check if it's already there
                        If InStr(1, strNick, strTemp3) <= 0 Then
                            'it's not set
                            'we have to add it
                            TheServer.NickList_Modes(ChannelId) = strNick & strTemp3
                        End If
                    ElseIf strTemp2 = "-" Then
                        'mode is being UNSET
                        'check to see if it's there
                        If InStr(1, strNick, strTemp3) > 0 Then
                            'it's set
                            'we need to remove it
                            TheServer.NickList_Modes(ChannelId) = Replace(strNick, strTemp3, vbNullString)
                        End If
                    Else
                        'modes parameters, for example +l 200
                    End If
                
                    'TO DO:
                    ' +/- h mode has two meanings:
                    ' HalfOp a nickname, or Hidden Channel
                               
                'it isn't; we'll have to extract the arguments
                Else
                    'nick/mode parameter
                    'for example this will be SomeOne if the mode is +o Someone
                    'or it will be *!*@* if the mode is +b *!*@*
                    AllModes(2, i - 1) = GetParameter(strText, i - intNoArgsModesCount)
                End If
                'now we'll have to take the appropriate action
                'depending on the mode letter
                Select Case AllModes(1, i - 1)
                    'if it's o that means that...
                    Case "o"
                        '...someone is being oped/deoped
                        'strNick is the nick itself
                        'strNick2 is the nick including priviledges(@ or +)
                        'get them.
                        strNick = AllModes(2, i - 1)
                        i2 = GetNickID(strNick, ChannelId, TheServer)
                        If i2 = -1 Then
                            'this nickname is not on that channel
                            'but the mode change was accepted by the server
                            'this means that the person just joined
                            'but we didn't get the JOIN message yet
                            'add him/her to the list
                            CurrentNickList.Add GetNick(strNick)
                            i2 = CurrentNickList.Count
                        End If
                        strNick2 = CurrentNickList.Item(i2)
                        'now the action(set or unset) depends on what's in AllModes(0, x)
                        'that's either + or -
                        If AllModes(0, i - 1) = "+" Then '+o
                            'it's +, someone is being oped
                            '+ symbol is always on the beginning
                            'even if someone has an oper status
                            '(if someone has both op and voice
                            'the internal nicklist will store
                            ' +@Nickname)
                            'so check if he/she has already an operator priviledge
                            If Strings.Left$(strNick2, 3) <> "+%@" And Strings.Left$(strNick2, 2) <> "+@" And Strings.Left$(strNick2, 2) <> "%@" And Strings.Left$(strNick2, 1) <> "@" Then
                                'he/she hasn't. Gets promotion
                                'update the internal nicklist...
                                CurrentNickList.Remove i2
                                If Strings.Left$(strNick2, 2) = "+%" Then
                                    AddEntry "+%@" & Strings.Right$(strNick2, Len(strNick2) - 2), CurrentNickList, i2
                                ElseIf Strings.Left$(strNick2, 1) = "%" Then
                                    AddEntry "%@" & Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                                ElseIf Strings.Left$(strNick2, 1) = "+" Then
                                    AddEntry "+@" & Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                                Else
                                    AddEntry "@" & strNick2, CurrentNickList, i2
                                End If
                            End If
                            'we'll need to run the op script
                            ScriptNick2 = strNick2
                            RunScript "Op"
                            ThisSoundSchemePlaySound "op"
                        Else 'mode -o
                            'this time the mode is being unset
                            'deoped
                            'if has already an op
                            'that will happen if
                            '( Left$(strNick2, 2) = "+@" Or Left$(strNick2, 1) = "@" )
                            'is true.
                            If Strings.Left$(strNick2, 3) = "+%@" Then
                                'looses his/her @ but keeps his/her voice and halfop.
                                'we don't use GetNick so as the halfop doesn't disappear
                                CurrentNickList.Remove i2
                                AddEntry "+%" & Strings.Right$(strNick2, Len(strNick2) - 3), CurrentNickList, i2
                                CurrentNickList.List(i2) = "+%" & Strings.Right$(strNick2, Len(strNick2) - 2)
                            ElseIf Strings.Left$(strNick2, 2) = "+@" Then
                                'lose op but keeps his/her halfop
                                CurrentNickList.Remove i2
                                AddEntry "+" & Strings.Right$(strNick2, Len(strNick2) - 2), CurrentNickList, i2
                                CurrentNickList.List(i2) = "+" & Strings.Right$(strNick2, Len(strNick2) - 2)
                            ElseIf Strings.Left$(strNick2, 2) = "%@" Then
                                'lose op but keeps his/her halfop
                                CurrentNickList.Remove i2
                                AddEntry "%" & Strings.Right$(strNick2, Len(strNick2) - 2), CurrentNickList, i2
                                CurrentNickList.List(i2) = "%" & Strings.Right$(strNick2, Len(strNick2) - 2)
                            ElseIf Strings.Left$(strNick2, 1) = "@" Then
                                'loses op, but hasn't a voice status
                                CurrentNickList.Remove i2
                                AddEntry Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                            End If
                            'here comes the DeOp script routine
                            ScriptNick2 = strNick2
                            RunScript "DeOp"
                            ThisSoundSchemePlaySound "deop"
                        End If
                    Case "v"
                        'someone is being voiced/devoiced
                        'get the strNick and strNick2 again...
                        strNick = AllModes(2, i - 1)
                        i2 = GetNickID(strNick, ChannelId, TheServer)
                        strNick2 = CurrentNickList.Item(i2)
                        If AllModes(0, i - 1) = "+" Then '+v
                            'voiced
                            If Strings.Left$(strNick2, 1) <> "+" Then
                                'gets promotion
                                CurrentNickList.Remove i2
                                AddEntry "+" & strNick2, CurrentNickList, i2
                            End If
                            ScriptNick2 = strNick2
                            RunScript "Voice"
                            ThisSoundSchemePlaySound "voice"
                        Else 'mode -v
                            'devoiced
                            If Strings.Left$(strNick2, 1) = "+" Then
                                'looses his/her +
                                CurrentNickList.Remove i2
                                AddEntry Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                            End If
                            ScriptNick2 = strNick2
                            RunScript "DeVoice"
                            ThisSoundSchemePlaySound "devoice"
                        End If
                    Case "h"
                        'someone is being halfoped/dehalfoped
                        'get the strNick and strNick2 again...
                        strNick = AllModes(2, i - 1)
                        i2 = GetNickID(strNick, ChannelId, TheServer)
                        strNick2 = CurrentNickList.Item(i2)
                        If AllModes(0, i - 1) = "+" Then '+h
                            'halfoped
                            If Strings.Left$(strNick2, 1) <> "%" And Strings.Left$(strNick2, 2) <> "+%" Then
                                'gets promotion
                                CurrentNickList.Remove i2
                                If Strings.Left$(strNick2, 1) = "+@" Then
                                    AddEntry "+%@" & Strings.Right$(strNick2, Len(strNick2) - 2), CurrentNickList, i2
                                ElseIf Strings.Left$(strNick2, 1) = "+" Then
                                    AddEntry "+%" & Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                                ElseIf Strings.Left$(strNick2, 1) = "@" Then
                                    AddEntry "%@" & Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                                Else
                                    AddEntry "%" & strNick2, CurrentNickList, i2
                                End If
                            End If
                            ScriptNick2 = strNick2
                            RunScript "HalfOp"
                            ThisSoundSchemePlaySound "halfop"
                        Else 'mode -h
                            'dehalfoped
                            If Strings.Left$(strNick2, 1) = "%" Then
                                'looses his/her &
                                CurrentNickList.Remove i2
                                AddEntry Strings.Right$(strNick2, Len(strNick2) - 1), CurrentNickList, i2
                            ElseIf Strings.Left$(strNick2, 2) = "+%" Then
                                'looses his/her & but keep the +
                                CurrentNickList.Remove i2
                                AddEntry "+" & Strings.Right$(strNick2, Len(strNick2) - 2), CurrentNickList, i2
                            ElseIf Strings.Left$(strNick2, 3) = "+%@" Then
                                'looses his/her & but keep + and @
                                CurrentNickList.Remove i2
                                AddEntry "+@" & Strings.Right$(strNick2, Len(strNick2) - 3), CurrentNickList, i2
                            End If
                            ScriptNick2 = strNick2
                            RunScript "DeHalfOp"
                            ThisSoundSchemePlaySound "dehalfop"
                        End If
                    Case Else
                        'a different channel mode changes
                        ThisSoundSchemePlaySound "modechange"
                End Select
            'and we move to the next mode...
            Next i
            SortCollection2 CurrentNickList
            'if that happened in the current channel, show changes
            If ChannelId = TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) Then
                buildStatus
            End If
        Case ndPrivMsg
            'that's a new message...
            i = 0
            If InStr(1, strText, MIRC_CTCP & "ACTION") <> 0 Then
                If Nick1 = TheServer.myNick Then
                    i = 1
                End If
            End If
            If LenB(Channel) = 0 Or i = 1 Then
                '...a private message(there is no channel)
                'if the message starts with the mirc ctcp request character
                'handle it as a ctcp request.
                If Strings.Left$(strText, 1) = MIRC_CTCP And Strings.Right$(strText, 1) = MIRC_CTCP Then
                    strTemp = Replace(GetStatement(strText), MIRC_CTCP, vbNullString)
                    
                    ClearCTCPMemory
                    intTemp = UBound(PastCTCPs) + 1
                    ReDim Preserve PastCTCPs(intTemp)
                    PastCTCPs(intTemp).strNickname = Nick1
                    PastCTCPs(intTemp).strType = strTemp
                    PastCTCPs(intTemp).lngTime = GetTickCount
                    
                    If Options.CTCPFloodProtect Then
                        'check to see if we're being flooded
                        CheckCTCPFlood
                        
                        For i = 0 To UBound(ProtectedFrom)
                            If LCase$(ProtectedFrom(i).strNickname) = LCase$(Nick1) Then
                                'we are
                                'check to see if the reason is CTCP requests
                                If ProtectedFrom(i).bReason = 0 Then
                                    'yeap.
                                    'ignore CTCP request
                                    CTCPIgnoring = True
                                    
                                    'check to see if we need to bounce it back
                                    If Options.CTCPFloodBounce Then
                                        AddStatus Replace(Language(651), "%1", Nick1) & vbNewLine, TheServer
                                        TheServer.SendData "PRIVMSG " & Nick1 & " " & strText & vbNewLine
                                    End If
                                End If
                            End If
                        Next i
                    End If
                    
                    'CTCP REPLIES
                    Select Case strTemp
                        Case "VERSION"
                            'CTCP Version; we'll need to tell them that we are using Node!
                            'unless the user has disabled CTCP Version replies...
                            
                            'check to see if we are being flooded
                            If CTCPIgnoring Then
                                'yes, don't reply just display warning
                                AddStatus Replace(Language(649), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If
                            
                            If Options.CTCPVersion Then

                                If Options.CTCPVersionToIgnored Then
                                    For i = 0 To frmOptions.lstIgnore(0).ListCount - 1
                                        If LCase$(Nick1) = LCase$(frmOptions.lstIgnore(i).List(i)) Then
                                            'ignored nickname: ingore version request
                                            AddStatus Replace(Language(636), "%1", Nick1) & vbNewLine, TheServer
                                            Exit Sub
                                        End If
                                    Next i
                                End If
                                TheServer.SendData "Notice " & Nick1 & " :" & MIRC_CTCP & "VERSION " & Options.CTCPVersionMessage & MIRC_CTCP & vbNewLine
                                AddNews Replace(Language(694), "%1", Nick1)
                            Else
                                'we are ignoring ALL ping requests
                                AddStatus Replace(Language(635), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If
                        Case "TIME"
                            'check to see if we are being flooded
                            If CTCPIgnoring Then
                                'yes, don't reply just display warning
                                AddStatus Replace(Language(650), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If
                            
                            'they're asking for our local time... let them know.
                            If Options.CTCPTime Then
                                If Options.CTCPTimeToIgnored Then
                                    For i = 0 To frmOptions.lstIgnore(0).ListCount - 1
                                        If LCase$(Nick1) = LCase$(frmOptions.lstIgnore(i).List(i)) Then
                                            'ignored nickname: ingore time request
                                            AddStatus Replace(Language(638), "%1", Nick1) & vbNewLine, TheServer
                                            Exit Sub
                                        End If
                                    Next i
                                End If
                                TheServer.SendData "Notice " & Nick1 & " :" & MIRC_CTCP & "PING " & Replace(GetParameter(strText), MIRC_CTCP, vbNullString) & MIRC_CTCP & vbNewLine
                                AddNews Replace(Language(695), "%1", Nick1)
                            Else
                                'we are ignoring ALL time requests
                                AddStatus Replace(Language(637), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If
                            
                            TheServer.SendData "Notice " & Nick1 & " :" & MIRC_CTCP & "TIME " & Language(53) & " " & DateTime.Time & ", " & DateTime.Date & MIRC_CTCP & vbNewLine
                        Case "PING"
                            'are we still online? ofcourse we are!
                            
                            'check to see if we are being flooded
                            If CTCPIgnoring Then
                                'yes, don't reply just display warning
                                AddStatus Replace(Language(648), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If

                            'providing we are replying to ping requests...
                            If Options.CTCPPing Then
                                If Options.CTCPPingToIgnored Then
                                    For i = 0 To frmOptions.lstIgnore(0).ListCount - 1
                                        If LCase$(Nick1) = LCase$(frmOptions.lstIgnore(i).List(i)) Then
                                            'ignored nickname: ingore ping request
                                            AddStatus Replace(Language(634), "%1", Nick1) & vbNewLine, TheServer
                                            Exit Sub
                                        End If
                                    Next i
                                End If
                                TheServer.SendData "Notice " & Nick1 & " :" & MIRC_CTCP & "PING " & Replace(GetParameter(strText), MIRC_CTCP, vbNullString) & MIRC_CTCP & vbNewLine
                                AddNews Replace(Language(696), "%1", Nick1)
                            Else
                                'we are ignoring ALL ping requests
                                AddStatus Replace(Language(632), "%1", Nick1) & vbNewLine, TheServer
                                Exit Sub
                            End If
                        Case "ACTION"
                            If Nick1 = TheServer.myNick Then
                                i = TheServer.Tabs.SelectedItem.index
                                GoTo Show_Channel_Msg
                                Exit Sub
                            End If
                            For i = 1 To TheServer.Tabs.Tabs.Count
                                'if the caption of this tab is the nickname of the person
                                'who is talking to us and it is a private tab...
                                If Strings.LCase$(TheServer.Tabs.Tabs(i).Caption) = (Strings.LCase$(Nick1)) And TheServer.TabType(i) = TabType_Private Then
                                    '...display the message there
                                    GoTo Show_Channel_Msg
                                    Exit Sub
                                End If
                                
                                'go to the next tab
                            Next i
                            TheServer.Tabs.Tabs.Add , , Nick1, TabImage_Private
                            'the ChannelID of that tab will be i.
                            'i was used in the previous loop and
                            'it has the value of tsTabs.Tabs.Count + 1
                            '(i.e. one more than the last tab)
                            'that's perfect for a new tab!
                            TheServer.TabInfo.Add xLet(ChannelId, i)
                            'the new tab is a private message tab.
                            TheServer.TabType.Add TabType_Private
                            'we should update the tabsbar
                            If TheServer Is CurrentActiveServer Then
                                UpdateTabsBar
                            End If
                            'the messages are going to be stored in a new textbox
                            If ChannelId > TheServer.IRCData_Count + 1 Then
                                TheServer.IRCData_ReDim ChannelId
                            End If
                            'which initially mustn't contain anything at all.
                            TheServer.IRCData(ChannelId) = vbNullString
                            i = TheServer.Tabs.Tabs.Count
                            GoTo Show_Channel_Msg
                            Exit Sub
                        Case "NDC*"
                            'NDC pre-request. Send back normal NDC request.
                            'TO DO: Use resolved Local IP!
                            TheServer.preExecute "/PRIVMSG " & Nick1 & " :" & Strings.ChrW$(1) & "NDC " & LocalIP & Strings.ChrW$(1), False
                            'save the connection request details
                            'in the PendingNDCConnectionRequests array
                            'which we will use when receiving the IntroPack
                            'in order to determine the ActiveServer the user
                            'is connected to.
                            intTemp = UBound(PendingNDCConnectionRequests) + 1
                            ReDim PendingNDCConnectionRequests(intTemp)
                            Set PendingNDCConnectionRequests(intTemp).ActiveServer = TheServer
                            PendingNDCConnectionRequests(intTemp).Nickname = Nick1
                            AddNews Replace(Language(697), "%1", Nick1)
                            Exit Sub
                        Case "NDC"
                            'Normal NDC request. Establish NDC connection.
                            'intTemp = wsNDC.Count
                            intTemp = GetNDCFromNickname(Nick1)
                            If intTemp = -1 Then
                                intTemp = wsNDC.LoadNew
                            End If
                            wsNDC.Item(intTemp).LocalPort = 0
                            wsNDC.Item(intTemp).Close
                            strTemp = Right$(strText, Len(strText) - Len("#NDC "))
                            strTemp = Left$(strTemp, Len(strTemp) - 1)
                            wsNDC.Item(intTemp).Connect strTemp, 8752
                            ReDim Preserve NDCConnections(intTemp)
                            Set NDCConnections(intTemp).ActiveServer = TheServer
                            NDCConnections(intTemp).strNicknameA = Nick1
                            AddNews Replace(Language(698), "%1", Nick1)
                            Exit Sub
                        Case "DCC"
                            Dim names2() As String
                            ReDim names2(5)
                            ReDim names(5)
                            strText = Replace(strText, MIRC_CTCP, vbNullString)
                            names2() = Split(strText, " ")
                            names(0) = names2(0)
                            names(1) = names2(1)
                            names(2) = names2(2)
                            names2() = Split(StrReverse(strText), " ")
                            names(5) = StrReverse(names2(0))
                            names(4) = StrReverse(names2(1))
                            names(3) = StrReverse(names2(2))
                            'intTemp = Strings.Len("DCC ") + Strings.Len(names(1)) + 2
                            'lTemp = Strings.Len(names(5) & names(4) & names(3)) + 2
                            'names(2) = Strings.Mid$(strText, intTemp, Strings.Len(strText) - intTemp - lTemp)
                            'MsgBox "/" & names(2) & "/"
                            intTemp = wsDCC.Count
                            If names(1) = "SEND" Then
                                AddNews Replace(Language(700), "%1", Nick1)
                                TheServer.DCCTransfer_Incoming Nick1, names(2), names(5), names(4)
                                'addStatus Nick1 & " sending you a file" & vbnewline
                            ElseIf names(1) = "ACCEPT" Then
                                TheServer.DCCTransfer_Accepted Nick1, names(2), names(4), names(5)
                                '2-filename 4-port 5-size
                                Exit Sub
                            ElseIf names(1) = "RESUME" Then
                                MsgBox names(3) & " " & names(4) & " " & names(5)
                                TheServer.DCCTransfer_Resuming Nick1, names(4), CLng(Val(names(5)))
                                Exit Sub
                            ElseIf names(1) = "CHAT" Then
                                sbar.Panels.Item(1).Text = Replace(Language(701), "%1", Nick1)
                                strTemp = MsgBox(Replace(Language(701), "%1", Nick1), vbYesNo Or vbQuestion, Language(205))
                                If strTemp = vbNo Then
                                    AddStatus IMPORTANT_PREFIX & Language(241) & IMPORTANT_SUFFIX & vbNewLine, TheServer
                                    Exit Sub
                                Else
                                    TheServer.DCCChat_Accepted Nick1, names(5), TheServer.IPSender
                                    Exit Sub
                                End If
                            End If
                    End Select
                    AddStatus EVENT_PREFIX & "CTCP by " & Nick1 & ": " & REASON_PREFIX & Replace(GetStatement(strText), MIRC_CTCP, vbNullString) & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer
                    Exit Sub 'do not display it as a private message
                End If
                
                If frmOptions.lstIgnore(0).ListCount > 0 Then
                    For Pos = 0 To frmOptions.lstIgnore(0).ListCount - 1
                        If Strings.LCase$(Nick1) = Strings.LCase$(frmOptions.lstIgnore(0).List(Pos)) Then
                            Exit Sub
                        End If
                    Next Pos
                End If
                
                'Private Message(Query/Whisper)
                'see if the private message tab exists
                'go through all tabs
                For i = 1 To TheServer.Tabs.Tabs.Count
                    'if the caption of this tab is the nickname of the person
                    'who is talking to us and it is a private tab...
                    If Strings.LCase$(TheServer.Tabs.Tabs(i).Caption) = (Strings.LCase$(Nick1)) And TheServer.TabType(i) = TabType_Private Then
                        'Need to go through dccchats to see if any are linked to this tab
                        If TheServer.DCCChats_Count > 1 Then
                            For i2 = 1 To TheServer.DCCChats_Count - 1
                                If LCase$(TheServer.DCCChats_UserName(i2)) = LCase$(Nick1) Then
                                    If TheServer.DCCChats_TabIndex(i2) <> i Then
                                        Exit For
                                    End If
                                End If
                            Next i2
                        Else
                            i2 = -1
                        End If
                        If i2 < TheServer.DCCChats_Count - 1 Then
                            'check to see if an NDC connection is present
                            intTemp = GetNDCFromNickname(TheServer.Tabs.Tabs(i).Caption)
                            If intTemp <> -1 Then
                                NDCConnections(intTemp).Typing = False
                                If TheServer.Tabs.SelectedItem.index = i Then
                                    sbar.Panels.Item(1).Text = Replace(Language(687), "%1", NDCConnections(intTemp).strNicknameA)
                                End If
                            End If
                            '...display the message there
                            GoTo Show_Private_Msg
                        Else
                            GoTo Show_Private_Msg
                        End If
                    End If
                    'go to the next tab
                Next i
                'there isn't such a tab. Create one.
                'add a tab with its caption set to the nickname of the person we are talking to.
                TheServer.Tabs.Tabs.Add , , Nick1, TabImage_Private
                'the ChannelID of that tab will be i.
                'i was used in the previous loop and
                'it has the value of tsTabs.Tabs.Count + 1
                '(i.e. one more than the last tab)
                'that's perfect for a new tab!
                TheServer.TabInfo.Add xLet(ChannelId, TheServer.GetEmptyChannelID)
                'the new tab is a private message tab.
                TheServer.TabType.Add TabType_Private
                tvConnections.Nodes.Add(TheServer.ServerNode, tvwChild, "p_" & GetServerIndexFromActiveServer(TheServer) & "_" & Nick1, Nick1, TabImage_Private).Parent.Expanded = True
                ThisSoundSchemePlaySound "whisper"
                If Options.NarrationInterface Then
                    MSSpeech.Speak Replace(Language(724), "%1", Nick1)
                End If
                'we should update the tabsbar
                If TheServer Is CurrentActiveServer Then
                    UpdateTabsBar
                End If
                'the messages are going to be stored in a new textbox
                If ChannelId > TheServer.IRCData_Count - 1 Then
                    TheServer.IRCData_ReDim ChannelId
                End If
                'which initially mustn't contain anything at all.
                TheServer.IRCData(ChannelId) = vbNullString
                SaveSession
                If GetSetting(App.EXEName, "InfoTips", "Private", "0") = "0" Then
                    ShowInfoTip PrivateMessage
                    SaveSetting App.EXEName, "InfoTips", "Private", "1"
                End If
Show_Private_Msg:
                'flash window
                FlashWindow Me.hwnd, 0
                
                'notify the plugins
                For i2 = 0 To NumToPlugIn.Count - 1
                    If Plugins(i2).boolLoaded Then
                        If Not Plugins(i2).objPlugIn Is Nothing Then
                            On Error Resume Next
                            Plugins(i2).objPlugIn.Receiving Nick1, strText
                        End If
                    End If
                Next i2
                
                'now, display that message on the(existing or new) private message tab.
                'in the format of <Nick> Message!
                
                'IRCAction is used to display
                'our messages as well. This
                'special call is made with the
                'Nick2 argument set to MyNick
                'Therefore, we have to check
                'if Nick2 is set, in order
                'to display our nick before
                'the text.
                
                If LenB(Nick2) > 0 Then
                    strNick = Nick2
                Else
                    strNick = Nick1
                End If
                
                AddLog "&lt; " & strNick & " &gt; " & strText & vbNewLine, TheServer.TabInfo(i), TheServerIndex, Options.TimeStampLogs And Options.TimeStampPrivates
                
                If LenB(Nick2) > 0 Then
                    'make nickname link if enabled
                    If Options.NickLinkMinePriv Then
                        strNick = NickLink(strNick)
                    End If
                Else
                    'make nickname link if enabled
                    If Options.NickLinkPriv Then
                        strNick = NickLink(strNick)
                    End If
                End If
                
                AddStatus Replace(Options.DisplayNormal, "%nick", strNick) & " " & strText & vbNewLine, TheServer, TheServer.TabInfo(i), False, False
                AddNarration Replace(Language(723), "%1", strNick) & "; " & strText, TheServer, TheServer.TabInfo(i)
                LastMessage = strText
                RunScript "PrivMsg"
            Else
                'message in a channel
                'see if we have that channel tab open.
                'go through the tabs...
                For i = 1 To TheServer.Tabs.Tabs.Count
                    'if the caption of that tab is the same as the channel and it is a channel...
                    If Strings.LCase$(TheServer.Tabs.Tabs(i).Caption) = Strings.LCase$(Channel) And TheServer.TabType(i) = TabType_Channel Then
                        '...display it there.
                        GoTo Show_Channel_Msg
                    End If
                    'go to the next tab
                Next i
                'the channel wasn't found
                'create a new tab with its caption set to the channel name
                TheServer.Tabs.Tabs.Add , , Channel, TabImage_Channel
                'i is suitable for a new tab's ChannelID(it's tsTabs.Tabs.Count + 1)
                'add TabInfo data
                TheServer.TabInfo.Add xLet(ChannelId, TheServer.GetEmptyChannelID)
                'the type of the tab is Channel.
                TheServer.TabType.Add TabType_Channel
                'update the tabsbar
                If TheServer Is CurrentActiveServer Then
                    UpdateTabsBar
                End If
                'load a new nicklist and an irc-textbox
                If ChannelId > TheServer.IRCData_Count - 1 Then
                    TheServer.IRCData_ReDim ChannelId
                End If
                TheServer.NickList_ReDim ChannelId
                TheServer.NickList_List_LoadNewCollection ChannelId
                'the tag of the listbox containing the nicknames
                'indicates its width... set it to the default value.
                TheServer.NickList_Size(ChannelId) = NickList_DefaultSize
                'the text initially contains nothing.
                TheServer.IRCData(ChannelId) = vbNullString
Show_Channel_Msg:
                'only if the channel is selected
                If TheServer.Tabs.SelectedItem.index = i Then
                    'flash window
                    'TO DO: Add option about flashing.
                    FlashWindow Me.hwnd, 0
                End If
                'display the message...
                If Strings.Left$(strText, 1) = "" Then
                    strText = Replace(strText, "", vbNullString)
                    Pos = InStr(1, strText, " ")
                    
                    strNick = Nick1
                    'make nick a link if enabled
                    If strNick = TheServer.myNick Then
                        If Options.NickLinkMineChan Then
                            strNick = NickLink(strNick)
                        End If
                    Else
                        If Options.NickLinkChan Then
                            strNick = NickLink(strNick)
                        End If
                    End If
                    
                    AddStatus HTML_OPEN & "font color=#FF00FF" & HTML_CLOSE & HTML_OPEN & "strong" & HTML_CLOSE & _
                              "* " & strNick & Strings.Mid$(strText, Pos) & HTML_OPEN & "/strong" & HTML_CLOSE & HTML_OPEN & "/font" & HTML_CLOSE & vbNewLine, TheServer, TheServer.TabInfo(i), False, False
                    AddLog Nick1 & Strings.Mid$(strText, Pos), TheServer.TabInfo(i), GetServerIndexFromActiveServer(TheServer), Options.TimeStampLogs And Options.TimeStampChannels
                    AddNarration Nick1 & Strings.Mid$(strText, Pos), TheServer, TheServer.TabInfo(i)
                Else
                    'notify the plugins
                    For i2 = 0 To NumToPlugIn.Count - 1
                        If Plugins(i2).boolLoaded Then
                            If Not Plugins(i2).objPlugIn Is Nothing Then
                                On Error Resume Next
                                Plugins(i2).objPlugIn.Receiving Nick1, strText, Channel
                            End If
                        End If
                    Next i2
                    
                    strNick = Nick1
                    'make nick a link if enabled
                    If strNick = TheServer.myNick Then
                        If Options.NickLinkMineChan Then
                            strNick = NickLink(strNick)
                        End If
                    Else
                        If Options.NickLinkChan Then
                            strNick = NickLink(strNick)
                        End If
                    End If
                    'AddStatus "&lt; " & strNick & " &gt; " & strText & vbnewline, TheServer, TheServer.TabInfo(i), False, False
                    AddStatus Replace(Options.DisplayNormal, "%nick", strNick) & " " & strText & vbNewLine, TheServer, TheServer.TabInfo(i), False, False
                    AddLog "&lt; " & Nick1 & " &gt; " & strText & vbNewLine, TheServer.TabInfo(i), TheServerIndex, Options.TimeStampLogs And Options.TimeStampChannels
                    AddNarration Replace(Language(723), "%1", Nick1) & "; " & strText, TheServer, TheServer.TabInfo(i)
                End If
            End If
            LastMessage = strText
            RunScript "ChanMsg"
        Case ndNick
            'someone changes his/her nick(this can be us or someone else)
            'that's the text displayed on every channel that person is in.
            strTemp = SpecialSmiley("Arrow") & " " & EVENT_PREFIX & Language(48) & REASON_PREFIX & Nick1 & REASON_SUFFIX & " " & Language(54) & " " & REASON_PREFIX & Nick2 & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine
            'go through all tabs
            For i = 1 To TheServer.Tabs.Tabs.Count
                'if that's a channel...
                If TheServer.TabType(i) = TabType_Channel Then
                    '...it's a channel, check if the nick is in the nicklist of it
                    'get the nicklist and store it in CurrentNickList object variable
                    Set CurrentNickList = TheServer.NickList_List(TheServer.TabInfo(i))
                    'go through the nicknames of this nicklist
                    For i2 = 1 To CurrentNickList.Count
                        'if the nickname exists(the case doesn't matter)
                        If Strings.LCase$(GetNick(CurrentNickList.Item(i2))) = Strings.LCase$(Nick1) Then
                            strTemp2 = GetPriviledges(CurrentNickList.Item(i2))
                            'found nick name, replace it
                            CurrentNickList.Remove i2
                            AddEntry strTemp2 & Nick2, CurrentNickList, i2
                            SortCollection2 CurrentNickList
                            'and show msg.
                            AddStatus strTemp, TheServer, TheServer.TabInfo(i), False
                            AddLog Language(48) & Nick1 & " " & Language(54) & " " & Nick2, TheServer.TabInfo(i)
                            'then, skip the rest of the nicknames
                            '(there can't be the same nickname there later)
                            Exit For
                        End If
                        'move to the next nickname
                    Next i2
                ElseIf TheServer.TabType(i) = TabType_Private Then
                    If Strings.LCase$(TheServer.Tabs.Tabs.Item(i).Caption) = Strings.LCase$(Nick1) Then
                        TheServer.Tabs.Tabs.Item(i).Caption = Nick2
                    End If
                End If
                'move to the next channel
            Next i
            For i = 1 To tvConnections.Nodes.Count
                If tvConnections.Nodes.Item(i).Key = "p_" & TheServerIndex & "_" & Nick1 Then
                    tvConnections.Nodes.Item(i).Key = "p_" & TheServerIndex & "_" & Nick2
                    tvConnections.Nodes.Item(i).Text = Nick2
                End If
            Next i
            'run the appropriate script routine.
            RunScript "Nick"
        Case ndUseAltNick
            'TO DO: DOESN'T WORK. Check + Fix
            '
            'TO DO: CONVERT TO MULTIPLE SERVERS FRAMEWORK
            If TheServer.AltNickVal = 2 Then
                AddStatus Language(139) & vbNewLine, TheServer, , False
                frmOptions.txtAltTwo.Text = InputBox(Language(148), Language(149))
                'Exit Sub
            Else
                TheServer.AltNickVal = TheServer.AltNickVal + 1
            End If
            TheServer.myNick = Switch(TheServer.AltNickVal = 1, frmOptions.txtAlt.Text, TheServer.AltNickVal >= 2, frmOptions.txtAltTwo.Text)
            
            'preExecute "/quit"
            Wait 1
        
            TheServer.SendData "NICK " & TheServer.myNick & vbNewLine
        Case ndSpecialNotice
            AddStatus SpecialSmiley("Arrow") & " " & REASON_PREFIX & strText & REASON_SUFFIX & vbNewLine, TheServer, ChannelId
    End Select
End Sub
Public Sub DataArrival(ByVal strData As String, ByVal TheServerIndex As Integer)
'this sub is called by the event-sub wsIRC_DataArrival
'The argument strData contains a single line of incoming data

    Dim Nick1 As String, Reason As String, mode As String 'temporary variables used to store a nickname, a reason and a mode modification
    Dim chan As String, Nick2 As String 'two more
    Dim Pos As Integer, pos2 As Integer, pos3 As Integer 'temporary variables used to store positions inside a string
    Dim t As Variant, c As Variant, e As Variant 'three temporary variables used to store array results from Split() function
    Dim msgID As String 'msgID variable, used to identify the type of the incoming data
    Dim msgString As String 'msgString, used to store the data between the second and the third space inside the incoming data
    Dim realData As String 'realData, used to store the data after the second : symbol
    Dim strMsgDisplay As String 'strMsgDisplay, used to store the message that will be displayed in the Status window
    Dim IsChannel As Boolean 'boolean variable used to determine if the current action is taken in a channel(works for ALMOST every message, NOT for every message)
    Dim i As Integer
    Dim names() As String
    Dim TheServer As clsActiveServer
    
    Set TheServer = ActiveServers(TheServerIndex)
    
    'remove the comment from this line to be able to see the data
    'that the server actually sent.
    'DO NOT UNCOMMENT THE FOLLOWING LINE
    'use /debug instead!
    'AddStatus strData & vbnewline, TheServer 'check messages
    'let windows update
    'DoEvents
    'get the first space in the incoming data
    Pos = InStr(1, strData, " ")
    'get the second space
    pos2 = InStr(Pos + 1, strData, " ")
    'if both spaces exist
    If Pos > 0 And pos2 > 0 Then
        'get msgID which is the string between the first and the second space
        msgID = Strings.Mid$(strData, Pos + 1, pos2 - Pos - 1)
        'get the third space
        pos3 = InStr(pos2 + 1, strData, " ")
        'if there's no third space
        If pos3 <= 0 Then
            'consider the end of the message as third space
            pos3 = Len(strData)
        End If
        'get the msgString which is the string between the second and the third space
        'or the second space and the end of the data
        msgString = Strings.Mid$(strData, pos2 + 1, pos3 - pos2 - 1)
    End If
    'if the symbol : exists for second time in the message and the messageID is not 005 or 322...
    If InStr(2, strData, ":") > 0 And msgID <> "005" And msgID <> "322" Then
        'get the realData which is the data between the second : symbol and the end of the message
        'but not including the last <CR><LF> symbols
        If InStr(1, strData, vbNewLine) > 0 Then
            '<CR><LF>
            realData = Strings.Mid$(strData, InStr(2, strData, ":") + 1, Len(strData) - InStr(2, strData, ":") - 2)
        Else
            'Single <CR> or single <LF>
            realData = Strings.Mid$(strData, InStr(2, strData, ":") + 1, Len(strData) - InStr(2, strData, ":") - 1)
        End If
        '(if you want the last two <CR><LF> symbols to be included use this string parsing code
        'Right$(strData, Len(strData) - InStr(2, strData, ":"))
        'to do it)
    Else
        'if there's no second : symbol or if the message is 005 or 322...
        'the realData is the string between the second space and
        'the end of the message
        realData = Strings.Mid$(strData, pos2 + 1, Len(strData) - pos2 - 2)
    End If
    'get the strMsgDisplay string, the data that is going to be displayed
    'which is the same as realData
    strMsgDisplay = realData
    'if realData starts with a # it must be a channel
    IsChannel = Strings.Left$(realData, 1) = "#"
    Select Case msgID
        Case "JOIN"
            If IsChannel Then
                Pos = InStr(1, strData, "!")
                If Pos > 0 Then
                    Nick1 = Strings.Mid$(strData, 2, Pos - 2)
                Else
                    'a server, for example something.cool.org
                    Nick1 = Strings.Left$(strData, InStr(1, strData, " ") - 1)
                End If
                IRCAction ndJoin, Nick1, realData, , , TheServerIndex
            End If
        Case "MODE"
            'channel modes
            If Strings.Left$(msgString, 1) = "#" Then
                t = Split(strData, "!", 2)
                t = Split(t(0), " ", 2)
                Nick1 = Replace(CStr(t(0)), ":", vbNullString)
                c = Split(strData, "#", 2)
                e = Split(c(1), " ", 2)
                mode = Strings.Left$(e(1), Len(e(1)) - IIf(Right$(e(1), 2) = vbNewLine, 2, 1))
                IRCAction ndMode, Nick1, "#" & e(0), , mode, TheServerIndex
            Else
                'to do: user modes
                'for example:
                ':NodeUser MODE NodeUser :+wx
            End If
        Case "QUIT"
            Pos = InStr(1, strData, "!")
            Nick1 = Strings.Mid$(strData, 2, Pos - 2)
            Pos = InStr(1, strData, "QUIT") + Len("QUIT") + 2
            Reason = Strings.Mid$(strData, Pos, Len(strData) - Pos - 1)
            IRCAction ndQuit, Nick1, , , Reason, TheServerIndex
        Case "PART"
            Pos = InStr(1, strData, "!")
            Nick1 = Strings.Mid$(strData, 2, Pos - 2)
            If IsChannel Then
                'no part reason
                IRCAction ndPart, Nick1, Strings.Left$(msgString, Len(msgString) - 1), , , TheServerIndex
            Else
                IRCAction ndPart, Nick1, msgString, , realData, TheServerIndex
            End If
        Case "NICK"
            pos2 = InStr(2, strData, ":")
            Pos = InStr(1, strData, "!")
            Nick1 = Strings.Mid$(strData, 2, Pos - 2)
            Nick2 = Strings.Mid$(strData, pos2 + 1, Len(strData) - pos2 - 2)
            IRCAction ndNick, Nick1, , Nick2, , TheServerIndex
        Case "KICK"
            Pos = InStr(1, strData, "!")
            Nick2 = Strings.Mid$(strData, 2, Pos - 2)
            pos2 = InStr(1, strData, "#")
            chan = Strings.Mid$(strData, pos2 + 1, InStr(pos2, strData, " ") - pos2 - 1)
            Nick1 = Strings.Mid$(strData, pos2 + Len(chan) + 2, InStr(pos2 + Len(chan) + 3, strData, " ") - (pos2 + Len(chan) + 2))
            Reason = Strings.Mid$(strData, pos2 + Len(chan) + 2 + Len(Nick1) + 2, Len(strData) - (pos2 + Len(chan) + 2 + Len(Nick1) + 2) - 1)
            IRCAction ndKick, Nick1, "#" & chan, Nick2, Reason, TheServerIndex
        Case "PRIVMSG"
            pos2 = InStr(1, strData, "PRIVMSG " & TheServer.myNick)
            Pos = InStr(1, strData, "!")
            Nick1 = Strings.Mid$(strData, 2, Pos - 2)
            If Strings.Left$(msgString, 1) = "#" Then
                If frmOptions.lstIgnore(0).ListCount > 0 Then
                    For Pos = 0 To frmOptions.lstIgnore(0).ListCount - 1
                        If Strings.LCase$(Nick1) = Strings.LCase$(frmOptions.lstIgnore(0).List(Pos)) Then
                            Exit Sub
                        End If
                    Next Pos
                End If
                'Channel Message
                Pos = pos2 + Len("PRIVMSG #") - 1
                chan = msgString
                IRCAction ndPrivMsg, Nick1, chan, , realData, TheServerIndex
            Else
                'Private Message
                IRCAction ndPrivMsg, Nick1, , , realData, TheServerIndex
                'set the public var to the message
                MsgRcvTxt = realData
            End If
        Case "TOPIC" 'Topic changed
            Pos = InStr(1, strData, "#")
            chan = Strings.Mid$(strData, Pos, InStr(Pos, strData, " ") - Pos)
            Nick1 = Strings.Mid$(strData, 2, InStr(1, strData, "!") - 2)
            If InStr(1, Nick1, "services.") > 0 Then Nick1 = "ChanServ"
            IRCAction ndTopic, Nick1, chan, "changed", realData, TheServerIndex
        Case "NOTICE" 'A notice came...
            If InStr(1, strData, "!") < InStr(2, strData, " :") And InStr(1, strData, "!") > 0 Then
                Nick1 = Strings.Mid$(strData, 2, InStr(1, strData, "!") - 2)
            Else
                Nick1 = vbNullString
            End If
            Pos = InStr(InStr(2, strData, " ") + 1, strData, " ")
            pos2 = InStr(Pos + 1, strData, " :")
            chan = Mid$(strData, Pos + 1, pos2 - Pos - 1)
            chan = Replace(chan, "@", vbNullString)
            chan = Replace(chan, "+", vbNullString)
            chan = Replace(chan, "%", vbNullString)
            If Left$(chan, 1) <> "#" Then
                'the notice came to a nick -- us (or AUTH before using USER)
                chan = vbNullString
            'Else
                'the notice case to a channel
            End If
            IRCAction ndNotice, Nick1, chan, , realData, TheServerIndex
        
        Case "INVITE" 'You've been invited to a channel
            Nick1 = Mid$(strData, 2, InStr(1, strData, "!") - 2)
            IRCAction ndInvite, Nick1, realData, , , TheServerIndex
        Case "=" 'listing nicks(alternative)
            ReDim names(0)
            Pos = InStr(1, strData, "#")
            chan = Strings.Mid$(strData, Pos, InStr(Pos, strData, " ") - Pos)
            If namesboolean = False Then
                IRCAction ndNames, , chan, , Strings.Right$(strData, Len(strData) - InStr(2, strData, " :") - 1), TheServerIndex
            Else
                IRCAction ndNamesSpecial, , , , strData
                Wait 0.009
            End If
        
        'NOTE: Keep these cases sorted by number please!
        Case "001"
            'welcome note, that's where we can get our initial nickname from
            Pos = InStr(1, strData, " ")
            pos2 = InStr(Pos, strData, ":")
            Pos = InStrRev(strData, " ", pos2 - 2)
            chan = Trim(Mid(strData, Pos, pos2 - Pos))
            If chan <> vbNullString Then
                TheServer.myNick = chan
            End If
            
        Case "005"
            'Server 005 Attributes
            Pos = InStr(1, realData, " ")
            pos2 = InStrRev(realData, ":")
            TheServer.Parse005Attributes Mid$(realData, Pos + 1, pos2 - Pos - 2)
            
        Case "252"
            '# of staff members
            realData = Strings.Mid$(strData, InStr(2, strData, ":") - 2, Len(strData) - InStr(2, strData, ":") + 1)
        
        Case "254"
            '# of registered channels
            Pos = InStr(2, strData, ":")
            pos2 = InStrRev(strData, " ", Pos - 2)
            TheServer.NumberOfChannels = Conversion.CLng(Strings.Mid$(strData, pos2, Pos - pos2))
            realData = Strings.Mid$(strData, InStr(2, strData, ":") - 2, Len(strData) - InStr(2, strData, ":") + 1)
            
        Case "302"
            'reply to /userhost
            If LenB(Trim$(realData)) = 0 Then
                AddStatus Language(816) & vbNewLine, TheServer
                Exit Sub
            End If
            Pos = InStrRev(realData, "@")
            If Pos > 0 Then
                'example 302 message:
                ':bear.freenode.net 302 dionyziz :dionyziz=+dionyziz@62.103.227.219
                pos2 = InStr(1, realData, "=")
                Nick1 = Left$(realData, pos2 - 1)
                Reason = Right$(realData, Len(realData) - Pos)
                strMsgDisplay = Replace(Replace(Language(818), "%1", Nick1), "%2", Reason)
                If Nick1 = LocalLookupNick Then
                    'resolved local IP
                    'store it
                    LocalIP = Reason
                    strMsgDisplay = strMsgDisplay & "(" & Language(817) & ")"
                End If
                AddStatus strMsgDisplay & vbNewLine, TheServer
            End If
            'nick1 =
        
        Case "303" 'ison reply from server
            IRCAction ndIson, , , , realData, TheServerIndex
        'Case "305"
            'marked as back(from away)
        'Case "306"
            'marked as away
            
        'Case "307"
            'this is handled below, together with RAW/320
            
        Case "311"
            'whois: real name
            'If showwhois Then
            Pos = InStr(2, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            pos2 = InStr(Pos + 1, strData, " ")
            Nick1 = Mid$(strData, Pos + 1, pos2 - Pos - 1)
            Pos = InStr(pos2 + 1, strData, " ")
            strMsgDisplay = Mid$(strData, pos2 + 1, Pos - pos2 - 1)
            pos2 = InStr(Pos + 1, strData, " ")
            Reason = Mid$(strData, Pos + 1, pos2 - Pos - 1)
            'Nick1 = nickname
            'strMsgDisplay = real name
            'Reason = connection IP
            'realData = email
            AddStatus Replace(Replace(Replace(Language(836), "%1", Nick1), "%2", strMsgDisplay), "%3", realData) & vbNewLine, TheServer
            AddStatus Replace(Replace(Language(818), "%1", Nick1), "%2", Reason) & vbNewLine, TheServer
            'End If
            'msgString = Strings.Mid$(strData, pos3 + 1, Len(strData) - pos2)
            'msgString = Strings.Left$(msgString, InStr(1, msgString, " ") - 1)
            'IRCAction ndBuddyName, msgString, , "name", realData

        Case "312" 'various /whois messages
            'whois: server
            Pos = InStr(2, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            pos2 = InStr(Pos + 1, strData, " ")
            Nick1 = Mid$(strData, Pos + 1, pos2 - Pos - 1)
            Pos = InStr(pos2 + 1, strData, " ")
            strMsgDisplay = Mid$(strData, pos2 + 1, Pos - pos2 - 1)
            AddStatus Replace(Replace(Replace(Language(837), "%1", Nick1), "%2", strMsgDisplay), "%3", realData) & vbNewLine, TheServer
        
        Case "317"
            'whois: seconds idle, signon time
            Pos = InStr(2, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            pos2 = InStr(Pos + 1, strData, " ")
            Nick1 = Mid$(strData, Pos + 1, pos2 - Pos - 1)
            Pos = InStr(pos2 + 1, strData, " ")
            Pos = InStr(Pos + 1, strData, " ")
            strMsgDisplay = Mid$(strData, pos2 + 1, Pos - pos2 - 1)
            Reason = GetStatement(strMsgDisplay)
            strMsgDisplay = GetParameter(strMsgDisplay)
            AddStatus Replace(Replace(Language(832), "%1", Nick1), "%2", REASON_PREFIX & Reason & REASON_SUFFIX) & vbNewLine, TheServer
            AddStatus Replace(Replace(Language(835), "%1", Nick1), "%2", REASON_PREFIX & LocalTime(strMsgDisplay) & " GMT" & REASON_SUFFIX) & vbNewLine, TheServer
        
        Case "318" 'End of whois
            'If strCurrentPanel = "buddylist" And showwhois = False Then
            '    buddieschecked = buddieschecked + 1
            '    executeCommand "buddy"
            'End If

        Case "319"
            'WhoIs: Is on channels...
            
            Pos = InStr(2, strData, ":")
            pos2 = InStrRev(strData, " ", Pos - 2)
            Nick1 = Mid$(strData, pos2 + 1, Pos - pos2 - 2)
            
            'check if user has joined multiple channels (in order to display the correct LangKey #)
            If InStr(InStr(1, realData, "#") + 1, realData, "#") > 0 Then
                'in multiple channels
                strMsgDisplay = Language(833)
            Else
                strMsgDisplay = Language(834)
            End If
            AddStatus Replace(Replace(strMsgDisplay, "%1", Nick1), "%2", REASON_PREFIX & realData & REASON_SUFFIX) & vbNewLine, TheServer
            'End If
            'If Strings.Right$(realData, 1) = " " Then realData = Strings.Left$(realData, Len(realData) - 1)
            'msgString = Strings.Mid$(strData, pos3 + 1, Len(strData) - pos2)
            'msgString = Strings.Left$(msgString, InStr(1, msgString, " ") - 1)
            'IRCAction ndBuddyName, msgString, , "channels", realData

        Case "320", "307"
            'whois: Identified User
            
            'get the nickname
            Pos = InStr(2, strData, ":")
            pos2 = InStrRev(strData, " ", Pos - 2)
            Nick1 = Mid$(strData, pos2 + 1, Pos - pos2 - 2)
            AddStatus Replace(Language(831), "%1", Nick1) & vbNewLine, TheServer

        Case "321"
            'start of channels list

        Case "322" 'list of channels
            realData = Strings.Mid$(realData, InStr(2, realData, " "))
            IRCAction ndChannelList, , , , realData, TheServerIndex
        Case "323" 'end of the list of channels
            IRCAction ndChannelList, , , , "end of list", TheServerIndex
        Case "324"
            'init channel modes
            Pos = InStr(1, strData, "#")
            pos2 = InStr(Pos, strData, " ")
            chan = Mid$(strData, Pos, pos2 - Pos)
            Reason = Mid$(strData, pos2 + 1, Len(strData) - pos2 - 1 - IIf(Right$(strData, 2) = vbNewLine, 1, 2))
            IRCAction ndInitModes, , chan, , Reason, TheServerIndex
        Case "329"
            'TO DO:
            'Handle these messages
            'example #329 message:
            ':pratchett.freenode.net 329 dionyziz #node-irc 1088282087
            'they are received after doing a /mode #node-irc
            'to check channel's modes
        Case "330"
            'TO DO:
            'Handle these messages
            'example #330 message:
            ':port80c.se.quakenet.org 330 dionyziz Snerf IceChat :is authed as
            'they are received only from some servers after /whois
            '
        Case "332" 'Topic of channel
            Pos = InStr(1, strData, "#")
            chan = Strings.Mid$(strData, Pos, InStr(Pos, strData, " ") - Pos)
            IRCAction ndTopic, , chan, "original", realData, TheServerIndex
        Case "333" 'Who set the topic and when
            IRCAction ndTopicTime, , , , realData, TheServerIndex
        Case "341" 'invitation to sombody to join a channel (sent by us)
            Pos = InStr(1, realData, " ")
            Nick1 = Left$(realData, Pos - 1)
            pos2 = InStr(Pos + 1, realData, " ")
            Nick2 = Mid$(realData, Pos + 1, pos2 - Pos - 1)
            chan = Mid$(realData, pos2 + 1, Len(realData) - pos2)
            IRCAction ndInvite, Nick1, chan, Nick2, realData, TheServerIndex
        Case "353" 'Listing Nicknames
            ReDim names(0)
            Pos = InStr(1, strData, "#")
            chan = Strings.Mid$(strData, Pos, InStr(Pos, strData, " ") - Pos)
            'If namesboolean = False Then
                IRCAction ndNames, , chan, , Strings.Right$(strData, Len(strData) - InStr(2, strData, " :") - 1), TheServerIndex
            'Else
            '    IRCAction ndNamesSpecial, , , , strData
            '    Wait 0.009
            'End If
        Case "366"
            namesboolean = False
            'End of Names list
            '(or end of channels list?)
        'Case "321", "366", "376"
            '321-start of channels list, 366-end of channels list, 376-end of motd
        Case "367" 'list of bans
            IRCAction ndBanList, , GetParameterQuick(realData), , GetParameterQuick(realData, 2, True), TheServerIndex
        Case "368" 'end of ban list
            IRCAction ndBanList, , GetParameterQuick(strData, 3), , "end of list"
        Case "376"
            'end of motd
        Case "379"
            'forwarding to another channel
            IRCAction ndSpecialNotice, , , , Language(815), TheServerIndex
            AddNews Language(815)
        Case "410", "440"
            'services are currently down, try again l8r
            IRCAction ndSpecialNotice, , , , Language(814), TheServerIndex
            AddNews Language(814)
        Case "433" 'nick already in use
            TheServer.myNick = vbNullString
            Nick1 = frmOptions.txtNickname.Text
            AddStatus IMPORTANT_PREFIX & realData & IMPORTANT_SUFFIX & vbNewLine, TheServer
            IRCAction ndUseAltNick, Nick1, , , , TheServerIndex
        Case "451" 'you have not registered
            IRCAction ndSpecialNotice, , , , Language(227), TheServerIndex
        Case "482" 'you are not a channel operator
            Pos = InStr(1, strData, "#")
            pos2 = InStr(2, strData, ":")
            chan = Strings.Mid$(strData, Pos, pos2 - Pos)
            If Right$(chan, 1) = " " Then
                chan = Left$(chan, Len(chan) - 1)
            End If
            IRCAction ndSpecialNotice, msgString, chan, , Language(153), TheServerIndex
        'case "475"
        'case
        Case vbNullString
            'here come the messages that cannot be identified using msgId
            Select Case Strings.Left$(strData, InStr(1, strData, " ") - 1)
                Case "ERROR"
                    'Error
                    IRCAction ndError, vbNullString, , , Strings.Right$(strData, Len(strData) - Len("ERROR")), TheServerIndex
                    Exit Sub
                Case "PING"
                    'PING from Server
                    'send back PONG
                    'TO DO: Check to see what we should send as an argument:
                    ' The data sent as an argument by PING or the local IP address! (DONE)
                    TheServer.SendData "PONG " & Strings.Right$(strData, Len(strData) - Len("PING ")) & vbNewLine
                    AddStatus EVENT_PREFIX & "Ping? " & REASON_PREFIX & "Pong!" & REASON_SUFFIX & EVENT_SUFFIX & vbNewLine, TheServer
                    AddNews "Ping? Pong!"
                    RunScript "Pong"
                    Exit Sub
            End Select
        Case Else
            'check to see if it's special server message
            For i = 0 To UBound(IRCMsg)
                If IRCMsg(i, 0) = msgID Then
                    If msgID = "401" And showwhois = False Then
                        Exit Sub
                    End If
                    'away
                    
                    'TO DO:
                    'don't display "from online to xxx"
                    'but instead "from yyy to xxx"
                    If InStr(1, Language(IRCMsg(i, 1)), "%s") > 0 Then
                        AddStatus EVENT_PREFIX & Replace(Language(IRCMsg(i, 1)), "%s", Language(397 + TheServer.MyStatus)) & EVENT_SUFFIX & vbNewLine, TheServer
                    Else
                        AddStatus EVENT_PREFIX & Language(IRCMsg(i, 1)) & EVENT_SUFFIX & vbNewLine, TheServer
                    End If
                    Exit Sub
                End If
            Next i
            'it's not; just display it
            AddStatus realData & vbNewLine, TheServer
            Exit Sub
    End Select
Already_Loaded:
End Sub
Public Sub buildStatus(Optional ByVal ChannelId As Long = -1)
    'this sub is used to read from an irc-textbox(with index ChannelID)
    'and create an HTML file containing the imported text
    Dim CurFile As Integer 'variable used to store the index of a new free file
    Dim CurFile2 As Integer 'another variable used to store another free file index
    'a Channel HTML is the frameHTML with frames mainHTML and nickHTML
    'a Private or Status HTML is the mainHTML only, without any frames.
    
    Dim intTXTInput As Integer 'where are we going to input the contents from.
    Dim lstNickList As Collection 'where are we going to input the nicknames for the nicklist from.
    Dim i As Integer 'a counter variable for loops
    Dim TabIndex As Integer 'the index of the tab that contains the channel with
                            'channel id the passed parameter ChannelID
    Dim strNick As String 'variable storing the current item while we are looping inside the nicklist
                          'of the channel with channel id ChannelID.
    Dim strNickListInnerHTML As String
    Dim boolTemp As Boolean
    Dim DCCIndex As Integer
    Dim DCCProgress As Byte
    Dim DCCFilename As String
    Dim RCV As Boolean
    Dim TheServer As clsActiveServer
    
    If BSCodeCall Then
        Exit Sub
    End If
    
    'there's no need to have a server parameter;
    'if we're BuildStatus()-ing
    'the information should be imported
    'from CurrentActiveServer
    Set TheServer = CurrentActiveServer
    
    'tabs switching have been reported to be causing some bugs
    'that have been fixed in past versions
    'in case there are some not so obvious bugs
    'put an error handler here and ask the user
    'to report it in case of a bug
    
    On Error GoTo Please_Report_Bug
    
    'if there wasn't a ChannelID passed, or if the number -1 was passed...
    If ChannelId = -1 Then
        '...use the selected tab
        'get the ChannelID of the current tab from the TabInfo collection.
        ChannelId = TheServer.TabInfo(TheServer.Tabs.SelectedItem.index)
        TabIndex = TheServer.Tabs.SelectedItem.index
    Else
        'use GetTab to get the index of the tab which contains the channel with channel id ChannelID.
        '(if the ChannelID passed was -1 then this will be the selected's tab index)
        'Won't work for DCC tabs, as TabInfo contains data about if the file is being send or recieved
        TabIndex = TheServer.GetTab(ChannelId)
    End If
    
    
    'if the tab passed is wrong...
    If TabIndex = 0 Then
        '...don't try to build anything
        Exit Sub
    End If
    
    'if the tab is a website tab...
    If TheServer.TabType(TabIndex) = TabType_WebSite Then
        '...we don't have anything to build!
        Exit Sub
    ElseIf TheServer.TabType(TabIndex) = TabType_DCCFile Then
        If TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) = 1 Then 'RCV: No
            RCV = False
        Else ' = 0; recieving; RCV = Yes
            RCV = True
        End If
        DCCIndex = TheServer.GetDCCFromTab(TheServer.Tabs.SelectedItem.index, RCV)
        DCCProgress = TheServer.DCCFile_Progress(DCCIndex, RCV)
        DCCFilename = TheServer.DCCFile_FileName(DCCIndex, RCV)
        'COMPATIBILITY ISSUE: True equals -1 in VB6, but 1 in VB.net.
        '                     this will have to become - RCV instead of + RCV if we move to VB.net at some point.
        webdocDCCs.All.Item("send_or_recieve").innerText = Language(209 + RCV)
        webdocDCCs.All.Item("textprogress").innerText = DCCProgress & "% " & Language(266)
        webdocDCCs.All.Item("shapeprogress").innerHTML = _
              "<table width=""200px"" height=""10px"" bgColor=""black"" cellspacing=""0"" cellpadding=""0"" class=""pbdcc"">" & _
              "<tr><td class=""progress"" width=""" & DCCProgress * 2 & "px"">" & _
              "</td><td class=""progressbg"" width=""" & (200 - (DCCProgress * 2)) & "px""></td></tr></table>"
        If DCCProgress = 100 Then
            If Not RCV Then
                webdocDCCs.All.Item("textprogress").innerHTML = "<br>" & Language(267) & " <br><input type=""button"" onClick='window.location.href=""NodeScript:/close"";' value='" & Language(515) & "'>"
            Else
                webdocDCCs.All.Item("textprogress").innerHTML = "<br>" & Language(267) & " <br>" & _
                "<input type=""button"" onClick='openfile(""" & Replace(App.Path, "\", "/") & "/downloads/" & DCCFilename & """);' value='" & Language(708) & "'> " & _
                "<input type=""button"" onClick='openfile(""" & Replace(App.Path, "\", "/") & "/downloads" & """);' value='" & Language(709) & "'> " & _
                "<input type=""button"" onClick='window.location.href=""NodeScript:/close"";' value='" & Language(515) & "'> "
                webdocDCCs.All.Item("shapeprogress").innerHTML = _
                  "<table width=""200px"" height=""10px"" bgColor=""black"" cellspacing=""0"" cellpadding=""0"" class=""pbdcc"">" & _
                  "<tr><td class=""progress"" width=""200px"">" & _
                  "</td><td class=""progressbg"" width=""0px""></td></tr></table>"
            End If
        End If
        
        ' display speed and ETL
        webdocDCCs.All.Item("tranfer_rate").innerHTML = Language(260) & ": " & Round(TheServer.DCCTransfer_Speed(DCCIndex, RCV), 3) & " KB/Sec<br>" & _
                                                        Language(259) & ": " & TheServer.DCCTransfer_ETL(DCCIndex, RCV)
        Exit Sub
    End If
    
    'the irc-textbox from which we will input
    intTXTInput = TheServer.TabInfo(TabIndex)
    'if it is a channel...
    If TheServer.TabType(TabIndex) = TabType_Channel Then
        '...get the listbox control containing the nicknames of the current channel
        Set lstNickList = TheServer.NickList_List(ChannelId)
    End If
    
    'if we are talking about a channel we'll have to create
    'a frameset containing the mainHTML and a nickHTML, which
    'we'll have to create as well.
    If TheServer.TabType(TabIndex) = TabType_Channel Then
        'the <table> of the nicklist was opened by the intro file
        'now we'll just have to print the contents of it.
        'we go through the current nicklist items
        strNickListInnerHTML = "<table>"
Fill_NickList:
        For i = 1 To lstNickList.Count
            'get the nickname which should be on the current position
            'and store it in strNick
            strNick = lstNickList.Item(i)
            If LenB(strNick) = 0 Or strNick = vbNewLine Then
                lstNickList.Remove i
                GoTo Fill_NickList
            End If
            'inform the parser that a new column starts
            strNickListInnerHTML = strNickListInnerHTML & "<tr>"
            'and a new field in the table as well.
            strNickListInnerHTML = strNickListInnerHTML & "<td class=""single_nick" & IIf(i = 1, "_1", "") & """>"
            'here we say that this nickname is actually a link
            'when it's clicked the page will navigate to a NodeScript
            'which won't really be a navigation: node will execute the
            'command without moving to that site. The command that
            'should be executed when the nickname-link is clicked is
            '/nickmenu NickName
            '(ExecuteCommand will handle the rest)
            strNickListInnerHTML = strNickListInnerHTML & "<a class=""nick"" href=""NodeScript:/nickmenu " & GetNick(strNick) & """>"
            'we print the nickname itself.
            strNickListInnerHTML = strNickListInnerHTML & "<nobr>" & GetNicklistNick(strNick) & "</nobr>"
            'and close the link tag
            strNickListInnerHTML = strNickListInnerHTML & "</a>"
            'close the field
            strNickListInnerHTML = strNickListInnerHTML & "</td>"
            'and the current column
            strNickListInnerHTML = strNickListInnerHTML & "</tr>"
        'go to the next nickname of the current channel
Continue:
        Next i
        strNickListInnerHTML = strNickListInnerHTML & "</table>"
        
        If webdocChanNicklist.All("nicklist") Is Nothing Then
            Set webdocChanNicklist = webdocChanFrameSet.parentWindow.frames(2).Document
        End If
        webdocChanNicklist.All("nicksText").innerHTML = strNickListInnerHTML
        
        'the cols property must be set to the actual width of the nicklist, which
        'is stored in the tag of the current nicklist :)
        'print the frames text
    End If
    
    'now type to print the
    'main text
    'if we're in full screen mode...
    If MaxMode Then
        '...only update that window
        'check if webdocFullScreen is valid
        If Not webdocFullScreen Is Nothing And Not frmFullScreen.wbCustom(1).Busy Then
            Select Case TheServer.TabType(TabIndex)
                Case TabType_Channel
                    webdocFullScreen.parentWindow.frames(1).Document.All.tags("div").Item(0).innerHTML = TheServer.IRCData(intTXTInput)
                    webdocFullScreen.parentWindow.frames(2).Document.All("nicksText").innerHTML = strNickListInnerHTML
                    webdocFullScreen.parentWindow.frames(1).Document.anchors.Item(webdocFullScreen.parentWindow.frames(0).Document.anchors.length - 1).scrollIntoView
                Case TabType_Private
                    webdocFullScreen.All.tags("div").Item(0).innerHTML = TheServer.IRCData(intTXTInput)
                    webdocFullScreen.anchors.Item(webdocFullScreen.anchors.length - 1).scrollIntoView
            End Select
        End If
        Exit Sub
    End If
    
    'in the file: read it from the irc-textbox and append it to the file
    Select Case TheServer.TabType(TabIndex)
        Case TabType_Channel
            If webdocChanMain.All.tags("div").Item(0) Is Nothing Then
                Set webdocChanMain = webdocChanFrameSet.parentWindow.frames(1).Document
            End If
            webdocChanMain.All.tags("div").Item(0).innerHTML = TheServer.IRCData(intTXTInput)
            webdocChanMain.anchors.Item(webdocChanMain.anchors.length - 1).scrollIntoView
            webdocChanFrameSet.parentWindow.frames(0).Document.All.tags("div").Item(0).innerHTML = TheServer.NickList_Topic_Parsed(ChannelId)
        Case TabType_Private, TabType_Status
            webdocPrivates.All.tags("div").Item(0).innerHTML = TheServer.IRCData(intTXTInput)
            webdocPrivates.anchors.Item(webdocPrivates.anchors.length - 1).scrollIntoView
    End Select
       
    'something has changed. Call the `something changed' script.
    RunScript "Knock"
    
    Exit Sub
Please_Report_Bug:
    CriticalError
End Sub
Public Sub BuildPrimary()
    Dim intFL As Integer
    Dim intFL2 As Integer
    Dim intFl3 As Integer
    
    boolBuildingPrimary = True
    
    '>Chans
    intFL = FreeFile
    Open App.Path & "\temp\main_chan.html" For Output As intFL
    
    intFL2 = FreeFile
    Open App.Path & "\data\html\imports\main_chan.html" For Input As intFL2
    PartCopyR intFL, intFL2
    Close #intFL2
    'the main html file won't need anything else. save it.
    Close #intFL
    
    'copy nicklist
    intFL = FreeFile
    Open App.Path & "\temp\nicklist.html" For Output As intFL
    
    intFL2 = FreeFile
    Open App.Path & "\data\html\imports\nicklist.html" For Input As intFL2
    PartCopyR intFL, intFL2
    Close #intFL2
    'the nicklist html file won't need anything else. save it.
    Close #intFL
    
    '>Privs
    'Note: It is important to BuildPrimary() the private
    'HTML DOM Document first and THEN the FrameSet HTML DOM Document
    'This is because we have to use DoEvents(), which will
    'fire BuildStatus, a procedure that needs access
    'to the current DOM Document. The current DOM Document
    'in this stage is either a Web Site document or
    'the Status Window (a private document)
    intFL = FreeFile
    Open App.Path & "\temp\priv.html" For Output As intFL
    
    intFL2 = FreeFile
    Open App.Path & "\data\html\imports\main.html" For Input As intFL2
    'so we copy all lines from CurFile2 to CurFile...
    'go through CurFile2
    PartCopyR intFL, intFL2
    'we copied all beginning lines... we don't need this file to be open any more.
    'close it.
    Close #intFL2
    Close #intFL
    wbStatus(WebBrowserIndex_Priv).Navigate2 App.Path & "\temp\priv.html"
    DoEvents
    Set webdocPrivates = wbStatus(WebBrowserIndex_Priv).Document
    
    'load the frameset containing both the nicklist and the main html
    wbStatus(WebBrowserIndex_Chan).Navigate2 App.Path & "\data\html\imports\frameset.html"
    DoEvents
    Set webdocChanFrameSet = wbStatus(WebBrowserIndex_Chan).Document
    On Error Resume Next 'we'll reload them later
    Set webdocChanMain = webdocChanFrameSet.parentWindow.frames(1).Document
    Set webdocChanNicklist = webdocChanFrameSet.parentWindow.frames(2).Document
    Set webdocChanTopic = webdocChanFrameSet.parentWindow.frames(0).Document
    
    'remote error handler
    On Error GoTo 0
          
    'DCC
    intFL = FreeFile
    Open App.Path & "\temp\dcc.html" For Output As intFL
    intFL2 = FreeFile
    Open App.Path & "\data\html\imports\dcc.html" For Input As intFL2
    PartCopyR intFL, intFL2
    Close #intFL2
    Close #intFL
    wbStatus(WebBrowserIndex_DCC).Navigate2 App.Path & "\temp\dcc.html"
    DoEvents
    Set webdocDCCs = wbStatus(WebBrowserIndex_DCC).Document

    boolBuildingPrimary = False
End Sub
Public Sub BuildSecondary()
    Dim intFL As Integer
    Dim intFL2 As Integer
    
    'Error Page
    intFL = FreeFile
    Open App.Path & "/temp/error.html" For Output Access Write Lock Write As #intFL
    intFL2 = FreeFile
    Open App.Path & "/data/html/imports/error.html" For Input Access Read Lock Write As #intFL2
    PartCopyR intFL, intFL2
    Close #intFL2
    Close #intFL
    
    'Loading Page
    intFL = FreeFile
    Open App.Path & "/temp/loading.html" For Output Access Write Lock Write As #intFL
    intFL2 = FreeFile
    Open App.Path & "/data/html/imports/loading.html" For Input Access Read Lock Write As #intFL2
    PartCopyR intFL, intFL2
    Close #intFL2
    Close #intFL
    
    'Skin's CSS
    FileCopy App.Path & "/data/skins/" & ThisSkin.TemplateFile, App.Path & "/temp/currentskin/style.css"
End Sub
Public Sub AddLog(ByVal strText As String, Optional ByVal ChannelId As Long = -1, Optional ByVal TheServerIndex As Integer, Optional bolUseTimeStamp As Boolean = False)
    Dim intFL As Integer 'free file index for the log file
    Dim strNet As String
    Dim strFile As String
    Dim intTabType As Integer
    Dim TheServer As clsActiveServer
    
    If TheServerIndex = -1 Then
        Set TheServer = CurrentActiveServer
    Else
        Set TheServer = ActiveServers(TheServerIndex)
    End If
    
    If ChannelId = -1 Then
        ChannelId = TheServer.GetChanID(TheServer.Tabs.SelectedItem.Caption)
    End If
    If TheServer.GetTab(ChannelId) = 0 Then
        'no such channel
        Exit Sub
    End If
    'if it's a channel and logging channels option is enabled...
    '...we'll have to log this channel
    'a tab can't be both a channel and a private, so we'll use OR.
    'if, else, it's a private and privates logging option is enabled...
    '...we'll have to log this private
    intTabType = TheServer.TabType(TheServer.GetTab(ChannelId))
    If (intTabType = TabType_Channel And Options.LogChannels) Or _
       (intTabType = TabType_Private And Options.LogPrivates) Then
        'so this is going to be logged
        
        
        'get a free file index
        intFL = FreeFile
        'open the appropriate log file
        Open GetLogFile(TheServer.Tabs.Tabs.Item(TheServer.GetTab(ChannelId)).Caption, TheServer) _
                            For Append Access Write Lock Write As #intFL
        'add the new contents
        Print #intFL, IIf(bolUseTimeStamp, "[ " & Time & "] ", vbNullString) & Replace(Replace(strText, "<", "&lt;"), ">", "&gt;") & "<BR>"
        'and close file
        Close #intFL
    End If
End Sub
Public Sub AddNarration(ByVal strText As String, ByRef TheServer As clsActiveServer, Optional ByVal ChannelId As Long = -1)
    Dim boolUseNarration As Boolean

    If ChannelId = -1 Then
        '...display the text on the status tab
        ChannelId = TheServer.GetStatusID
    End If
    
    Select Case TheServer.TabType(TheServer.GetTab(ChannelId))
        Case TabType_Channel
            boolUseNarration = Options.NarrationChannels
        Case TabType_Status
            boolUseNarration = Options.NarrationStatus
        Case TabType_Private
            boolUseNarration = Options.NarrationPrivates
    End Select
    boolUseNarration = boolUseNarration And Options.Narration

    If boolUseNarration Then
        MSSpeech.Speak CreateNarrationText(strText)
    End If
End Sub
Public Sub AddStatus(ByVal strText As String, ByRef TheServer As clsActiveServer, Optional ByVal ChannelId As Long = -1, Optional ByVal boolLogIt As Boolean = True, Optional ByVal boolNarrateIt As Boolean = True)
    'This sub adds the specified text to a channel, private or status tab.
    'if logging options are enabled it also appends the new text to the logs
    Dim strResultText As String
    Dim boolUseTimeStamp As Boolean
    Dim boolUseNarration As Boolean
    Dim TabIndex As Integer
    Dim i As Integer
    Dim CurrentTabType As Integer
    Dim NodeKey As String
    
    If ChannelId = -1 Then
        '...display the text on the status tab
        ChannelId = TheServer.GetStatusID
    End If
    
    Select Case TheServer.TabType(TheServer.GetTab(ChannelId))
        Case TabType_Channel
            boolUseTimeStamp = Options.TimeStampChannels
            boolUseNarration = Options.NarrationChannels
        Case TabType_Status
            boolUseTimeStamp = Options.TimeStampStatus
            boolUseNarration = Options.NarrationStatus
        Case TabType_Private
            boolUseTimeStamp = Options.TimeStampPrivates
            boolUseNarration = Options.NarrationPrivates
    End Select
    boolUseTimeStamp = boolUseTimeStamp And Options.TimeStamp
    boolUseNarration = boolUseNarration And Options.Narration
    
    If boolUseNarration And boolNarrateIt Then
        MSSpeech.Speak CreateNarrationText(strText)
    End If
    
    If boolUseTimeStamp Then
        strResultText = SPECIAL_PREFIX & "[ " & Time & " ]" & SPECIAL_SUFFIX & strText
    Else
        strResultText = strText
    End If
    
    'if no 2nd argument was passed or if the value -1 was passed
    If boolLogIt Then
        'TO DO:
        '  Seperate settings for Logging and
        '  Displaying TimeStamps.
        '
        ' i.e.
        '  Display TimeStamps:
        '  [x] Channels
        '  [x] Privates
        '  [x] Status
        '
        '  Log TimeStamps:
        '  [x] Channels
        '  [x] Privates
        '  [x] Status
        
        AddLog strText, ChannelId, GetServerIndexFromActiveServer(TheServer), boolUseTimeStamp And Options.TimeStampLogs
    End If
       
    'irc-textboxes
    'these textboxes hold all the current messages
    'on irc tabs: channels, privates, and status.
    'BuildStatus uses these textboxes to generate
    'the HTML result.
    'Here, we use CreateMainText function to convert
    'the `mIRC special characters' of the passed text to
    'the HTML ones, for example the `mIRC character' for
    'bold(Ctrl + B) is converted to <strong> and </strong>
    'This function also converts smiley characters to
    'the actual smileys.
    'Then, we add the returned text to the current one.
    'If the ChannelID indicates the status window then
    'we won't replace smileys(UseSmileys parameter is False)
    'Else, we will.
    strResultText = CreateMainText(strResultText, TheServer, Not ChannelId = TheServer.GetStatusID And Options.UseSmileys, TheServer.TabType(TheServer.GetTab(ChannelId)) = TabType_Channel)
    TheServer.IRCData(ChannelId) = TheServer.IRCData(ChannelId) & strResultText
    
    'If the tab where the change was made is selected...
    If ChannelId = TheServer.TabInfo(TheServer.Tabs.SelectedItem.index) And TheServer Is CurrentActiveServer Then
        'Reset Make-It-Quicker Timer
        'This timer will refresh the tab in a while.
        'We use this timer so mass-changes need only
        'one refresh and not several
        '(which would crash the program or make it extremely slow)
        'We actually first set the Enabled property to Flase
        'and then to True so the timer actually RESETS.
        'This means that the interval starts to count
        'from the beginning again.
        tmrMakeItQuicker.Enabled = False
        tmrMakeItQuicker.Enabled = True
    Else
        'if the tab where the change was made isn't selected
        '(we may not be able to highlight it for any reason: error trap)
        On Error Resume Next
        'highlight that tab
        'TO DO:
        '   Use different highlight color on the tab
        '   if only new text, and different
        '   if highlighted word (i.e. user's nickname)
        
        TabIndex = TheServer.GetTab(ChannelId)
        TheServer.Tabs.Tabs.Item(TabIndex).Tag = "HighLighted"
        TheServer.Tabs.Tabs.Item(TabIndex).Image = TabImage_Look
        
        CurrentTabType = TheServer.TabType(TheServer.GetTab(ChannelId))
        
        NodeKey = Switch(CurrentTabType = TabType_Channel, "c_", CurrentTabType = TabType_Private, "p_", CurrentTabType = TabType_Status, "s")
        NodeKey = NodeKey & GetServerIndexFromActiveServer(TheServer)
        If Left$(NodeKey, 1) <> "s" Then
            NodeKey = NodeKey & "_" & TheServer.Tabs.Tabs.Item(TabIndex).Caption
        End If
        
        For i = 1 To tvConnections.Nodes.Count
            If tvConnections.Nodes.Item(i).Key = NodeKey Then
                tvConnections.Nodes.Item(i).Image = TabImage_Look
                tvConnections.Nodes.Item(i).Parent.Expanded = True
                Exit For
            End If
        Next i
    End If
End Sub
Public Sub SaveSession()
    'TO DO:
    '   Convert to multiple server framework
    '   Save ALL servers the user is connected to.
    '   (DONE)
    '
    'TO DO:
    '   Check if it works.
    
    Dim intFL As Integer
    Dim objTab As Object
    Dim ActiveServer As Variant
    
    
    intFL = FreeFile
    Open App.Path & "\temp\session.xml" For Output As #intFL
    Print #intFL, "<?xml version=""1.0""?>"
    Print #intFL, "<session when=""" & Now & """>"
    For Each ActiveServer In ActiveServers
        If Not ActiveServer Is Nothing Then
            Print #intFL, "<server hostname=""" & ActiveServer.WinSockConnection.RemoteHost & """ port=""" & ActiveServer.WinSockConnection.RemotePort & """>"
            For Each objTab In ActiveServer.Tabs.Tabs
                If Not (ActiveServer.TabType(objTab.index) = TabType_DCCFile Or ActiveServer.TabType(objTab.index) = TabType_Status) Then
                    Print #intFL, "<tab type=" & ActiveServer.TabType(objTab.index) & """ "
                    Select Case ActiveServer.TabType(objTab.index)
                        Case TabType_Channel, TabType_Private
                            Print #intFL, "info=""" & objTab.Caption & """ "
                        Case TabType_DCCFile
                        Case TabType_Status
                        Case TabType_WebSite
                            Print #intFL, "info=""" & wbStatus(ActiveServer.TabInfo(objTab.index)).LocationURL & """ "
                    End Select
                    Print #intFL, " />"
                End If
            Next objTab
            Print #intFL, "</server>"
        End If
    Next ActiveServer
    Print #intFL, "</session>"
    Close #intFL
End Sub
Public Sub LoadSession(ByVal strFileName As String)
    Dim XMLFile As DOMDocument
    Dim XMLSession As IXMLDOMElement
    Dim XMLTab As IXMLDOMElement
    Dim XMLServer As IXMLDOMElement
    Dim strServer As String
    Dim strPort As String * 5
    Dim strInfo As String
    Dim i As Integer, i2 As Integer
    Dim ChannelId As Integer
    Dim intFL As Integer
    Dim ActiveServer As Variant
       
    Set XMLFile = New DOMDocument
    XMLFile.Load strFileName
    Set XMLSession = XMLFile.documentElement
       
'    If strServer = vbnullstring Or strPort = vbnullstring Then
'        'unable to resume: not connected to server
'        AddStatus SPECIAL_PREFIX & Language(282) & SPECIAL_SUFFIX & vbnewline
'        Exit Sub
'    End If
    
    'unload current session
    For Each ActiveServer In ActiveServers
        ActiveServer.WinSockConnection.Close
        Set ActiveServer = Nothing
    Next ActiveServer
       
    For Each XMLServer In XMLSession.childNodes
        With NewServer
            i = i + 1
            .GetStatusID
            
            intFL = FreeFile
            Open App.Path & "\temp\session_perform_" & i & ".dat" For Output As #intFL
            For Each XMLTab In XMLServer.childNodes
                If XMLTab.nodeName = "tab" Then
                    strInfo = XMLTab.getAttribute("info")
                    Select Case XMLTab.getAttribute("type")
                        Case TabType_Channel
                            Print #intFL, "/join " & strInfo
                        Case TabType_Private
                            Print #intFL, "/query " & strInfo
                        Case TabType_WebSite
                            Print #intFL, "/browse """ & strInfo & """"
                    End Select
                End If
            Next XMLTab
            Close #intFL
            .DoSessionPerform = True
            .SessionPerformFile = App.Path & "\temp\session_perform_" & i & ".dat"
            
            'connect to the server
            .preExecute "/connect " & XMLServer.getAttribute("hostname") & " " & XMLTab.getAttribute("port")
        End With
    Next XMLServer
    
    'Display
    '"resuming previous session"
    'at the last server
    '(which should be the currently active one)
    AddStatus SPECIAL_PREFIX & Language(262) & SPECIAL_SUFFIX & vbNewLine, CurrentActiveServer
End Sub
Public Sub ShowPane()
    IsPaneOpen = Not IsPaneOpen
    fraPane.Visible = IsPaneOpen
    fraPane.Left = 0
    fraPane.Top = nmnuMain.Top + nmnuMain.Height + 20
    imgPaneBegin.Left = 0
    imgPaneBegin.Top = 0
    imgPane.Left = imgPaneBegin.Left + imgPaneBegin.Width
    imgPane.Top = 0
    imgPaneEnd.Top = 0
    lblPaneTitle.Top = 0
    lblPaneTitle.Caption = Language(793)
    picPaneResize.Top = imgPane.Height
    tvConnections.Left = 0
    tvConnections.Top = lblPaneTitle.Top + lblPaneTitle.Height + 40
    Form_Resize
End Sub
Public Sub LoadPanel(ByVal Panel As String)
    Dim i As Integer
    Dim bolPanelLoaded As Boolean
    Dim intPanelCaption As Integer
    Static bolCallByCode As Boolean
    
    If Not fraPanel.Visible Then
        For i = 1 To ntsPanel.NumberOfTabs
            ntsPanel.DeleteTab
        Next i
    End If
    
    If bolCallByCode Then
        bolCallByCode = False
        Exit Sub
    End If
    For i = 0 To UBound(LoadedPanels)
        If LoadedPanels(i) = Panel Then
            bolPanelLoaded = True
            Exit For
        End If
    Next i
    If Not bolPanelLoaded Then
        ReDim Preserve LoadedPanels(UBound(LoadedPanels) + 1)
        LoadedPanels(UBound(LoadedPanels)) = Panel
        
        Select Case LCase$(Panel)
            Case "connect"
                intPanelCaption = 2
            Case "join"
                intPanelCaption = 323
            Case "buddylist"
                intPanelCaption = 310
            Case "favorites"
                intPanelCaption = 488
            Case "avatar"
                intPanelCaption = 456
        End Select
        
        If fraPanel.Visible Then
            ntsPanel.SelectedTab = ntsPanel.AddTab(, Language(intPanelCaption))
        Else
            ntsPanel.TabCaption(1) = Language(intPanelCaption)
        End If
        'ntsPanel.TabIndex = UBound(LoadedPanels)
    Else
        bolCallByCode = True
        ntsPanel.SelectedTab = GetPanelTabIndexFromKey(Panel)
    End If
    fraPanel.Visible = True
    fraPanel.Width = GetSetting("Node", "Remember", "Panel_" & Panel, 2775)
    picPanelResize.Visible = True
    If wbPanel.LocationURL <> "file:///" & Replace(App.Path, "\", "/") & "/data/html/panels/" & Panel & ".html" Then
        wbPanel.Navigate2 App.Path & "\data\html\panels\" & Panel & ".html"
    End If
    wbPanel_DocumentComplete Nothing, wbPanel.LocationURL
    strCurrentPanel = LCase$(Panel)
    Form_Resize
End Sub
Public Sub ShowInfoTip(ByVal WhichOne As NodeInfoTips)
    Dim strTip As String
    Dim strTitle As String
    If Options.InfoTips = False Then
        Exit Sub
    End If
    Select Case WhichOne
        Case WelcomeToNode
            'strTip = Language(284) <-- 0.31
            'strTip = Language(524) <-- 0.32
            'strTip = Language(884) <-- 0.33
            strTip = "Welcome to Node " & App.Major & "." & App.Minor & "!<br>Thanks for downloading. To begin, use the panel on the right!"
            strTitle = Replace(Language(285), "%1", App.Major & "." & App.Minor)
        Case Connected
            strTip = Language(288)
            strTitle = Language(287)
        Case Joined
            strTip = Language(290)
            strTitle = Language(289)
        Case PrivateMessage
            strTip = Language(292)
            strTitle = Language(291)
        Case DCCIncoming
            strTip = Language(294)
            strTitle = Language(293)
        Case Kicked
            strTip = Language(296)
            strTitle = Language(295)
        Case CrashDis
            strTip = Language(298)
            strTitle = Language(297)
        Case NickInUse
            strTip = Language(302)
            strTitle = Language(301)
        Case LangChange
            strTip = Language(304)
            strTitle = Language(303)
        Case SkinChange
            strTip = Language(306)
            strTitle = Language(305)
        Case CrashAsk
            strTip = Language(308)
            strTitle = Language(307)
        Case TrayExit
            strTip = Language(316)
            strTitle = Language(315)
        Case BuddySignOn
            strTip = ScriptNick & " " & Language(517) & vbNewLine & Language(559)
            strTitle = Language(560)
        Case BuddySignOff
            strTip = ScriptNick & " " & Language(516) & vbNewLine & Language(559)
            strTitle = Language(561)
    End Select
    ThisTip = WhichOne
    xpBalloon.ShowBalloonTip Replace(strTip, "<br>", vbNewLine), strTitle, NIIF_INFO, 15000
End Sub
Private Function GetPanelTabIndexFromKey(ByVal Key As String) As Integer
    Dim i As Integer
    For i = 1 To ntsPanel.NumberOfTabs
        If LCase$(LoadedPanels(i)) = LCase$(Key) Then
            GetPanelTabIndexFromKey = i
            Exit Function
        End If
    Next i
    GetPanelTabIndexFromKey = -1
End Function
Private Function ImageFromType(ByVal bTabType As Byte) As Byte
    Select Case bTabType
        Case TabType_Channel
            ImageFromType = TabImage_Channel
        Case TabType_Private
            ImageFromType = TabImage_Private
        Case TabType_Status
            ImageFromType = TabImage_Status
        Case TabType_WebSite
            ImageFromType = TabImage_WebSite
        Case TabType_DCCFile
            ImageFromType = TabImage_DCC
    End Select
End Function
Private Sub ShowBrowser(ByVal BrowserIndex As Integer, ByVal Visibility As Boolean)
    Dim NewLeft As Integer
    
    If Not Visibility Then
        If Not wbStatus(BrowserIndex).Busy Then
            If wbStatus(BrowserIndex).Visible Then
                wbStatus(BrowserIndex).Visible = False
            End If
        End If
    Else
        If Not wbStatus(BrowserIndex).Visible Then
            wbStatus(BrowserIndex).Visible = True
        End If
    End If
    
    NewLeft = IIf(Visibility, ThisSkin.Resize_LeftOffset + IIf(IsPaneOpen, fraPane.Width, 0), -wbStatus(BrowserIndex).Width - 1000)
    If Not wbStatus(BrowserIndex).Left = NewLeft Then
        wbStatus(BrowserIndex).Left = NewLeft
    End If
End Sub
Private Sub ShowMe(Optional doShow As Boolean = True)
    Dim ActiveServer As Variant
    Dim bFade As Byte 'current opacity
    Dim t As Long 'the current system timer
    Dim bValue As Long 'the speed of the transaction; either 5 or -5, depends on if we are fading in or out.
    Dim i As Byte 'a counter variable for the internal transaction loop(for)
    Dim bTransFinal As Byte
    Dim strQuit As String
    Dim intFL As Integer
    Dim intLnCount As Integer
    Dim a As Integer
    
    If Not doShow Then
        ThisSoundSchemePlaySound "end"
    End If
    
    If Options.FadeTransaction Then
        'Alternative: (wont' refresh good)
        'Me.Visible = False
        'AnimateWindow Me.hwnd, 1000, AW_BLEND
        'GoTo End_Transaction
        
        'do the fade transaction; doShow indicates whether we are showing or hiding the window
        
        'the final opacity is 255 if we are fading in, else it's 0; set it.
        bTransFinal = IIf(doShow, 255, 0)
        'the current opacity should be 0 if we are fading in, else it should be 255.
        bFade = IIf(Not doShow, 255, 0)
        'get the current system timer value and store it in t variable.
        t = GetTickCount + 1
        'the speed value depends on doShow
        bValue = IIf(doShow, 5, -5)
        'if it reaches -1 or 256 then the fade is over
        On Error GoTo Byte_Overflow
        'begin the fade loop
        Do
            'if it's time to increase/decrease the opacity do it;
            If t <= GetTickCount Then
                'we'll have to update the system timer variable in order to
                'be able to find when the next increasement or decreasement
                'will occur
                t = GetTickCount + 1
                'change opacity 10 * bValue times
                For i = 1 To 10
                    'update opacity: use the Alpha class
                    'also change the current value variable, bFade.
                    SetLayered Me.hwnd, xLet(bFade, bFade + bValue)
                Next i
                'update view
                DoEvents
            End If
        Loop
Byte_Overflow:
        'the final point was reached
        'set the opacity to its final value
        SetLayered Me.hwnd, bTransFinal
        'if we're not on a web site
        'If currentWB = 0 Then
            'update the webBrowser view
        '    frmMain.wbStatus(0).Refresh2
        'End If
    Else
        Me.Visible = doShow
    End If
    
'End_Transaction:
    'if the program is being hidden it should be ended now
    If Not doShow Then
        
        'close all active connections
        For Each ActiveServer In ActiveServers
            If Not ActiveServer Is Nothing Then
                If ActiveServer.WinSockConnection.State = sckConnected Then
                    If Options.QuitMultiple Then
                        If FS.FileExists(Options.QuitFile) Then
                            intFL = FreeFile
                            Open Options.QuitFile For Input As #intFL
                            Do Until EOF(intFL)
                                xLineInput intFL
                                intLnCount = intLnCount + 1
                            Loop
                            Close #intFL
                            
                            a = Rnd() * intLnCount
                            
                            intLnCount = 0
                            intFL = FreeFile
                            Open Options.QuitFile For Input As #intFL
                            Do Until EOF(intFL)
                                If intLnCount = a Then
                                    Line Input #intFL, strQuit
                                    Exit Do
                                Else
                                    xLineInput intFL
                                End If
                                intLnCount = intLnCount + 1
                            Loop
                            Close #intFL
                        Else
                            strQuit = Options.QuitMsg
                            DB.XWarning "Quit List File does not exist!"
                        End If
                    Else
                        strQuit = Options.QuitMsg
                    End If
                    On Error Resume Next
                    ActiveServer.SendData "QUIT :" & strQuit & vbNewLine
                End If
            End If
        Next ActiveServer
        
        'wait until we get disconnected
        'or until two seconds have passed
        'so that the program can have time
        'to send quit message
        t = GetTickCount
        Do
            DoEvents
            If t + 2000 < GetTickCount Or Not ConnectionIsPresent Then
                Exit Do
            End If
        Loop
        
        'terminate program
        mdlNode.Node_Unload
    End If
End Sub
Private Sub TextToolbarLoadPics()
    Set tbText.ImageList = ilTbText
    tbText.Buttons(1).Image = 1
    tbText.Buttons(3).Image = 2
    tbText.Buttons(4).Image = 3
    tbText.Buttons(5).Image = 4
    tbText.Buttons(7).Image = 5
    tbText.Buttons(9).Image = 6
    tbText.Buttons(10).Image = 7
    tbText.Buttons(11).Image = 8
End Sub
Private Sub MainLoadLanguage()
    mnuMode.Caption = Language(331)
    mnuGiveOp.Caption = Language(332)
    mnuTakeOp.Caption = Language(333)
    mnuGiveVoice.Caption = Language(334)
    mnuTakeVoice.Caption = Language(335)
    mnuKick.Caption = Language(336) & "..."
    mnuInformation.Caption = Language(337)
    mnuInfo.Caption = Language(338)
    mnuWhoIs.Caption = Language(339)
    mnuCTCP.Caption = Language(341)
    mnuCTCPVer.Caption = Language(342)
    mnuCTCPTime.Caption = Language(343)
    mnuCTCPPing.Caption = "Ping" '* TO DO: NEW LANGUAGE ENTRY!!!
    mnuDCC.Caption = Language(344)
    mnuDCCSend.Caption = Language(345)
    mnuDCCChat.Caption = Language(346)
    mnuNDC.Caption = Language(347)
    mnuNDCConnect.Caption = Language(86)
    mnuWhisper.Caption = Language(348)
    mnuCloseTab.Caption = Language(349)
    mnuEnd.Caption = Language(5)
    mnuShow.Caption = Language(350)
    'mnuBuddy.Caption = Language(310)
    mnuAddBuddy.Caption = Language(84)
    mnuNDCAudio.Caption = Language(425)
    mnuWebTabBack.Caption = Language(448)
    mnuWebTabForward.Caption = Language(449)
    mnuWebTabRefresh.Caption = Language(446)
    mnuWebTabStop.Caption = Language(447)
    mnuWebTabFav.Caption = Language(885)
    mnuTabConnect.Caption = Language(2)
    mnuTabDisconnect.Caption = Language(3)
    imgClosePanel.ToolTipText = Language(487)
    tbText.Buttons.Item(1).ToolTipText = Language(513)
    tbText.Buttons.Item(3).ToolTipText = Language(510)
    tbText.Buttons.Item(4).ToolTipText = Language(511)
    tbText.Buttons.Item(5).ToolTipText = Language(512)
    tbText.Buttons.Item(7).ToolTipText = Language(563)
    tbText.Buttons.Item(9).ToolTipText = Language(562)
    tbText.Buttons.Item(10).ToolTipText = Language(732)
    tbText.Buttons.Item(11).ToolTipText = Language(735)
    mnuChanTabLeave.Caption = Language(619)
    mnuChanTabRejoin.Caption = Language(620)
    mnuNickClear.Caption = Language(93)
    mnuNickViewLogs.Caption = Language(249)
    mnuClear.Caption = Language(93)
    mnuViewLogs.Caption = Language(249)
    mnuChanProperties.Caption = Language(738)
    mnuNewServer.Caption = Language(882)
    mnuRemoveServer.Caption = Language(883)
    mnuGiveHalfOp.Caption = Language(760)
    mnuTakeHalfOp.Caption = Language(761)
    mnuUserHost.Caption = Language(819)
    mnuChanModes.Caption = Language(877) & "..."
End Sub
'<section title="Node_Scripting">
Public Function Scripting_GetWindowTitle(TheHWnd As Long) As String
    Dim Title As String
    If IsWindow(TheHWnd) Then
        Title = Strings.Space$(GetWindowTextLength(TheHWnd) + 1)
        Call GetWindowText(TheHWnd, Title, Len(Title))
        Title = Strings.Left$(Title, Len(Title) - 1)
    End If
    Scripting_GetWindowTitle = Title
End Function
Public Function Scripting_GetWindowTextLength(ByVal hwnd As Long) As Long
    Scripting_GetWindowTextLength = GetWindowTextLength(hwnd)
End Function
Public Function Scripting_GetWindowText(ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Scripting_GetWindowText = GetWindowText(hwnd, lpString, cch)
End Function
Public Function Scripting_FindWindow(ByVal lpClassName As String) As Long
    Scripting_FindWindow = FindWindow(lpClassName, vbNullString)
End Function
Public Function Scripting_IsWindow(ByVal hwnd As Long) As Long
    Scripting_IsWindow = IsWindow(hwnd)
End Function
Public Function Scripting_SendMessageM(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Scripting_SendMessageM = SendMessageM(hwnd, wMsg, wParam, lParam)
End Function
Public Function Scripting_CreateEllipticRgn(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Scripting_CreateEllipticRgn = CreateEllipticRgn(X1, Y1, X2, Y2)
End Function
Public Function Scripting_CreateRectRgn(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Scripting_CreateRectRgn = CreateRectRgn(X1, Y1, X2, Y2)
End Function
Public Function Scripting_CombineRgn(ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Scripting_CombineRgn = CombineRgn(hDestRgn, hSrcRgn1, hSrcRgn2, nCombineMode)
End Function
Public Function CAS() As Object
    Set CAS = CurrentActiveServer
End Function
Public Function Scripting_SetWindowRgn(ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Scripting_SetWindowRgn = SetWindowRgn(hwnd, hRgn, bRedraw)
End Function
Public Function Scripting_CreatePolygonRgn(lpPointX As Variant, lpPointY As Variant, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
    Dim apiPoint() As POINTAPI
    Dim i As Integer
    
    ReDim apiPoint(nCount - 1)
    For i = 0 To nCount - 1
        apiPoint(i).X = lpPointX(i)
        apiPoint(i).Y = lpPointY(i)
    Next i
    Scripting_CreatePolygonRgn = CreatePolygonRgn(apiPoint(0), nCount, nPolyFillMode)
End Function
Public Function Scripting_DeleteObject(ByVal hObject As Long) As Long
    Scripting_DeleteObject = DeleteObject(hObject)
End Function
Public Function Scripting_LoadPicture(ByVal Path As String) As IPictureDisp
    Set Scripting_LoadPicture = LoadPicture(Path)
End Function
Public Function Scripting_Lang(ByVal KeyID As Integer) As String
    Scripting_Lang = Language(KeyID)
End Function
Public Function sEVENT_PREFIX() As String
    sEVENT_PREFIX = EVENT_PREFIX
End Function
Public Function sEVENT_SUFFIX() As String
    sEVENT_SUFFIX = EVENT_SUFFIX
End Function
Public Function sREASON_PREFIX() As String
    sREASON_PREFIX = REASON_PREFIX
End Function
Public Function sREASON_SUFFIX() As String
    sREASON_SUFFIX = REASON_SUFFIX
End Function
Public Function sSPECIAL_PREFIX() As String
    sSPECIAL_PREFIX = SPECIAL_PREFIX
End Function
Public Function sSPECIAL_SUFFIX() As String
    sSPECIAL_SUFFIX = SPECIAL_SUFFIX
End Function
Public Function sShellExecute(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    ShellExecute hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd
End Function
'</section>
Public Sub UpdateTabsBar()
    Dim LastTab As Object
    LockWindowUpdate CurrentActiveServer.Tabs.hwnd
    Set LastTab = CurrentActiveServer.Tabs.Tabs(CurrentActiveServer.Tabs.Tabs.Count)
    If CurrentActiveServer.Tabs.Width <> wbStatus(currentWB).Width Then
        CurrentActiveServer.Tabs.Width = wbStatus(currentWB).Width
    End If
    If LastTab.Selected Then
        CurrentActiveServer.Tabs.Width = LastTab.Left + LastTab.Width - ThisSkin.Resize_TabFocusOffset - IIf(IsPaneOpen, fraPane.Width, 0)
    Else
        CurrentActiveServer.Tabs.Width = LastTab.Left + LastTab.Width - ThisSkin.Resize_TabNonFocusOffset - IIf(IsPaneOpen, fraPane.Width, 0)
    End If
    If CurrentActiveServer.Tabs.Width > wbStatus(currentWB).Width Then
        CurrentActiveServer.Tabs.Width = wbStatus(currentWB).Width
    End If
    LockWindowUpdate 0&
End Sub
Private Sub xpBalloon_SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
    On Error Resume Next
    Me.Show
    Me.WindowState = vbMaximized
    Form_Resize
End Sub
Private Sub xpBalloon_SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
    If eButton = vbRightButton Then
        PopupMenu mnuTray, , , , mnuShow
    End If
End Sub
Public Sub AddNews(ByVal strText As String)
    sbar.Panels(1).Text = strText
    tmrClearToolbar.Enabled = False
    tmrClearToolbar.Enabled = True
End Sub
Public Sub CreateWSRoot()
    'CreateWSRoot()
    'Create web sites root node
    'in case it doesn't exist
    Dim i As Integer
    Dim nodeRootNode As ComctlLib.Node
    
    'check if it exists
    'go through the current nodes
    For i = 1 To tvConnections.Nodes.Count
        'if this item is the WSRoot
        If tvConnections.Nodes.Item(i).Key = "wbroot" Then
            'we don't have to create it
            Exit Sub
        ElseIf tvConnections.Nodes.Item(i).Key = "root" Then
            'this is the root now
            'store it for later use
            Set nodeRootNode = tvConnections.Nodes.Item(i)
        End If
    Next i
    
    'there weren't any nodes that were WSRoot
    'we have to create it
    
    'create it and expand it
    'tvConnections.Nodes.Add(nodeRootNode, tvwChild, "wbroot", Language(802), TabImage_WebSite).Expanded = True
    tvConnections.Nodes.Add(, tvwFirst, "wbroot", Language(802), TabImage_WebSite).Expanded = True
End Sub
Public Sub RemoveWSRoot()
    'RemoveWSRoot()
    'Remove web sites root node
    'if it exists in case there
    'are no web sites in it
    Dim i As Integer
    Dim nodeWebRoot As ComctlLib.Node
    
    'check if web sites exist
    For i = 1 To tvConnections.Nodes.Count
        'a web site exists
        If Left$(tvConnections.Nodes.Item(i).Key, 2) = "w_" Then
            'don't remove it
            Exit Sub
        'this is the node we have to remove
        ElseIf tvConnections.Nodes.Item(i).Key = "wbroot" Then
            'store it for later us
            Set nodeWebRoot = tvConnections.Nodes.Item(i)
        End If
    Next i
    
    'there are no web sites
    'check if the web sites node exists
    If Not nodeWebRoot Is Nothing Then
        'it exists
        'remove it
        tvConnections.Nodes.Remove nodeWebRoot.index
    End If
End Sub
Public Sub LoadDialog(ByVal FileName As String, Optional ByRef Window As Form)
    mdlScripting.LoadDialog FileName, Window
End Sub

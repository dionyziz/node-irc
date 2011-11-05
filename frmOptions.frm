VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options Dialog"
   ClientHeight    =   9345
   ClientLeft      =   4140
   ClientTop       =   6045
   ClientWidth     =   18060
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   18060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ident Options"
      Height          =   5175
      Index           =   8
      Left            =   2760
      TabIndex        =   273
      Top             =   0
      Width           =   5895
      Begin VB.Frame fraIdent 
         Height          =   1575
         Index           =   0
         Left            =   360
         TabIndex        =   287
         Top             =   1440
         Width           =   5055
         Begin VB.TextBox txtIdent 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFEFEF&
            ForeColor       =   &H00332222&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   290
            Text            =   "MyOS"
            Top             =   1080
            Width           =   4455
         End
         Begin VB.OptionButton optIdent 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   289
            Top             =   360
            Width           =   200
         End
         Begin VB.OptionButton optIdent 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   288
            Top             =   720
            Width           =   200
         End
         Begin VB.Label lblIdent 
            AutoSize        =   -1  'True
            Caption         =   "Operating System ID"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   293
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label lblIdent 
            BackStyle       =   0  'Transparent
            Caption         =   "Use the default (UNIX)"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   292
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label lblIdent 
            BackStyle       =   0  'Transparent
            Caption         =   "Use this Operating System ID:"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   291
            Top             =   720
            Width           =   4575
         End
      End
      Begin VB.Frame fraIdent 
         Height          =   1575
         Index           =   1
         Left            =   360
         TabIndex        =   280
         Top             =   3120
         Width           =   5055
         Begin VB.OptionButton optIdent 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   285
            Top             =   720
            Width           =   200
         End
         Begin VB.OptionButton optIdent 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   283
            Top             =   360
            Width           =   200
         End
         Begin VB.TextBox txtIdent 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFEFEF&
            ForeColor       =   &H00332222&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   282
            Text            =   "NodeIRC"
            Top             =   1080
            Width           =   4455
         End
         Begin VB.Label lblIdent 
            BackStyle       =   0  'Transparent
            Caption         =   "Use this User ID:"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   286
            Top             =   720
            Width           =   4575
         End
         Begin VB.Label lblIdent 
            BackStyle       =   0  'Transparent
            Caption         =   "Use my nickname"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   284
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label lblIdent 
            AutoSize        =   -1  'True
            Caption         =   "User ID"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   281
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Left            =   600
         TabIndex        =   279
         Text            =   "UNIX"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtIdent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   275
         Text            =   "113"
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkIdent 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   274
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Operating System ID:"
         Height          =   255
         Left            =   480
         TabIndex        =   278
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label lblIdent 
         Caption         =   "On TCP Port"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   277
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblIdent 
         Caption         =   "Enable Ident"
         Height          =   255
         Index           =   0
         Left            =   525
         TabIndex        =   276
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Away Actions"
      Height          =   5175
      Index           =   27
      Left            =   2760
      TabIndex        =   247
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtAwayPerform 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   1005
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   251
         Top             =   1800
         Width           =   5055
      End
      Begin VB.CheckBox chkAwayChangeNick 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   250
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkAwayPerform 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   249
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox txtAwayNick 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Left            =   360
         TabIndex        =   248
         Text            =   "NodeUser_Away"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.Label lblAwayGo 
         Caption         =   "When going away"
         Height          =   255
         Left            =   240
         TabIndex        =   262
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label lblAwayChangeNick 
         Caption         =   "Change my Nickname"
         Height          =   255
         Left            =   525
         TabIndex        =   253
         Top             =   735
         Width           =   4815
      End
      Begin VB.Label lblAwayPerform 
         Caption         =   "Execute these IRC Commands:"
         Height          =   255
         Left            =   525
         TabIndex        =   252
         Top             =   1455
         Width           =   4815
      End
   End
   Begin ComctlLib.TreeView tvOptions 
      Height          =   6135
      Left            =   0
      TabIndex        =   130
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   10821
      _Version        =   327682
      Indentation     =   882
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin SHDocVwCtl.WebBrowser wbNav 
      Height          =   615
      Left            =   1440
      TabIndex        =   132
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
      ExtentX         =   1508
      ExtentY         =   1085
      ViewMode        =   6
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2655
      Left            =   14040
      TabIndex        =   131
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   4683
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picCustom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   54
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O&K"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Apply"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame fraLine 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   2760
      TabIndex        =   41
      Top             =   5100
      Width           =   6015
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Sessions"
      Height          =   4575
      Index           =   9
      Left            =   8640
      TabIndex        =   62
      Top             =   0
      Width           =   5295
      Begin VB.Frame fraSessionC 
         Height          =   1335
         Left            =   120
         TabIndex        =   64
         Top             =   1920
         Width           =   5055
         Begin VB.OptionButton optSessionCA 
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   960
            Width           =   200
         End
         Begin VB.OptionButton optSessionCD 
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   600
            Width           =   200
         End
         Begin VB.OptionButton optSessionCR 
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   200
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Ask"
            Height          =   255
            Index           =   14
            Left            =   360
            TabIndex        =   82
            Top             =   960
            Width           =   4575
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            Height          =   255
            Index           =   13
            Left            =   360
            TabIndex        =   81
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Resume"
            Height          =   255
            Index           =   12
            Left            =   360
            TabIndex        =   80
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label lblSessionCrash 
            AutoSize        =   -1  'True
            Caption         =   "Crash"
            Height          =   195
            Left            =   120
            TabIndex        =   71
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraSessionN 
         Height          =   1335
         Left            =   120
         TabIndex        =   63
         Top             =   360
         Width           =   5055
         Begin VB.OptionButton optSessionNA 
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   960
            Width           =   200
         End
         Begin VB.OptionButton optSessionND 
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   600
            Width           =   200
         End
         Begin VB.OptionButton optSessionNR 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   200
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Ask"
            Height          =   255
            Index           =   11
            Left            =   360
            TabIndex        =   79
            Top             =   960
            Width           =   4575
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   78
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label lblOption 
            BackStyle       =   0  'Transparent
            Caption         =   "Resume"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   77
            Top             =   240
            Width           =   4575
         End
         Begin VB.Label lblSessionNormal 
            AutoSize        =   -1  'True
            Caption         =   "Normal"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Proxy"
      Height          =   5055
      Index           =   13
      Left            =   8760
      TabIndex        =   123
      Top             =   240
      Width           =   5295
      Begin VB.TextBox txtProxyPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   128
         Text            =   "0000"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cboProxy 
         Height          =   315
         Left            =   360
         TabIndex        =   125
         Text            =   "Combo1"
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chkProxyEnable 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Proxy"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   124
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblWarning 
         Height          =   1095
         Left            =   120
         TabIndex        =   129
         Top             =   2520
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label lblProxyPort 
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy Port"
         Height          =   255
         Left            =   360
         TabIndex        =   127
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblProxy 
         BackStyle       =   0  'Transparent
         Caption         =   "Proxy Address"
         Height          =   255
         Left            =   360
         TabIndex        =   126
         Top             =   840
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buddy List"
      Height          =   5055
      Index           =   6
      Left            =   8520
      TabIndex        =   32
      Top             =   0
      Width           =   5535
      Begin VB.CommandButton cmdEditBdy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   61
         Top             =   3480
         Width           =   975
      End
      Begin VB.CheckBox chkLeaveWin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pop-up Dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   51
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CheckBox chkLeaveMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display Text"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         Top             =   4320
         Width           =   2055
      End
      Begin VB.CheckBox chkEnterMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display Text"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   4320
         Width           =   2055
      End
      Begin VB.ListBox lstBdyNk 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   2955
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   1695
      End
      Begin VB.ListBox lstBdyWT 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00854A34&
         Height          =   2955
         Left            =   1920
         TabIndex        =   35
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmdAddnick 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add"
         Height          =   375
         Left            =   2760
         TabIndex        =   34
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CommandButton cmdREBdy 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remove"
         Height          =   375
         Left            =   4080
         TabIndex        =   33
         Top             =   3480
         Width           =   1215
      End
      Begin VB.CheckBox chkEnterWin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pop-up Dialog"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Line lnHighlight 
         BorderColor     =   &H80000014&
         X1              =   5280
         X2              =   0
         Y1              =   3945
         Y2              =   3945
      End
      Begin VB.Label lblBuddyLeave 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "When buddy leaves"
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblBuddyEnter 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "When buddy enters"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblNicklist 
         BackStyle       =   0  'Transparent
         Caption         =   "Nick List"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblWelcome 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Text:"
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.Line lnShadow 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   5280
         Y1              =   3930
         Y2              =   3930
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Timestamps"
      Height          =   5415
      Index           =   18
      Left            =   8520
      TabIndex        =   179
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox chkTimeStamp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Timestamps"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   360
         Width           =   4815
      End
      Begin VB.CheckBox chkTimeStampStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Status"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox chkTimeStampChannels 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Channels"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   181
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chkTimeStampPrivates 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Privates"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   180
         Top             =   1440
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Plugins"
      Height          =   5055
      Index           =   12
      Left            =   8640
      TabIndex        =   119
      Top             =   0
      Width           =   5295
      Begin MSFlexGridLib.MSFlexGrid fgPlugins 
         Height          =   3975
         Left            =   240
         TabIndex        =   122
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7011
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
      Begin VB.Label lblInstalledPlugs 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Installed Plugins"
         Height          =   255
         Left            =   240
         TabIndex        =   121
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblDownloadPlugs 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download More Plugins"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   3540
         MouseIcon       =   "frmOptions.frx":5F32
         MousePointer    =   99  'Custom
         TabIndex        =   120
         Top             =   4680
         Width           =   1680
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Accessibility Options"
      Height          =   5055
      Index           =   17
      Left            =   8520
      TabIndex        =   154
      Top             =   0
      Width           =   5535
      Begin VB.CheckBox chkNarrationStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   163
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chkNarrationPrivates 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   161
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkNarrationChannels 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   159
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkNarrationInterface 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   156
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chkNarration 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Narration"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   155
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblNarrationStatus 
         Caption         =   "Status"
         Height          =   255
         Left            =   645
         TabIndex        =   164
         Top             =   1935
         Width           =   4815
      End
      Begin VB.Label lblNarrationPrivates 
         Caption         =   "Privates"
         Height          =   255
         Left            =   645
         TabIndex        =   162
         Top             =   1575
         Width           =   4815
      End
      Begin VB.Label lblNarrationChannels 
         Caption         =   "Channels"
         Height          =   255
         Left            =   645
         TabIndex        =   160
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label lblNarrationInterface 
         Caption         =   "Interface"
         Height          =   255
         Left            =   645
         TabIndex        =   158
         Top             =   855
         Width           =   4815
      End
      Begin VB.Label lblNarration 
         Caption         =   "Narration"
         Height          =   255
         Left            =   520
         TabIndex        =   157
         Top             =   380
         Width           =   4935
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Log"
      Height          =   4815
      Index           =   2
      Left            =   8640
      TabIndex        =   23
      Top             =   0
      Width           =   5295
      Begin VB.CheckBox chkLogRAW 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log RAW msgs"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   153
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CheckBox chkTimeStampsLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log Timestamps"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   151
         Top             =   2160
         Width           =   4815
      End
      Begin VB.CheckBox chkLogByNet 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log by Network"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   1800
         Width           =   4935
      End
      Begin VB.CheckBox chkLogChannels 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log Channels"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   4935
      End
      Begin VB.CheckBox chkLogPrivates 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Log Privates"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblViewLogs 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Logs"
         ForeColor       =   &H00332222&
         Height          =   315
         Left            =   240
         MouseIcon       =   "frmOptions.frx":6084
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   1320
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scripts"
      Height          =   4695
      Index           =   1
      Left            =   8400
      TabIndex        =   21
      Top             =   600
      Width           =   5295
      Begin VB.CheckBox chkCodeBehind 
         Caption         =   "Allow CodeBehind for Skins"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   4080
         Width           =   4935
      End
      Begin VB.CheckBox chkScripting 
         Caption         =   "Enable Node Scripting"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   3720
         Width           =   4935
      End
      Begin VB.CommandButton cmdDeleteScript 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete"
         Height          =   375
         Left            =   2400
         TabIndex        =   52
         Top             =   3240
         Width           =   1095
      End
      Begin VB.FileListBox flScripts 
         Height          =   2625
         Left            =   120
         Pattern         =   "*.vbs;*.xml"
         TabIndex        =   26
         Top             =   600
         Width           =   5055
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Edit"
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblScripts 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Available Scripts:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ignore"
      Height          =   7935
      Index           =   4
      Left            =   8640
      TabIndex        =   25
      Top             =   0
      Width           =   5295
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         CausesValidation=   0   'False
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   135
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remove"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   134
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add..."
         CausesValidation=   0   'False
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   133
         Top             =   6360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstIgnore 
         Height          =   2400
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add..."
         CausesValidation=   0   'False
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   9
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remove"
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   10
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear"
         CausesValidation=   0   'False
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   11
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ListBox lstIgnore 
         Height          =   2400
         Index           =   1
         Left            =   120
         TabIndex        =   136
         Top             =   3840
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblIgnore 
         Caption         =   "Ignore List #2 (IPs/Hostnames):"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   137
         Top             =   3600
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblIgnore 
         Caption         =   "Ignore List:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "HOT Keys"
      Height          =   4935
      Index           =   11
      Left            =   8640
      TabIndex        =   115
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdRemoveHotKey 
         Caption         =   "Remove"
         Height          =   375
         Left            =   2400
         TabIndex        =   118
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddHotKey 
         Caption         =   "Add..."
         Height          =   375
         Left            =   3840
         TabIndex        =   117
         Top             =   3240
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid fgHotKeys 
         Height          =   2775
         Left            =   120
         TabIndex        =   116
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4895
         _Version        =   393216
         FixedCols       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Smiley Packs"
      Height          =   5175
      Index           =   23
      Left            =   2760
      TabIndex        =   225
      Top             =   0
      Width           =   5895
      Begin VB.CheckBox chkEnableSmileys 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   271
         Top             =   360
         Width           =   255
      End
      Begin MSComctlLib.ImageCombo icSmileyPacks 
         Height          =   330
         Left            =   120
         TabIndex        =   226
         Top             =   960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "icSmileyPacks"
      End
      Begin VB.Label lblShowSmileys 
         Caption         =   "Enable Smileys"
         Height          =   255
         Left            =   405
         TabIndex        =   272
         Top             =   375
         Width           =   4815
      End
      Begin VB.Label lblSmileyPacks 
         Caption         =   "Smiley Pack"
         Height          =   255
         Left            =   120
         TabIndex        =   228
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label lblDownloadSmileyPacks 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download More Smiley Packs"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   3465
         MouseIcon       =   "frmOptions.frx":61D6
         MousePointer    =   99  'Custom
         TabIndex        =   227
         Top             =   4680
         Width           =   2115
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Misc"
      Height          =   4935
      Index           =   14
      Left            =   2880
      TabIndex        =   39
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox chkParseMemoServ 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   269
         Top             =   4320
         Width           =   255
      End
      Begin VB.CheckBox chkKeepChannelsOpen 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   220
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkModesOnJoin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   216
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox chkJoinOnKick 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   214
         Top             =   3240
         Width           =   255
      End
      Begin VB.CheckBox chkJoinOnInvite 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   212
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox chkFocusJoined 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Focus Joined Channels"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   152
         Top             =   2520
         Width           =   4815
      End
      Begin VB.CheckBox chkJoinPanel 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Join Panel"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   113
         Top             =   2160
         Width           =   4815
      End
      Begin VB.CheckBox chkTray 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Keep Node running on the tray"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   75
         Top             =   1800
         Width           =   4815
      End
      Begin VB.CheckBox chkInfoTips 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show InfoTips"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   1440
         Width           =   4815
      End
      Begin VB.CheckBox chkAutocomplete 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Autocomplete"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   59
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chkXPCommonControls 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Use XP-Style controls"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox chkFade 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fade Transaction"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   4815
      End
      Begin VB.Label lblParseMemoServ 
         Caption         =   "Parse MemoServ messages (experimental)"
         Height          =   255
         Left            =   525
         TabIndex        =   270
         Top             =   4320
         Width           =   4815
      End
      Begin VB.Label lblKeepChannelsOpen 
         Caption         =   "Keep channels open"
         Height          =   255
         Left            =   525
         TabIndex        =   219
         Top             =   3960
         Width           =   4815
      End
      Begin VB.Label lblModesOnJoin 
         Caption         =   "/modes on JOIN"
         Height          =   255
         Left            =   525
         TabIndex        =   217
         Top             =   3615
         Width           =   4695
      End
      Begin VB.Label lblRejoinOnKick 
         Caption         =   "Rejoin Channels when Kicked"
         Height          =   255
         Left            =   525
         TabIndex        =   215
         Top             =   3255
         Width           =   4695
      End
      Begin VB.Label lblJoinOnInvite 
         Caption         =   "Join Channels after Invited"
         Height          =   255
         Left            =   525
         TabIndex        =   213
         Top             =   2895
         Width           =   4695
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "User"
      Height          =   5175
      Index           =   0
      Left            =   2760
      TabIndex        =   16
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtReal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   4
         Top             =   2760
         Width           =   3495
      End
      Begin VB.TextBox txtNickname 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   0
         Tag             =   "not used"
         Text            =   "x"
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtAlt 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtAltTwo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   2
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   3495
      End
      Begin VB.Label lblWelcomeWizard 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome Wizard"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   4320
         MouseIcon       =   "frmOptions.frx":6328
         MousePointer    =   99  'Custom
         TabIndex        =   218
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label lblReal 
         BackStyle       =   0  'Transparent
         Caption         =   "Real Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label lblNickname 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblAlt 
         BackStyle       =   0  'Transparent
         Caption         =   "Alternative:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblAlt2 
         BackStyle       =   0  'Transparent
         Caption         =   "Alternative 2:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Import/Export"
      Height          =   5175
      Index           =   19
      Left            =   2760
      TabIndex        =   184
      Top             =   120
      Width           =   5295
      Begin MSComDlg.CommonDialog cdExport 
         Left            =   1200
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   375
         Left            =   2760
         TabIndex        =   187
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export"
         Height          =   375
         Left            =   480
         TabIndex        =   186
         Top             =   2280
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog cdImport 
         Left            =   3480
         Top             =   2760
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label lblExportDescription 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Here goes a small description of the Import/Export feature!"
         Height          =   1935
         Left            =   120
         TabIndex        =   185
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Startup"
      Height          =   5175
      Index           =   20
      Left            =   2760
      TabIndex        =   193
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cboConnectServer 
         Height          =   315
         Left            =   240
         TabIndex        =   201
         Text            =   "Choose Server"
         Top             =   2640
         Width           =   4935
      End
      Begin VB.CheckBox chkLatest 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Latest Version on Startup"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   198
         Top             =   360
         Width           =   4815
      End
      Begin VB.CheckBox chkTOD 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show Tips"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   197
         Top             =   720
         Width           =   4815
      End
      Begin VB.CheckBox chkRestoreStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Restore my Status"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   196
         Top             =   1080
         Width           =   4815
      End
      Begin VB.CheckBox chkStartPage 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Startup Web Site"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   195
         Top             =   1440
         Width           =   4815
      End
      Begin VB.TextBox txtStartPage 
         Height          =   285
         Left            =   240
         TabIndex        =   194
         Text            =   "Text1"
         Top             =   1800
         Width           =   4935
      End
      Begin VB.CheckBox chkStartupConnect 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Startup Web Site"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   199
         Top             =   2280
         Width           =   255
      End
      Begin VB.Label lblStartupConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   360
         TabIndex        =   200
         Top             =   2280
         Width           =   5055
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Browsing"
      Height          =   5175
      Index           =   25
      Left            =   2760
      TabIndex        =   234
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton optBrowseInternal 
         Height          =   255
         Left            =   240
         TabIndex        =   239
         Top             =   1080
         Width           =   200
      End
      Begin VB.OptionButton optBrowseDefaultBrowser 
         Height          =   255
         Left            =   240
         TabIndex        =   238
         Top             =   720
         Width           =   200
      End
      Begin VB.CheckBox chkBrowseParseLinks 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   235
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblBrowseInternal 
         BackStyle       =   0  'Transparent
         Caption         =   "Node internal browser"
         Height          =   255
         Left            =   480
         TabIndex        =   241
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label lblBrowseDefaultBrowser 
         BackStyle       =   0  'Transparent
         Caption         =   "My default browser"
         Height          =   255
         Left            =   480
         TabIndex        =   240
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label lblBrowseUsing 
         Caption         =   "Browse Using"
         Height          =   255
         Left            =   240
         TabIndex        =   237
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label lblBrowseParseLinks 
         Caption         =   "Parse Links"
         Height          =   255
         Left            =   525
         TabIndex        =   236
         Top             =   1575
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Perform"
      Height          =   5175
      Index           =   3
      Left            =   2760
      TabIndex        =   24
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cboPerformServer 
         Height          =   315
         Left            =   360
         TabIndex        =   192
         Text            =   "Choose Server"
         Top             =   1440
         Width           =   5295
      End
      Begin VB.OptionButton optPerformSingle 
         Height          =   255
         Left            =   240
         TabIndex        =   189
         Top             =   720
         Width           =   200
      End
      Begin VB.OptionButton optPerformMultiple 
         Height          =   255
         Left            =   240
         TabIndex        =   188
         Top             =   1080
         Width           =   200
      End
      Begin VB.CheckBox chkPerform 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Perform"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   4335
      End
      Begin RichTextLib.RichTextBox RTB 
         Height          =   3255
         Left            =   360
         TabIndex        =   5
         Top             =   1800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5741
         _Version        =   393217
         BackColor       =   15724527
         Enabled         =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmOptions.frx":647A
      End
      Begin VB.Label lblPerformSingle 
         BackStyle       =   0  'Transparent
         Caption         =   "Single Perform"
         Height          =   255
         Left            =   480
         TabIndex        =   191
         Top             =   720
         Width           =   5415
      End
      Begin VB.Label lblPerformMultiple 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiple Performs"
         Height          =   255
         Left            =   480
         TabIndex        =   190
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CTCP"
      Height          =   5055
      Index           =   15
      Left            =   3000
      TabIndex        =   138
      Top             =   0
      Width           =   5895
      Begin VB.CheckBox chkCTCPFloodBounce 
         Caption         =   "Bounce requests back to flooders"
         Height          =   255
         Left            =   360
         TabIndex        =   148
         Top             =   4080
         Width           =   5295
      End
      Begin VB.CheckBox chkCTCPFlood 
         Caption         =   "CTCP Flood Protection"
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   3720
         Width           =   5535
      End
      Begin VB.CheckBox chkCTCPTimeIgnore 
         Caption         =   "Ignore Time requests from Ignored nicknames"
         Height          =   255
         Left            =   360
         TabIndex        =   146
         Top             =   2880
         Width           =   5295
      End
      Begin VB.CheckBox chkCTCPTime 
         Caption         =   "Reply to Time requests"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox txtCTCPVersionCustom 
         Height          =   285
         Left            =   360
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   1800
         Width           =   4815
      End
      Begin VB.CheckBox chkCTCPVersionCustom 
         Caption         =   "Custom Reply"
         Height          =   255
         Left            =   360
         TabIndex        =   143
         Top             =   1440
         Width           =   5295
      End
      Begin VB.CheckBox chkCTCPVersion 
         Caption         =   "Reply to Version requests"
         Height          =   255
         Left            =   120
         TabIndex        =   142
         Top             =   1080
         Width           =   5535
      End
      Begin VB.CheckBox chkCTCPVersionIgnore 
         Caption         =   "Ignore Version requests from Ignored nicknames"
         Height          =   255
         Left            =   360
         TabIndex        =   141
         Top             =   2160
         Width           =   5295
      End
      Begin VB.CheckBox chkCTCPPingIgnore 
         Caption         =   "Ignore Ping requests from Ignored nicknames"
         Height          =   255
         Left            =   360
         TabIndex        =   140
         Top             =   720
         Width           =   5295
      End
      Begin VB.CheckBox chkCTCPPing 
         Caption         =   "Reply to Ping requests"
         Height          =   255
         Left            =   120
         TabIndex        =   139
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Language"
      Height          =   5055
      Index           =   5
      Left            =   2880
      TabIndex        =   27
      Top             =   0
      Width           =   5295
      Begin MSComctlLib.ImageList ilLanguage 
         Left            =   240
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ImageCombo icLanguage 
         Height          =   330
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "icLanguage"
      End
      Begin VB.Label lblDownloadLang 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download More Languages Online"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   2760
         MouseIcon       =   "frmOptions.frx":64FE
         MousePointer    =   99  'Custom
         TabIndex        =   114
         Top             =   4680
         Width           =   2460
      End
      Begin VB.Label lblSelectLang 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Please select your Language"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Skin"
      Height          =   5055
      Index           =   7
      Left            =   2880
      TabIndex        =   31
      Top             =   0
      Width           =   5415
      Begin VB.CheckBox chkLoading 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Loading screen when loading web sites"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   4815
      End
      Begin VB.CheckBox chkHTMLError 
         Caption         =   "Error screens in skins colors."
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   4935
      End
      Begin VB.ComboBox cboSkin 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblDownloadSkins 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download More Skins Online"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   3240
         MouseIcon       =   "frmOptions.frx":6650
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblSkin 
         BackStyle       =   0  'Transparent
         Caption         =   "Select your Skin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sound Scheme"
      Height          =   5175
      Index           =   22
      Left            =   2760
      TabIndex        =   221
      Top             =   0
      Width           =   5895
      Begin MSComctlLib.ImageCombo icSoundSchemes 
         Height          =   330
         Left            =   120
         TabIndex        =   223
         Top             =   600
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "icSoundSchemes"
      End
      Begin VB.Label lblDownloadSoundSchemes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download More Sound Schemes"
         ForeColor       =   &H00332222&
         Height          =   195
         Left            =   3240
         MouseIcon       =   "frmOptions.frx":67A2
         MousePointer    =   99  'Custom
         TabIndex        =   224
         Top             =   4680
         Width           =   2340
      End
      Begin VB.Label lblSoundScheme 
         Caption         =   "Sound Scheme"
         Height          =   255
         Left            =   120
         TabIndex        =   222
         Top             =   360
         Width           =   5535
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00EFEFEF&
      Caption         =   "DCC Options"
      Height          =   5295
      Index           =   10
      Left            =   2760
      TabIndex        =   73
      Top             =   0
      Width           =   5895
      Begin VB.CheckBox chkAutoNDC 
         Caption         =   "Autorequest NDC"
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   4200
         Width           =   4815
      End
      Begin MSComDlg.CommonDialog cdAntiV 
         Left            =   1680
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "Antivirus Executables (*.exe)|*.exe|Norton Antivirus Executable|NAVW32.EXE"
         Flags           =   4
      End
      Begin VB.Frame fraDCCRecieve 
         Height          =   3735
         Left            =   120
         TabIndex        =   83
         Top             =   360
         Width           =   5055
         Begin VB.Frame fraPortR 
            Caption         =   "Port Range"
            Height          =   1335
            Left            =   2520
            TabIndex        =   208
            Top             =   1680
            Width           =   2415
            Begin VB.TextBox TxtRangeH 
               Height          =   285
               Left            =   960
               TabIndex        =   210
               Text            =   "7000"
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox TxtRangeL 
               Height          =   285
               Left            =   240
               TabIndex        =   209
               Text            =   "6000"
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblSeperator 
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   161
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1320
               TabIndex        =   211
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.Frame fraDCCB 
            Height          =   1335
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton optDCCB 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   102
               Top             =   240
               Width           =   200
            End
            Begin VB.OptionButton optDCCB 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   103
               Top             =   600
               Width           =   200
            End
            Begin VB.OptionButton optDCCB 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   104
               Top             =   960
               Width           =   200
            End
            Begin VB.Label lblBuddies 
               AutoSize        =   -1  'True
               Caption         =   "Buddies"
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   0
               Width           =   570
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Accept"
               Height          =   255
               Index           =   0
               Left            =   360
               TabIndex        =   107
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ignore"
               Height          =   255
               Index           =   1
               Left            =   360
               TabIndex        =   106
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ask"
               Height          =   255
               Index           =   2
               Left            =   360
               TabIndex        =   105
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.Frame fraDCCI 
            Height          =   1335
            Left            =   2520
            TabIndex        =   93
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton optDCCI 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   96
               Top             =   960
               Width           =   200
            End
            Begin VB.OptionButton optDCCI 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   200
            End
            Begin VB.OptionButton optDCCI 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   95
               Top             =   600
               Width           =   200
            End
            Begin VB.Label lblIgnored 
               AutoSize        =   -1  'True
               Caption         =   "Ignored Nicks"
               Height          =   195
               Left            =   120
               TabIndex        =   100
               Top             =   0
               Width           =   990
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Accept"
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   99
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ignore"
               Height          =   255
               Index           =   4
               Left            =   360
               TabIndex        =   98
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ask"
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   97
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.Frame fraDCCE 
            Height          =   1335
            Left            =   120
            TabIndex        =   85
            Top             =   1680
            Width           =   2415
            Begin VB.OptionButton optDCCE 
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   88
               Top             =   960
               Width           =   200
            End
            Begin VB.OptionButton optDCCE 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   87
               Top             =   600
               Width           =   200
            End
            Begin VB.OptionButton optDCCE 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   86
               Top             =   240
               Width           =   200
            End
            Begin VB.Label lblEverybody 
               AutoSize        =   -1  'True
               Caption         =   "Everybody"
               Height          =   195
               Left            =   120
               TabIndex        =   92
               Top             =   0
               Width           =   750
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Accept"
               Height          =   255
               Index           =   6
               Left            =   360
               TabIndex        =   91
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ignore"
               Height          =   255
               Index           =   7
               Left            =   360
               TabIndex        =   90
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label lblOption 
               BackStyle       =   0  'Transparent
               Caption         =   "Ask"
               Height          =   255
               Index           =   8
               Left            =   360
               TabIndex        =   89
               Top             =   960
               Width           =   1935
            End
         End
         Begin VB.CheckBox chkAntivirus 
            Caption         =   "Antivirus"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   3120
            Width           =   2415
         End
         Begin VB.Label lblDCCReceive 
            AutoSize        =   -1  'True
            Caption         =   "DCC Recieve"
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblViewDownloads 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "View Downloads"
            ForeColor       =   &H00332222&
            Height          =   315
            Left            =   120
            MouseIcon       =   "frmOptions.frx":68F4
            MousePointer    =   99  'Custom
            TabIndex        =   112
            Top             =   3360
            Width           =   4815
         End
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Retry"
      Height          =   5175
      Index           =   21
      Left            =   2760
      TabIndex        =   202
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtRetryDelay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   360
         TabIndex        =   206
         Text            =   "3"
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkRetry 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   203
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblRetryDelaySeconds 
         Caption         =   "seconds"
         Height          =   255
         Left            =   1560
         TabIndex        =   207
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblRetryDelay 
         Caption         =   "Retry Delay"
         Height          =   255
         Left            =   360
         TabIndex        =   205
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblRetry 
         Caption         =   "Enable Retry"
         Height          =   255
         Left            =   525
         TabIndex        =   204
         Top             =   375
         Width           =   4935
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Display"
      Height          =   5175
      Index           =   26
      Left            =   2760
      TabIndex        =   242
      Top             =   0
      Width           =   5895
      Begin VB.TextBox txtDisplaySkinFontSize 
         Height          =   285
         Left            =   120
         TabIndex        =   245
         Text            =   "Text1"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txtDisplayNormal 
         Height          =   285
         Left            =   120
         TabIndex        =   243
         Text            =   "Text1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblDisplaySkinFontSize 
         Caption         =   "Font Size"
         Height          =   255
         Left            =   120
         TabIndex        =   246
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lblDisplayNormal 
         Caption         =   "How shall Node Display Nicknames"
         Height          =   255
         Left            =   120
         TabIndex        =   244
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Away System"
      Height          =   6375
      Index           =   24
      Left            =   3000
      TabIndex        =   229
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtAwayMinutes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Left            =   360
         TabIndex        =   233
         Text            =   "10"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkUseAway 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   120
         TabIndex        =   230
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblAwayMinutes 
         Caption         =   "Minutes:"
         Height          =   255
         Left            =   240
         TabIndex        =   232
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblUseAway 
         Caption         =   "Use Away System"
         Height          =   255
         Left            =   405
         TabIndex        =   231
         Top             =   375
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Back Actions"
      Height          =   5175
      Index           =   28
      Left            =   2400
      TabIndex        =   254
      Top             =   120
      Width           =   5895
      Begin VB.TextBox txtAwayBackNick 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Left            =   360
         TabIndex        =   258
         Text            =   "NodeUser_Away"
         Top             =   1080
         Width           =   4935
      End
      Begin VB.CheckBox chkAwayBackPerform 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   257
         Top             =   1440
         Width           =   255
      End
      Begin VB.CheckBox chkAwayBackChangeNick 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Interface"
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   256
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtAwayBackPerform 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   1005
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   255
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label lblAwayBackGo 
         Caption         =   "When backing from away"
         Height          =   255
         Left            =   240
         TabIndex        =   261
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label lblAwayBackPerform 
         Caption         =   "Execute these IRC Commands:"
         Height          =   255
         Left            =   525
         TabIndex        =   260
         Top             =   1455
         Width           =   4815
      End
      Begin VB.Label lblAwayBackChangeNick 
         Caption         =   "Change my Nickname"
         Height          =   255
         Left            =   525
         TabIndex        =   259
         Top             =   735
         Width           =   4815
      End
   End
   Begin VB.Frame fraOptions 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Messages"
      Height          =   5295
      Index           =   16
      Left            =   2760
      TabIndex        =   149
      Top             =   0
      Width           =   5535
      Begin MSComDlg.CommonDialog cdQuitList 
         Left            =   4200
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.CommandButton cmdQuitBrowse 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   268
         Top             =   1320
         Width           =   375
      End
      Begin VB.OptionButton optQuitMulti 
         Height          =   255
         Left            =   240
         TabIndex        =   265
         Top             =   1080
         Width           =   200
      End
      Begin VB.TextBox txtQuit 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFEFEF&
         ForeColor       =   &H00332222&
         Height          =   285
         Left            =   360
         TabIndex        =   150
         Text            =   "I am using Node IRC, http://node.sourceforge.net"
         Top             =   720
         Width           =   4815
      End
      Begin VB.OptionButton optQuitSingle 
         Height          =   255
         Left            =   240
         TabIndex        =   263
         Top             =   435
         Width           =   200
      End
      Begin VB.CheckBox chkNickLinkMinePriv 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   177
         Top             =   4575
         Width           =   255
      End
      Begin VB.CheckBox chkNickLinkPriv 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   175
         Top             =   4200
         Width           =   255
      End
      Begin VB.CheckBox chkNickLinkMineChan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   173
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox chkNickLinkChan 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   171
         Top             =   3225
         Width           =   255
      End
      Begin VB.CheckBox chkSelectionCopy 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   240
         TabIndex        =   168
         Top             =   2040
         Width           =   255
      End
      Begin VB.CheckBox chkSelectionClear 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00332222&
         Height          =   255
         Left            =   360
         TabIndex        =   165
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label lblQuitMultiFile 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C:\Program Files\Node\ExampleList.lst"
         Height          =   255
         Left            =   480
         TabIndex        =   267
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label lblQuitMulti 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiple Quit Messages (random)"
         Height          =   255
         Left            =   480
         TabIndex        =   266
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblQuitSingle 
         BackStyle       =   0  'Transparent
         Caption         =   "Single Quit Message"
         Height          =   255
         Left            =   480
         TabIndex        =   264
         Top             =   435
         Width           =   4575
      End
      Begin VB.Label lblNickLinkMinePriv 
         Caption         =   "my nick as link"
         Height          =   255
         Left            =   645
         TabIndex        =   178
         Top             =   4590
         Width           =   4815
      End
      Begin VB.Label lblNickLinkPriv 
         Caption         =   "nix as linx"
         Height          =   255
         Left            =   645
         TabIndex        =   176
         Top             =   4215
         Width           =   4815
      End
      Begin VB.Label lblNickLinkMineChan 
         Caption         =   "Display my nick as link"
         Height          =   255
         Left            =   645
         TabIndex        =   174
         Top             =   3615
         Width           =   4815
      End
      Begin VB.Label lblNickLinkChan 
         Caption         =   "Display nix as linx"
         Height          =   255
         Left            =   645
         TabIndex        =   172
         Top             =   3240
         Width           =   4815
      End
      Begin VB.Label lblInPrivates 
         BackStyle       =   0  'Transparent
         Caption         =   "In Privs"
         Height          =   255
         Left            =   240
         TabIndex        =   170
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Label lblInChannels 
         BackStyle       =   0  'Transparent
         Caption         =   "In Chans"
         Height          =   255
         Left            =   240
         TabIndex        =   169
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label lblSelectionCopy 
         Caption         =   "Copy Selection"
         Height          =   255
         Left            =   525
         TabIndex        =   167
         Top             =   2055
         Width           =   4935
      End
      Begin VB.Label lblSelectionClear 
         Caption         =   "Clear Selection after Copy"
         Height          =   255
         Left            =   645
         TabIndex        =   166
         Top             =   2535
         Width           =   4815
      End
   End
   Begin VB.Menu mnuPlugIn 
      Caption         =   "&PlugIn"
      Visible         =   0   'False
      Begin VB.Menu mnuLoadPlugin 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuUnloadPlugin 
         Caption         =   "&Unload"
      End
      Begin VB.Menu mnuPluginStartup 
         Caption         =   "Load on &Startup"
      End
      Begin VB.Menu mnuPlugInsSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlugInOptions 
         Caption         =   "&Options..."
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright (c)
'
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

'Options Dialog
'
'To add a new option:
'1)Create the interface about the option(textboxes, checkboxes or maybe a new options category).
'2)Add your options in the NodeOptions type in module mdlNode.
'3)Write code to load the option in LoadAll. (make sure you set a default value)
'4)Write code to save the option in SaveAll.
'5)Make the option take effect. Get the value of the option and take the appropriate action.
'6)Update the english language file by adding the keys with the captions needed for your option.
'7)Add the necessary code in OptionsLoadLanguage to load the captions.

'Allow only declared variables
Option Explicit
'This stores the servers' hostnames and tcp ports to connect to, loaded from external file
Private Servers() As Variant
'This stores the skins filename with the same order as in the combo list
Private Skins() As String
'two constants used to get the normal height and width of the options dialog
'as it will need to be resized after it is loaded
Private Const OPTIONS_HEIGHT As Integer = 6210
Private Const OPTIONS_WIDTH As Integer = 8500
Private Sub cboPerformServer_Click()
    Dim intFL As Integer
    Dim i As Integer
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As intFL
        Do Until EOF(intFL)
            If i = cboPerformServer.ListIndex Then
                RTB.Text = XMLPerformReadServer(GetStatement(xLineInput(intFL)))
                Exit Do
            Else
                'skip this line
                xLineInput intFL
            End If
            i = i + 1
        Loop
        Close intFL
    End If
End Sub
Private Sub chkAwayBackPerform_Click()
    chkUseAway_Click
End Sub
Private Sub chkAwayPerform_Click()
    chkUseAway_Click
End Sub
Private Sub chkEnableSmileys_Click()
    icSmileyPacks.Enabled = chkEnableSmileys.value = vbChecked
End Sub
Private Sub chkRetry_Click()
    Dim boolEnabled As Boolean
    
    boolEnabled = chkRetry.value = vbChecked
    lblRetryDelaySeconds.Enabled = boolEnabled
    txtRetryDelay.Enabled = boolEnabled
    lblRetryDelay.Enabled = boolEnabled
End Sub
Private Sub chkStartupConnect_Click()
    cboConnectServer.Enabled = chkStartupConnect.value = vbChecked
End Sub
Private Sub chkUseAway_Click()
    Dim bolEnabled As Boolean
    
    bolEnabled = chkUseAway.value = vbChecked
    
    txtAwayMinutes.Enabled = bolEnabled
    lblAwayMinutes.Enabled = bolEnabled
    
    chkAwayChangeNick.Enabled = bolEnabled
    chkAwayPerform.Enabled = bolEnabled
    lblAwayChangeNick.Enabled = bolEnabled
    lblAwayPerform.Enabled = bolEnabled
    txtAwayPerform.Enabled = chkAwayPerform.value = vbChecked And bolEnabled
    txtAwayNick.Enabled = chkAwayChangeNick.value = vbChecked And bolEnabled

    chkAwayBackChangeNick.Enabled = bolEnabled
    chkAwayBackPerform.Enabled = bolEnabled
    lblAwayBackChangeNick.Enabled = bolEnabled
    lblAwayBackPerform.Enabled = bolEnabled
    txtAwayBackPerform.Enabled = chkAwayBackPerform.value = vbChecked And bolEnabled
    txtAwayBackNick.Enabled = chkAwayBackChangeNick.value = vbChecked And bolEnabled
End Sub
Private Sub cmdQuitBrowse_Click()
    On Error GoTo Cancel_was_selected
    cdQuitList.ShowOpen
    
    Options.QuitFile = xLet(lblQuitMultiFile.Caption, cdQuitList.FileName)
    
    Exit Sub
Cancel_was_selected:
End Sub
Private Sub lblAwayBackChangeNick_Click()
    chkAwayBackChangeNick.value = -(Not CBool(chkAwayBackChangeNick.value))
End Sub
Private Sub lblAwayBackPerform_Click()
    chkAwayBackPerform.value = -(Not CBool(chkAwayBackPerform.value))
End Sub
Private Sub lblAwayChangeNick_Click()
    chkAwayChangeNick.value = -(Not CBool(chkAwayChangeNick.value))
End Sub
Private Sub lblAwayPerform_Click()
    chkAwayPerform.value = -(Not CBool(chkAwayPerform.value))
End Sub
Private Sub lblBrowseDefaultBrowser_Click()
    optBrowseDefaultBrowser.value = True
End Sub
Private Sub lblBrowseInternal_Click()
    optBrowseInternal.value = True
End Sub
Private Sub lblBrowseParseLinks_Click()
    chkBrowseParseLinks.value = -(Not CBool(chkBrowseParseLinks.value))
End Sub
Private Sub lblDownloadSmileyPacks_Click()
    'the user wants to download more sound schemes
    'navigate him/her to the smiley packs page
    CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=smileypacks"
    frmOptions.Hide
End Sub
Private Sub lblIdentCustomOSID_Click()
    optIdent(1).value = True
End Sub
Private Sub lblIdent_Click(Index As Integer)
    Select Case Index
        Case 0
            chkIdent.value = -(Not CBool(chkIdent.value))
        Case 1
            optIdent(0).value = True
        Case 2
            optIdent(1).value = True
        Case 4
            optIdent(2).value = True
        Case 5
            optIdent(3).value = True
    End Select
End Sub
Private Sub lblJoinOnInvite_Click()
    chkJoinOnInvite.value = -(Not CBool(chkJoinOnInvite.value))
End Sub
Private Sub lblKeepChannelsOpen_Click()
    chkKeepChannelsOpen.value = -(Not CBool(chkKeepChannelsOpen.value))
End Sub
Private Sub lblModesOnJoin_Click()
    chkModesOnJoin.value = -(Not CBool(chkModesOnJoin.value))
End Sub
Private Sub lblParseMemoServ_Click()
    chkParseMemoServ.value = -(Not CBool(chkParseMemoServ.value))
End Sub
Private Sub lblQuitMulti_Click()
    optQuitMulti.value = True
End Sub
Private Sub lblQuitSingle_Click()
    optQuitSingle.value = True
End Sub
Private Sub lblRejoinOnKick_Click()
    chkJoinOnKick.value = -(Not CBool(chkJoinOnKick.value))
End Sub
Private Sub lblRetry_Click()
    chkRetry.value = -(Not CBool(chkRetry.value))
End Sub
Private Sub lblShowSmileys_Click()
    chkEnableSmileys.value = -(Not CBool(chkEnableSmileys.value))
End Sub
Private Sub lblStartupConnect_Click()
    chkStartupConnect.value = -(Not CBool(chkStartupConnect.value))
End Sub
Private Sub lblUseAway_Click()
    chkUseAway.value = -(Not CBool(chkUseAway.value))
End Sub
Private Sub lblWelcomeWizard_Click()
    LoadWizard "Welcome"
End Sub
Private Sub optIdent_Click(Index As Integer)
    If optIdent(Index).value Then
        Select Case Index
            Case 0
                txtIdent(0).Enabled = False
            Case 1
                txtIdent(0).Enabled = True
                txtIdent(0).SetFocus
                txtIdent(0).SelStart = 0
                txtIdent(0).SelLength = Len(txtIdent(0).Text)
            Case 2
                txtIdent(1).Enabled = False
            Case 3
                txtIdent(1).Enabled = True
                txtIdent(1).SetFocus
                txtIdent(1).SelStart = 0
                txtIdent(1).SelLength = Len(txtIdent(1).Text)
        End Select
    End If
End Sub
Private Sub optPerformMultiple_Click()
    cboPerformServer.Enabled = True
    cboPerformServer_Click
End Sub
Private Sub optPerformSingle_Click()
    cboPerformServer.Enabled = False
    RTB.Text = XMLPerformReadServer(" default")
End Sub
Private Sub optQuitMulti_Click()
    lblQuitMultiFile.Enabled = optQuitMulti.value
    txtQuit.Enabled = Not optQuitMulti.value
End Sub
Private Sub optQuitSingle_Click()
    optQuitMulti_Click
End Sub
Private Sub RTB_LostFocus()
    Dim strResultHostName As String
    Dim intFL As Integer
    Dim i As Integer
    
    If optPerformSingle.value Then
        strResultHostName = " default"
    Else
        If FS.FileExists(App.Path & "\conf\servers.lst") Then
            intFL = FreeFile
            Open App.Path & "\conf\servers.lst" For Input Access Read Shared As intFL
            Do Until EOF(intFL)
                If i = cboPerformServer.ListIndex Then
                    strResultHostName = GetStatement(xLineInput(intFL))
                    Exit Do
                Else
                    'skip this line
                    xLineInput intFL
                End If
                i = i + 1
            Loop
            Close intFL
        End If
    End If
    
    XMLPerformSaveServer strResultHostName, RTB.Text
End Sub
Private Sub chkAntivirus_Click()
    If chkAntivirus.value = vbChecked Then
        'only if we haven't just loaded options...
        If LenB(Options.DCCAntivirus) = 0 Then
            'pick antivirus executable
            cdAntiV.DialogTitle = Language(434)
            cdAntiV.Filter = _
                Language(435) & " (*.exe)|*.exe|" & _
                Language(436) & "|NAVW32.EXE"
                
            On Error GoTo Cancel_Occured
            cdAntiV.ShowOpen
            Options.DCCAntivirus = cdAntiV.FileName
            If Right$(LCase$(Options.DCCAntivirus), Len("NAVW32.EXE")) = "navw32.exe" Then
                'using norton antivirus
                Options.DCCAntivirus = Options.DCCAntivirus & " /NORESULTS"
            End If
        End If
    Else
        Options.DCCAntivirus = vbNullString
    End If
    Exit Sub
Cancel_Occured:
    'this will call this procedure again, so that Options.DCCAntivirus is vbnullstring
    chkAntivirus.value = vbUnchecked
End Sub
Private Sub chkCTCPFlood_Click()
    Dim boolChksEnabled As Boolean
    
    boolChksEnabled = chkCTCPFlood.value = vbChecked
    chkCTCPFloodBounce.Enabled = boolChksEnabled
End Sub
Private Sub chkCTCPPing_Click()
    Dim boolChksEnabled As Boolean
    
    boolChksEnabled = chkCTCPPing.value = vbChecked
    chkCTCPPingIgnore.Enabled = boolChksEnabled
End Sub
Private Sub chkCTCPTime_Click()
    Dim boolChksEnabled As Boolean
    
    boolChksEnabled = chkCTCPTime.value = vbChecked
    chkCTCPTimeIgnore.Enabled = boolChksEnabled
End Sub
Private Sub chkCTCPVersion_Click()
    Dim boolChksEnabled As Boolean
    
    boolChksEnabled = chkCTCPVersion.value = vbChecked
    chkCTCPVersionIgnore.Enabled = boolChksEnabled
    chkCTCPVersionCustom.Enabled = boolChksEnabled
    
    txtCTCPVersionCustom.Enabled = boolChksEnabled And chkCTCPVersionCustom.value = vbChecked
End Sub
Private Sub chkCTCPVersionCustom_Click()
    chkCTCPVersion_Click
End Sub
Private Sub chkNarration_Click()
    Dim boolNarrationEnabled As Boolean
    
    boolNarrationEnabled = chkNarration.value = vbChecked
    chkNarrationInterface.Enabled = boolNarrationEnabled
    lblNarrationInterface.Enabled = boolNarrationEnabled
    chkNarrationChannels.Enabled = boolNarrationEnabled
    lblNarrationChannels.Enabled = boolNarrationEnabled
    chkNarrationPrivates.Enabled = boolNarrationEnabled
    lblNarrationPrivates.Enabled = boolNarrationEnabled
    chkNarrationStatus.Enabled = boolNarrationEnabled
    lblNarrationStatus.Enabled = boolNarrationEnabled
End Sub
Private Sub chkProxyEnable_Click()
    Dim boolShowProxy As Boolean
    boolShowProxy = chkProxyEnable.value = vbChecked
    lblProxy.Visible = boolShowProxy
    lblProxyPort.Visible = boolShowProxy
    txtProxyPort.Visible = boolShowProxy
    cboProxy.Visible = boolShowProxy
End Sub
Private Sub chkSelectionCopy_Click()
    Dim boolChkEnabled As Boolean
    
    boolChkEnabled = chkSelectionCopy.value = vbChecked
    chkSelectionClear.Enabled = boolChkEnabled
    lblSelectionClear.Enabled = boolChkEnabled
End Sub
Private Sub chkStartPage_Click()
    Dim boolChkEnabled As Boolean
    
    boolChkEnabled = chkStartPage.value = vbChecked
    txtStartPage.Enabled = boolChkEnabled
End Sub
Private Sub chkTimeStamp_Click()
    Dim boolTimeStampEnabled As Boolean
    
    boolTimeStampEnabled = chkTimeStamp.value = vbChecked
    chkTimeStampChannels.Enabled = boolTimeStampEnabled
    chkTimeStampPrivates.Enabled = boolTimeStampEnabled
    chkTimeStampStatus.Enabled = boolTimeStampEnabled
End Sub
Private Sub cmdAdd_Click(Index As Integer)
    'someone is being added in the ignore list
    Dim Nick As String 'who's added; can be a nick or ip
    
    'input nickname
    Nick = InputBox(IIf(Index = 0, Language(100), Language(625)), Language(101))
    
    'if cancel was selected or no nickname entered then skip the addition to the list.
    If LenB(Nick) = 0 Then
        Exit Sub
    End If
    
    'else, add the item
    lstIgnore(Index).AddItem Nick
End Sub
Private Sub cmdAddHotKey_Click()
    LoadDialog App.Path & "/data/dialogs/hotkey.xml"
End Sub
Public Sub cmdAddnick_Click()
    'Add nickname into buddy list
    Dim intFL As Integer 'a temporary variable for the file index
    Dim buddyname As String 'the nickname of the buddy the user wants to add
    Dim buddytext As String 'the welcome text for the buddy the user wants to add
    Dim i As Integer 'a counter variable for the loops
    
    'input the buddy name
    buddyname = InputBox(Language(142), Language(84))
    'if the user entered a nickname and didn't click cancel...
    If LenB(buddyname) > 0 Then
        '... go through the buddy list to see if the nickname already exists
        For i = 0 To lstBdyNk.ListCount - 1
            'if the buddy nickname already exists...
            If buddyname = lstBdyNk.List(i) Then
                '...display warning
                MsgBox Language(157), vbInformation
                'and do not add the buddy
                Exit Sub
            End If
        'go to the next nickname
        Next i
        'if we reached this point, it means that the buddy doesn't exist
        'input the welcome text for the buddy
        buddytext = InputBox(Language(143), Language(84))
        'add an item in the buddy list
        lstBdyNk.AddItem buddyname
        'add an item in the welcome text list
        lstBdyWT.AddItem buddytext
        'get a free file index
        intFL = FreeFile
        'open the configuration file with the buddys to write to
        Open App.Path & "\conf\buddy.info" For Output As #intFL
            'if there are some buddys to add...
            If lstBdyNk.ListCount > 0 Then
                '...go through the list
                For i = 0 To lstBdyNk.ListCount - 1
                    'and save each buddy's
                    'nick
                    Print #intFL, lstBdyNk.List(i)
                    'and welcome text
                    lstBdyWT.List(i) = Replace(lstBdyWT.List(i), ",", ChrW$(1))
                    Print #intFL, lstBdyWT.List(i)
                    lstBdyWT.List(i) = Replace(lstBdyWT.List(i), ChrW$(1), ",")
                'next buddy in the list
                Next i
            End If
        'close file and save changes
        Close #intFL
        intFL = FreeFile
        Open App.Path & "\conf\buddies\" & buddyname & ".info" For Output As #intFL
            Print #intFL, "Real Name: x"
            Print #intFL, "Version: x"
            Print #intFL, "Last Seen: x"
            Print #intFL, "Additional Info: x"
        Close #intFL
    End If
End Sub
Private Sub cmdApply_Click()
    'Applying settings
    Dim Window As Form 'an object variable used to store the current window when we loop through all the windows
    Dim t As Long
    
    DB.Enter "cmdApply_Click"
    
    Screen.MousePointer = vbHourglass
    
    DB.X "Checking for User Data"
    
    If LenB(txtNickname.Text) = 0 Or LenB(txtEmail.Text) = 0 Or LenB(txtReal.Text) = 0 Then
        DB.X "No User Data"
        'user information is missing; prompt him/her to fill it in later
        MsgBox Language(129), vbInformation, Language(103)
        'also select the appropriate option category
        Set lvList.SelectedItem = lvList.ListItems(1)
        'and display it
        'lvList_Click
    End If
    
    DB.X "Checking port ranges"
    If Val(TxtRangeL.Text) > Val(TxtRangeH.Text) Then
        'alert the user that their port range is incorrect
        MsgBox Language(795), vbInformation, Language(794)
        DB.X "Bad port ranges."
        Screen.MousePointer = vbDefault
        'do not save settings, port ranges must be corrected
        Exit Sub
    End If
    DB.X "Port ranges ok"
    
    DB.X "Checking Display Settings"
    If InStr(1, txtDisplayNormal.Text, "%nick") = 0 Then
        MsgBox Language(924), vbInformation, Language(923)
        Exit Sub
    End If
    DB.X "Display Settings OK"
    
    DB.X "Saving All Settings"
    'save all settings
    't = GetTickCount()
    SaveAll
    'Debug.Print "SaveAll Time: " & GetTickCount() - t
       
    'if the XP-style setting has changed...
    If chkXPCommonControls.value <> IIf(Options.XPCommonControls, vbChecked, vbUnchecked) Then
        'prompt that restart is necessary
        If MsgBox(Language(169) & " " & Language(170), vbQuestion Or vbYesNo, Language(71)) = vbYes Then
            'if the user agrees to restart
            'reload the application
            Shell """" & App.Path & "\" & App.EXEName & ".exe"""
            'set restarting variable to true
            Restarting = True
            'go through the loaded windows
            For Each Window In Forms
                'and unload each
                Unload Window
            'go to the next window
            Next Window
            'end the program
            End
        End If
    'if the language setting was changed...
    ElseIf Replace(Strings.LCase$(Options.LanguageFile), "/", "\") <> Replace(Strings.LCase$(icLanguage.SelectedItem.Tag), "/", "\") Then
        If GetSetting(App.EXEName, "InfoTips", "LangChange") = "0" Then
            'show tip
            frmMain.ShowInfoTip LangChange
            SaveSetting App.EXEName, "InfoTips", "LangChange", "1"
        End If
        'prompt that restart is necessary
        If MsgBox(Language(160) & " " & Language(170), vbQuestion Or vbYesNo, Language(78)) = vbYes Then
            'if the user agrees to restart
            'load all settings
            LoadAll
            'hide apply/Ok buttons
            cmdApply.Visible = False
            cmdOK.Visible = False
            'hide the options form
            Me.Hide
            'call the soft reload method; it will show the options dialog and make Apply/OK visible again
            Reload
            'do not load settings again
            Exit Sub
        End If
    'if the skin setting was changed...
    ElseIf Replace(Strings.LCase$(Skins(cboSkin.ListIndex)), "/", "\") <> Replace(Strings.LCase$(ThisSkin.FileName), "/", "\") Then
        If GetSetting(App.EXEName, "InfoTips", "SkinChange") = "0" Then
            frmMain.ShowInfoTip SkinChange
            SaveSetting App.EXEName, "InfoTips", "SkinChange", "1"
        End If
        If MsgBox(Language(215) & " " & Language(170), vbQuestion Or vbYesNo, Language(166)) = vbYes Then
            'if the user agrees to restart
            'load all settings
            LoadAll
            'hide apply/Ok buttons
            cmdApply.Visible = False
            cmdOK.Visible = False
            'hide the options form
            Me.Hide
            'call the soft reload method; it will show the options dialog and make Apply/OK visible again
            Reload
            'do not load settings again
            Exit Sub
        End If
    End If
    
    DB.X "Loading All Settings"
    'load settings(the new ones)
    't = GetTickCount()
    LoadAll
    'Debug.Print "LoadAll Time: " & GetTickCount() - t
    
    Screen.MousePointer = vbDefault
    
    DB.Leave "cmdApply_Click"
End Sub
Private Sub cmdCancel_Click()
    'canceled: don't save; just hide the options dialog.
    Me.Hide
End Sub
Private Sub cmdClear_Click(Index As Integer)
    'the ignore list is being cleared
    'ask first
    If MsgBox(Language(626), vbYesNo Or vbQuestion, Language(105)) = vbYes Then
        'answered yes; clear the list
        lstIgnore(Index).Clear
    End If
End Sub
Private Sub cmdDeleteScript_Click()
    'delete a script
    'if there is a script seelcted...
    If flScripts.ListIndex > 0 Then
        'ask the user before deleting the file...
        If MsgBox(Language(118), vbQuestion Or vbYesNo, Language(197)) = vbYes Then
            'if he/she agress delete it
            Kill flScripts.Path & "\" & flScripts.FileName
            'refresh the file list
            flScripts.Refresh
        End If
    'if there's nothing selected
    Else
        'display warning
        MsgBox Language(225), vbInformation, Language(87)
    End If
End Sub
Private Sub cmdEdit_Click()
    'the user wants to edit a script
    Dim ScriptToEdit As String 'the filename of the script to edit
    ScriptToEdit = flScripts.Path & "\" & flScripts.FileName 'the script to edit is the selected script
    'if there is a script seelcted...
    If flScripts.ListIndex >= 0 Then
        'edit the script; if the filename ends to `.xml' set the last argument to true.
        EditScript ScriptToEdit, Strings.LCase$(Strings.Right$(ScriptToEdit, Len(".xml"))) = ".xml"
    'if there's nothing selected
    Else
        'display warning
        MsgBox Language(225), vbInformation, Language(87)
    End If
End Sub

Private Sub cmdEditBdy_Click()
    'Edit Buddy
    Dim i As Integer 'a counter variable for the loops
    Dim intFL As Integer 'variable to store the index of the configuration file
    Dim boolSelected As Boolean 'if at least one buddy was selected
    Dim strTemp 'a tmeporary string
    
    'go through the buddy's list to see which items are selected
    'starting from buddy zero
    i = 0
    Do
        'if there are no more buddys
        If i = lstBdyNk.ListCount Then
            'end the loop
            Exit Do
        End If
        'if the current item is selected
        If lstBdyNk.Selected(i) Then
            'Show prompts to edit Nickname and welcome text
            strTemp = InputBox(Language(142), Language(89), lstBdyNk.List(i))
            If LenB(strTemp) > 0 Then
                lstBdyNk.List(i) = strTemp
            Else
                Exit Sub
            End If
            strTemp = InputBox(Language(143), Language(89), lstBdyWT.List(i))
            lstBdyWT.List(i) = strTemp
            'and refresh the two lists
            lstBdyNk.Refresh
            lstBdyWT.Refresh
            'at least one buddy was selected; set it to true
            boolSelected = True
        End If
        'moving to the next buddy
        i = i + 1
    Loop
    
    'if there were no buddys selected at all...
    If Not boolSelected Then
        '...display warning
        MsgBox Language(226), vbInformation, Language(130)
        'and don't save anything
        Exit Sub
    End If
    
    'someone has been changed, we need to save changed
    'get a free file index
    intFL = FreeFile
    'open the configuration file
    Open App.Path & "\conf\buddy.info" For Output As #intFL
        'if there's something in the list to save...
        If lstBdyNk.ListCount > 0 Then
            '...go through the list
            For i = 0 To lstBdyNk.ListCount - 1
                'and save its contents
                Print #intFL, lstBdyNk.List(i)
                lstBdyWT.List(i) = Replace(lstBdyWT.List(i), ",", ChrW$(1))
                Print #intFL, lstBdyWT.List(i)
                lstBdyWT.List(i) = Replace(lstBdyWT.List(i), ChrW$(1), ",")
            Next i
        End If
    'clode and save the configuration file
    Close #intFL
End Sub
Private Sub cmdExport_Click()
    On Error GoTo Cancel_was_selected
    cdExport.ShowSave
    If FS.FileExists(cdExport.FileName) Then
        If MsgBox(Language(940), vbExclamation Or vbYesNo, Language(633)) = vbNo Then
            Exit Sub
        End If
    End If
    ExportSettings cdExport.FileName
Cancel_was_selected:
End Sub
Private Sub cmdImport_Click()
    On Error GoTo Cancel_was_selected
    cdImport.ShowSave
    ImportSettings cdImport.FileName
Cancel_was_selected:
End Sub
Public Sub ExportSettings(ByVal FileName As String)
    'export all settings
    
    Dim XMLDoc As DOMDocument
    Dim XMLNode As IXMLDOMElement
    Dim XMLVersion As IXMLDOMProcessingInstruction
    Dim XMLComment As IXMLDOMComment
    
    'Create a new DOM Document
    Set XMLDoc = New DOMDocument
    
    'Init the DOM Document
    XMLDoc.async = False
    XMLDoc.validateOnParse = False
    XMLDoc.resolveExternals = False
    XMLDoc.preserveWhiteSpace = True
    
    'Create the ProcessingInstruction, this is <?xml version='1.0'>
    'That appears at the first line of the XML file
    Set XMLVersion = XMLDoc.createProcessingInstruction("xml", "version='1.0'")
    
    'and add it to the Document
    XMLDoc.appendChild XMLVersion
    
    Set XMLVersion = Nothing
    Set XMLComment = XMLDoc.createComment("This file was automatically created by Node")
    
    XMLDoc.appendChild XMLComment
    
    Set XMLComment = Nothing
    
    'build the Document Element
    Set XMLNode = XMLDoc.createElement("NodeSettings")
    
    'save the version of Node using an attribute of the Document Element
    XMLNode.setAttribute "Vesion", App.Major & "." & App.Minor
    
    'set the root element
    Set XMLDoc.documentElement = XMLNode
        
    'save all conf files
    Export_CreateFileNode XMLDoc, "info.dat"
    Export_CreateFileNode XMLDoc, "ignore.dat"
    Export_CreateFileNode XMLDoc, "alias.dat"
    Export_CreateFileNode XMLDoc, "buddy.info"
    Export_CreateFileNode XMLDoc, "channels.lst"
    Export_CreateFileNode XMLDoc, "ignore_nick.lst"
    Export_CreateFileNode XMLDoc, "perform.dat"
    Export_CreateFileNode XMLDoc, "servers.lst"
    Export_CreateFileNode XMLDoc, "favwebs.lst"
    Export_CreateFileNode XMLDoc, "hotkeys.dat"
    
    'save all registry settings
    Export_XMLAppendRegistrySettings XMLDoc, "Options"
    Export_XMLAppendRegistrySettings XMLDoc, "Options\Accessibility\Narration"
    
    'save the result
    XMLDoc.save FileName

    'unload the XML objects
    Set XMLDoc = Nothing
    Set XMLNode = Nothing
End Sub
Public Sub Export_CreateFileNode(ByRef XMLDoc As DOMDocument, ByVal FileName As String)
    Dim XMLNode As IXMLDOMElement
    
    'create UserInfo node
    Set XMLNode = XMLDoc.createElement("ConfigFile")
    
    XMLNode.setAttribute "FileName", FileName
        
    'add it as Text to XMLNode
    XMLNode.Text = GetFile(App.Path & "/conf/" & FileName)
    
    'append it to our document
    XMLDoc.documentElement.appendChild XMLNode
End Sub
Public Sub Export_XMLAppendRegistrySettings(ByRef XMLDoc As DOMDocument, ByVal RegKey As String)
    Dim cNodeSettings As Collection
    Dim intRegistrySettingIndex As Integer
    Dim XMLNode As IXMLDOMElement
    
    Set cNodeSettings = EnumerateRegistryValues(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\Node\" & RegKey)
    
    'append each one to the XML file
    For intRegistrySettingIndex = 1 To cNodeSettings.Count
        Set XMLNode = XMLDoc.createElement("NodeSetting")
        XMLNode.setAttribute "ValuePath", RegKey
        XMLNode.setAttribute "ValueName", cNodeSettings.Item(intRegistrySettingIndex)(0)
        XMLNode.setAttribute "ValueData", cNodeSettings.Item(intRegistrySettingIndex)(1)
        XMLDoc.documentElement.appendChild XMLNode
    Next intRegistrySettingIndex
End Sub
Public Sub ImportSettings(ByVal FileName As String)
    'import all settings
    
    Dim XMLDoc As DOMDocument
    Dim XMLNode As IXMLDOMElement
    Dim XMLAttribute As IXMLDOMAttribute
    Dim strFileName As String
    Dim strFileData As String
    Dim strRegKey As String
    Dim strRegValue As String
    Dim strRegData As String
    Dim i As Integer
    
    'Create a new DOM Document Object
    Set XMLDoc = New DOMDocument
    
    'Load
    If Not XMLDoc.Load(FileName) Then
        'invalid xml file
        MsgBox Language(752), vbCritical, Language(744)
        Exit Sub
    End If
    
    'build the Document Element
    Set XMLNode = XMLDoc.documentElement
    
    If XMLNode.nodeName <> "NodeSettings" Then
        'invalid root element
        MsgBox Language(753), vbCritical, Language(744)
    End If
    
    For i = 0 To XMLNode.Attributes.length - 1
        If XMLNode.Attributes.Item(i).nodeName = "Vesion" Then
            If XMLNode.Attributes.Item(i).nodeValue <> App.Major & "." & App.Minor Then
                'invalid XML settings file version
                '(a later or older version of Node)
                If MsgBox(Language(753), vbQuestion Or vbYesNo, Language(744)) = vbNo Then
                    Exit Sub
                End If
            Else
                Exit For
            End If
        End If
    Next
        
    For Each XMLNode In XMLDoc.documentElement.childNodes
        Select Case XMLNode.nodeName
            Case "ConfigFile"
                strFileName = vbNullString
                For i = 0 To XMLNode.Attributes.length - 1
                    If XMLNode.Attributes.Item(i).nodeName = "FileName" Then
                        strFileName = XMLNode.Attributes.Item(i).nodeValue
                        Exit For
                    End If
                Next i
                If LenB(strFileName) > 0 Then
                    strFileData = XMLNode.Text
                    SetFile App.Path & "/conf/" & strFileName, strFileData
                End If
            Case "NodeSetting"
                strRegKey = vbNullString
                strRegValue = vbNullString
                strRegData = vbNullString
                For i = 0 To XMLNode.Attributes.length - 1
                    Select Case XMLNode.Attributes.Item(i).nodeName
                        Case "ValuePath"
                            strRegKey = XMLNode.Attributes.Item(i).nodeValue
                        Case "ValueName"
                            strRegValue = XMLNode.Attributes.Item(i).nodeValue
                        Case "ValueData"
                            strRegData = XMLNode.Attributes.Item(i).nodeValue
                    End Select
                Next i
                If LenB(strRegKey) > 0 And LenB(strRegValue) > 0 Then
                    SaveSetting "Node", strRegKey, strRegValue, strRegData
                End If
        End Select
    Next XMLNode
    
    'unload the XML objects
    Set XMLDoc = Nothing
    Set XMLNode = Nothing
    
    'load new settings
    LoadAll
End Sub
Private Sub LoadPerformServers()
    'add all server descriptions to the perform depending on server combo list
    Dim intFL As Integer
    Dim strServerDescription As String
    
    cboPerformServer.Clear
    cboConnectServer.Clear
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As intFL
        Do Until EOF(intFL)
            strServerDescription = GetParameter(xLineInput(intFL), 2)
            cboPerformServer.AddItem strServerDescription
            cboConnectServer.AddItem strServerDescription
        Loop
        Close intFL
    End If
End Sub
Private Function GetServerListItemDetailsFromIndex(ByVal Index As Integer, ByVal DetailID As Integer) As String
    Dim intFL As Integer
    Dim i As Integer
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As intFL
        Do Until EOF(intFL)
            If i = Index Then
                GetServerListItemDetailsFromIndex = GetParameter(xLineInput(intFL), DetailID)
                Exit Do
            Else
                xLineInput intFL
            End If
            i = i + 1
        Loop
        Close intFL
    End If
End Function
Private Function GetServerListItemIndexFromHostname(ByVal strHostname As String) As Integer
    Dim intFL As Integer
    Dim i As Integer
    
    If LenB(strHostname) = 0 Then
        i = -1
        Exit Function
    End If
    
    If FS.FileExists(App.Path & "\conf\servers.lst") Then
        intFL = FreeFile
        Open App.Path & "\conf\servers.lst" For Input Access Read Shared As intFL
        Do Until EOF(intFL)
            If strHostname = GetStatement(xLineInput(intFL)) Then
                GetServerListItemIndexFromHostname = i
                Exit Do
            End If
            i = i + 1
        Loop
        Close intFL
    End If
End Function
Private Sub cmdOK_Click()
    'The user clicked OK.
    'Apply settings
    cmdApply_Click
    'and hide dialog
    cmdCancel_Click
End Sub
Private Sub cmdREBdy_Click()
    'Removing Buddy
    Dim i As Integer 'a counter variable for the loops
    Dim intFL As Integer 'variable to store the index of the configuration file
    Dim boolSelected As Boolean 'if at least one buddy was selected
    Dim buddyname As String 'this is the buddy's nickname that is to be removed
    Dim FS As FileSystemObject
    
    'go through the buddy's list to see which items are selected
    'starting from buddy zero
    i = 0
    Do
        'if there are no more buddys
        If i = lstBdyNk.ListCount Then
            'end the loop
            Exit Do
        End If
        'if the current item is selected
        If lstBdyNk.Selected(i) Then
            'set buddyname = the nickname to be removed
            buddyname = lstBdyNk.List(i)
            'remove both the nickname and the text
            lstBdyNk.RemoveItem i
            lstBdyWT.RemoveItem i
            'and refresh the two lists
            lstBdyNk.Refresh
            lstBdyWT.Refresh
            'at least one buddy was selected; set it to true
            boolSelected = True
            'go to the previous buddy, as the current one was removed
            i = i - 1
        End If
        'moving to the next buddy
        i = i + 1
    Loop
    
    'if there were no buddys selected at all...
    If Not boolSelected Then
        '...display warning
        MsgBox Language(226), vbInformation, Language(130)
        'and don't save anything
        Exit Sub
    End If
    
    'someone was removed, we need to save changed
    'get a free file index
    intFL = FreeFile
    'open the configuration file
    Open App.Path & "\conf\buddy.info" For Output As #intFL
        'if there's something in the list to save...
        If lstBdyNk.ListCount > 0 Then
            '...go through the list
            For i = 0 To lstBdyNk.ListCount - 1
                'and save its contents
                Print #intFL, lstBdyNk.List(i)
                Print #intFL, lstBdyWT.List(i)
            Next i
        End If
    'clode and save the configuration file
    Close #intFL
    
    Set FS = New FileSystemObject
    If FS.FileExists(App.Path & "/conf/buddies/" & buddyname & ".info") Then Kill App.Path & "/conf/buddies/" & buddyname & ".info"
End Sub
Private Sub cmdRemove_Click(Index As Integer)
    'someone is being removed from the ignore list.
    Dim Result As Integer
    If lstIgnore(Index).ListIndex = -1 Then
        'there isn't anyone selected; skip procedure
        MsgBox Language(106), vbInformation, Language(107)
    Else
        'ask first whether the user is sure to remove him or her.
        If MsgBox(Language(118), vbYesNo + vbQuestion, Language(119)) = vbYes Then
            'answered yes: remove.
            lstIgnore(Index).RemoveItem lstIgnore(0).ListIndex
        End If
    End If
End Sub
Private Sub cmdRemoveHotKey_Click()
    HotKeysRemove fgHotKeys.TextMatrix(fgHotKeys.Row, 0)
    
    If fgHotKeys.rows > 2 Then
        fgHotKeys.RemoveItem fgHotKeys.Row
    Else
        fgHotKeys.TextMatrix(1, 0) = vbNullString
        fgHotKeys.TextMatrix(1, 1) = vbNullString
    End If
End Sub
Private Sub LoadHotKeys()
    Dim strTemp As String
    Dim intFL As Integer
    Dim i As Integer
    
    'TO DO: convert IDs to the correct language captions
    
    fgHotKeys.Clear
    
    For i = 3 To fgHotKeys.rows
        fgHotKeys.RemoveItem 1
    Next i
    
    fgHotKeys.TextMatrix(0, 0) = Language(490)
    fgHotKeys.TextMatrix(0, 1) = Language(491)
    
    intFL = FreeFile
    
    If Not FS.FileExists(App.Path & "\conf\hotkeys.dat") Then
        intFL = FreeFile
        DB.X "hotkeys.dat does not exist"
        DB.X "Creating hotkeys.dat"
        Open App.Path & "\conf\hotkeys.dat" For Output As intFL
        Close intFL
    End If
    
    Open App.Path & "/conf/hotkeys.dat" For Input Access Read Shared As #intFL
    Do Until EOF(intFL)
        Line Input #intFL, strTemp
        fgHotKeys.AddItem GetStatement(strTemp) & vbTab & GetParameter(strTemp)
    Loop
    If LenB(fgHotKeys.TextMatrix(1, 1)) = 0 Then
        If fgHotKeys.rows > 2 Then
            fgHotKeys.RemoveItem 1
        End If
    End If
    Close #intFL
End Sub
Private Sub fgPlugins_DblClick()
    If LoadPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll") Then
        fgPlugins.TextMatrix(fgPlugins.Row, 1) = Language(280) 'Loaded: Yes
    Else
        'failed
        'display warning
        MsgBox Replace(Language(528), "%1", Err.Description), vbCritical, Language(529)
    End If
End Sub
Private Sub fgPlugins_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intRowNum As Integer
    
    'get the plugin which the mouse is over(Exclude #1)
    intRowNum = Y \ fgPlugins.RowHeight(1) - 1
    'add TopRow index in the case the user has scrolled
    intRowNum = intRowNum + fgPlugins.TopRow
    'select it
    On Error Resume Next
    fgPlugins.Row = intRowNum
    
    If intRowNum = 0 Then
        Exit Sub
    End If
    
    If Button = 2 Then
        'right-click
        
        If Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(intRowNum, 0) & ".dll")).boolLoaded Then
            mnuLoadPlugin.Visible = False
            mnuUnloadPlugin.Visible = True
        Else
            mnuLoadPlugin.Visible = True
            mnuUnloadPlugin.Visible = False
        End If
        mnuPluginStartup.Checked = Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(intRowNum, 0) & ".dll")).boolLoadOnStartup
        PopupMenu mnuPlugIn
    End If
End Sub

Private Sub flScripts_DblClick()
    cmdEdit_Click
End Sub
Private Sub Form_Initialize()
    'load XP common controls
    InitCommonControls
    'set the scripts path so the files are displayed correctly
    flScripts.Path = App.Path & "\scripts"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'so the skin loaded is not lost.
        Me.Hide
        Cancel = True
    End If
End Sub
Private Sub lblDownloadLang_Click()
    'the user wants to download more languages
    'navigate him/her to the languages page
    CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=languages"
    frmOptions.Hide
End Sub
Private Sub lblDownloadSoundSchemes_Click()
    'the user wants to download more sound schemes
    'navigate him/her to the sound schemes page
    CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=soundschemes"
    frmOptions.Hide
End Sub
Private Sub lblDownloadPlugs_Click()
    'the user wants to download more plugins
    'navigate him/her to the plugins page
    CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=plugins"
    frmOptions.Hide
End Sub
Private Sub lblDownloadSkins_Click()
    'the user wants to download more skins
    'navigate him/her to the skins page
    CurrentActiveServer.preExecute "/browse http://node.sourceforge.net/link.php?p=skins"
    frmOptions.Hide
End Sub
Private Sub lblNarration_Click()
    chkNarration.value = -(Not CBool(chkNarration.value))
End Sub
Private Sub lblNarrationChannels_Click()
    chkNarrationChannels.value = -(Not CBool(chkNarrationChannels.value))
End Sub
Private Sub lblNarrationInterface_Click()
    chkNarrationInterface.value = -(Not CBool(chkNarrationInterface.value))
End Sub
Private Sub lblNarrationPrivates_Click()
    chkNarrationPrivates.value = -(Not CBool(chkNarrationPrivates.value))
End Sub
Private Sub lblNarrationStatus_Click()
    chkNarrationStatus.value = -(Not CBool(chkNarrationStatus.value))
End Sub
Private Sub lblNickLinkChan_Click()
    chkNickLinkChan.value = -(Not CBool(chkNickLinkChan.value))
End Sub
Private Sub lblNickLinkMineChan_Click()
    chkNickLinkMineChan.value = -(Not CBool(chkNickLinkMineChan.value))
End Sub
Private Sub lblNickLinkMinePriv_Click()
    chkNickLinkMinePriv.value = -(Not CBool(chkNickLinkMinePriv.value))
End Sub
Private Sub lblNickLinkPriv_Click()
    chkNickLinkPriv.value = -(Not CBool(chkNickLinkPriv.value))
End Sub
Private Sub lblOption_Click(Index As Integer)
    Dim objOption As OptionButton
    
    Select Case Index
        Case 0
            Set objOption = optDCCB(0)
        Case 1
            Set objOption = optDCCB(1)
        Case 2
            Set objOption = optDCCB(2)
        Case 3
            Set objOption = optDCCI(0)
        Case 4
            Set objOption = optDCCI(1)
        Case 5
            Set objOption = optDCCI(2)
        Case 6
            Set objOption = optDCCE(0)
        Case 7
            Set objOption = optDCCE(1)
        Case 8
            Set objOption = optDCCE(2)
        Case 9
            Set objOption = optSessionNR
        Case 10
            Set objOption = optSessionND
        Case 11
            Set objOption = optSessionNA
        Case 12
            Set objOption = optSessionCR
        Case 13
            Set objOption = optSessionCD
        Case 14
            Set objOption = optSessionCA
    End Select
    objOption.value = True
End Sub
Private Sub lblPerformMultiple_Click()
    RTB_LostFocus
    optPerformMultiple.value = True
End Sub
Private Sub lblPerformSingle_Click()
    RTB_LostFocus
    optPerformSingle.value = True
End Sub
Private Sub lblSelectionClear_Click()
    chkSelectionClear.value = -(Not CBool(chkSelectionClear.value))
End Sub
Private Sub lblSelectionCopy_Click()
    chkSelectionCopy.value = -(Not CBool(chkSelectionCopy.value))
End Sub
Private Sub lblViewDownloads_Click()
    Shell "explorer " & App.Path & "\downloads", vbMaximizedFocus
End Sub
Private Sub lblViewLogs_Click()
    Shell "explorer " & App.Path & "\logs", vbMaximizedFocus
End Sub
Private Sub lstBdyNk_Click()
    'the user clicked on a buddy in the buddy list
    'also select the appropriate welcome text
    On Error Resume Next
    lstBdyWT.ListIndex = lstBdyNk.ListIndex
End Sub
Private Sub lstBdyWT_Click()
    'the user clicked on a welcome text in the buddy list
    'also select the appropriate buddy
    On Error Resume Next
    lstBdyNk.ListIndex = lstBdyWT.ListIndex
End Sub
Private Sub Form_Load()
    Dim Control As Object
    
    'The options form is being loaded
    
    DB.Enter "frmOptions.Form_Load"

    DB.X "LoadAll"
    'load all options
    LoadAll
    
    DB.X "OptionsLoadLanguage"
    'the options dialog is being loaded
    'load the captions of the controls on the dialog
    OptionsLoadLanguage
    
    DB.X "LoadOptionsList"
    'here we add all the options categories
    LoadOptionsList
    
    DB.X "FixPositions"
    'the first frame should be the default-selected
    'select it
    'Set lvList.SelectedItem = lvList.ListItems.Item(1)
    'update the frame on the right
    'lvList_Click

    FixPositions 'call this procedure which changes the position of the controls
                 'on the form in order to make it appear as it appears.
    
    DB.X "Selecting default Options Category"
    Set tvOptions.SelectedItem = tvOptions.Nodes.Item(1)
    tvOptions_NodeClick tvOptions.SelectedItem

    For Each Control In Controls
        If TypeOf Control Is Label Then
            Control.Font.Charset = LangCharSet
        End If
    Next Control
    
    DB.Leave "frmOptions.Form_Load"
End Sub
Public Sub LoadOptionsList()
    Dim i As Integer
    Dim nodeRoot As ComctlLib.Node
    Dim nodeSection As ComctlLib.Node
    Dim nodeSubSection As ComctlLib.Node
    
    DB.Enter "LoadOptionsList"
    'Set lvList.Icons = ilIcons
    
    For i = 1 To tvOptions.Nodes.Count
        tvOptions.Nodes.Remove i
        'lvList.ListItems.Remove 1
    Next i
    
    With tvOptions.Nodes
        Set nodeRoot = .Add(, tvwFirst, "root", Language(4))
            Set nodeSection = .Add(nodeRoot, tvwChild, "s01", Language(319))  'IRC
                .Add nodeSection, tvwChild, "k00", Language(72) ', 1 '"User"
                .Add nodeSection, tvwChild, "k02", Language(75) ', 4 '"Logs"
                .Add nodeSection, tvwChild, "k03", Language(76) ', 5 '"Perform"
                .Add nodeSection, tvwChild, "k18", Language(743) ', "TimeStamps"
                Set nodeSubSection = .Add(nodeSection, tvwChild, "k24", Language(900))  ''"Away System"
                    .Add nodeSubSection, tvwChild, "k27", Language(929) 'Away Actions
                    .Add nodeSubSection, tvwChild, "k28", Language(930) 'Back Actions
                .Add nodeSection, tvwChild, "k06", Language(130) ', 8 '"Buddy List"
                .Add nodeSection, tvwChild, "k16", Language(654) ', '"General"
            Set nodeSection = .Add(nodeRoot, tvwChild, "s02", Language(598))  'Personalization
                .Add nodeSection, tvwChild, "k05", Language(78) ', 7 '"Language"
                .Add nodeSection, tvwChild, "k07", Language(166) ', 9 '"Skin"
                '.Add nodeSection, tvwChild, "k08", Language(156) ', 10 '"Smileys"
                .Add nodeSection, tvwChild, "k11", Language(489) ', 13 '"Hot Keys"
                .Add nodeSection, tvwChild, "k17", Language(720) ', '"Accessibility"
                .Add nodeSection, tvwChild, "k22", Language(866) ', "Sound Scheme"
                .Add nodeSection, tvwChild, "k23", Language(886) ', "Smiley Packs"
                .Add nodeSection, tvwChild, "k26", Language(923) ', "Display"
            Set nodeSection = .Add(nodeRoot, tvwChild, "s03", Language(600))  'Connection
                .Add nodeSection, tvwChild, "k10", Language(309) ', 12 '"DCC/NDC Options"
                .Add nodeSection, tvwChild, "k13", Language(573) ', 15 '"Proxy"
                .Add nodeSection, tvwChild, "k21", Language(776) ', "Retry"
                .Add nodeSection, tvwChild, "k08", Language(941) ' "Ident"
            Set nodeSection = .Add(nodeRoot, tvwChild, "s05", Language(622))  'Security
                .Add nodeSection, tvwChild, "k04", Language(77) ', 6 '"Ignore"
                .Add nodeSection, tvwChild, "k15", Language(341) ', 12 '"CTCP"
            Set nodeSection = .Add(nodeRoot, tvwChild, "s04", Language(601))  'Other Options
                .Add nodeSection, tvwChild, "k01", Language(74) ', 3 '"Scripts"
                .Add nodeSection, tvwChild, "k09", Language(270) ', 11 '"Sessions"
                .Add nodeSection, tvwChild, "k12", Language(507) ', 14 '"Plug-Ins"
                .Add nodeSection, tvwChild, "k20", Language(756) ', '"Startup"
                .Add nodeSection, tvwChild, "k25", Language(908) '"Browsing"
                .Add nodeSection, tvwChild, "k14", Language(161) ', 16 '"Misc"
                
                'leave this one last
                .Add nodeSection, tvwChild, "k19", Language(744) ', "Import/Export"
    End With
    
    nodeRoot.Expanded = True

    For i = 0 To fraOptions.Count - 1
        'fraOptions(i).Caption = lvList.ListItems.Item(i + 1).Text
    Next i

    DB.Leave "LoadOptionsList"
End Sub
Public Sub SaveAll()
    'this sub is used to save all options
    Dim intFL As Integer 'a variable to store the free file index
    Dim i As Integer 'a counter variable for the loops
    Dim strServerHostname As String
    
    DB.Enter "SaveAll()"
    
    DB.X "Saving info.dat"
    'get a free file index
    intFL = FreeFile
    'open the personal information file and write in
    Open App.Path & "\conf\info.dat" For Output As #intFL
        'clear the file and write the new information
        'nickname, alternative, alternative 2, real name and email address
        Print #intFL, txtNickname.Text
        Print #intFL, txtAlt.Text
        Print #intFL, txtAltTwo.Text
        Print #intFL, txtEmail.Text
        Print #intFL, txtReal.Text
        'close the file
    Close #intFL
        
    DB.X "Saving ignore_nick.dat"
    
    'save the ignore list items
    'get a free file index
    intFL = FreeFile
    'open the ignore list file
    Open App.Path & "\conf\ignore_nick.dat" For Output As #intFL
        'clear the current file, and then go throught the list and save each item
        For i = 0 To lstIgnore(0).ListCount - 1
            'save the current item to the file
            Print #intFL, lstIgnore(0).List(i)
        Next i
    'close the file - save changes
    Close #intFL
    
    DB.X "Saving ignore_ip.dat"
    
    'get a free file index
    intFL = FreeFile
    'open the ignore list file
    Open App.Path & "\conf\ignore_ip.dat" For Output As #intFL
        'clear the current file, and then go throught the list and save each item
        For i = 0 To lstIgnore(1).ListCount - 1
            'save the current item to the file
            Print #intFL, lstIgnore(1).List(i)
        Next i
    'close the file - save changes
    Close #intFL
       
    DB.X "Saving Log Settings"
    
    'now we save the logger options to registry
    'if it's checked, save a True value, else save a false value
    '(this is caused by the result of the comparison between the Value Property
    ' and the constant vbChecked, which returns either True or False)
    SaveSetting App.EXEName, "Options", "LogChannels", chkLogChannels.value = vbChecked
    SaveSetting App.EXEName, "Options", "LogPrivates", chkLogPrivates.value = vbChecked
    SaveSetting App.EXEName, "Options", "TimeStamp/Log", chkTimeStampsLog.value = vbChecked
    SaveSetting App.EXEName, "Options", "Debug\Raw", chkLogRAW.value = vbChecked
    
    DB.X "Saving Browsing Settings"
    
    SaveSetting App.EXEName, "Options", "BrowseInternal", optBrowseInternal.value
    SaveSetting App.EXEName, "Options", "BrowseParseLinks", chkBrowseParseLinks.value = vbChecked
    
    DB.X "Saving Display Settings"
    
    SaveSetting App.EXEName, "Options", "DisplayNormal", txtDisplayNormal.Text
    SaveSetting App.EXEName, "Options", "DisplaySkinFontSize", txtDisplaySkinFontSize.Text

    
    DB.X "Saving Misc Settings"
    
    'here we use the same method to save the Latest Version option, the Fade and the Perform options
    SaveSetting App.EXEName, "Options", "CheckLatest", chkLatest.value = vbChecked
    SaveSetting App.EXEName, "Options", "Fade", chkFade.value = vbChecked
    SaveSetting App.EXEName, "Options", "Smileys", chkEnableSmileys.value = vbChecked
    SaveSetting App.EXEName, "Options", "Scripting", chkScripting.value = vbChecked
    SaveSetting App.EXEName, "Options", "CodeBehind", chkCodeBehind.value = vbChecked
    SaveSetting App.EXEName, "Options", "HTMLLoading", chkLoading.value = vbChecked
    SaveSetting App.EXEName, "Options", "HTMLError", chkHTMLError.value = vbChecked
    SaveSetting App.EXEName, "Options", "AutoComplete", chkAutocomplete.value = vbChecked
    SaveSetting App.EXEName, "Options", "TOD", chkTOD.value = vbChecked
    SaveSetting App.EXEName, "Options", "LogByNet", chkLogByNet.value = vbChecked
    SaveSetting App.EXEName, "Options", "InfoTips", chkInfoTips.value = vbChecked
    SaveSetting App.EXEName, "Options", "Tray", chkTray.value = vbChecked
    SaveSetting App.EXEName, "Options", "RestoreStatus", chkRestoreStatus.value = vbChecked
    SaveSetting App.EXEName, "Options", "AutoNDC", chkAutoNDC.value = vbChecked
    
    SaveSetting App.EXEName, "Options", "PortRangeLow", TxtRangeL.Text
    SaveSetting App.EXEName, "Options", "PortRangeHigh", TxtRangeH.Text
    
    SaveSetting App.EXEName, "Options", "JoinPanel", chkJoinPanel.value = vbChecked
       
    SaveSetting App.EXEName, "Options", "SelectCopy", chkSelectionCopy.value = vbChecked
    SaveSetting App.EXEName, "Options", "SelectClear", chkSelectionClear.value = vbChecked
           
    SaveSetting App.EXEName, "Options", "NickLinkChan", chkNickLinkChan.value = vbChecked
    SaveSetting App.EXEName, "Options", "NickLinkPriv", chkNickLinkPriv.value = vbChecked
    SaveSetting App.EXEName, "Options", "NickLinkMineChan", chkNickLinkMineChan.value = vbChecked
    SaveSetting App.EXEName, "Options", "NickLinkMinePriv", chkNickLinkMinePriv.value = vbChecked
    
    SaveSetting App.EXEName, "Options", "RejoinOnKick", chkJoinOnInvite.value = vbChecked
    SaveSetting App.EXEName, "Options", "JoinOnInvite", chkJoinOnKick.value = vbChecked
    SaveSetting App.EXEName, "Options", "ModesOnJoin", chkModesOnJoin.value = vbChecked
    SaveSetting App.EXEName, "Options", "KeepChannelsOpen", chkKeepChannelsOpen.value = vbChecked

    DB.X "Saving Perform Options"
    'the perform itself is saved when it's changed
    SaveSetting App.EXEName, "Options", "EnablePerform", chkPerform.value = vbChecked
    SaveSetting App.EXEName, "Options", "PerformSingle", optPerformSingle.value
    'save the current perform text in case it hasn't been saved
    RTB_LostFocus
    
    'startup: connect to server
    SaveSetting App.EXEName, "Options", "StartupConnect", chkStartupConnect.value = vbChecked
    strServerHostname = GetServerListItemDetailsFromIndex(cboConnectServer.ListIndex, 0) '0 = hostname
    If LenB(strServerHostname) > 0 And cboConnectServer.ListIndex <> -1 Then
        SaveSetting App.EXEName, "Options", "StartupConnectServer", strServerHostname
        SaveSetting App.EXEName, "Options", "StartupConnectPort", GetServerListItemDetailsFromIndex(cboConnectServer.ListIndex, 1) '1 = port
    Else
        SaveSetting App.EXEName, "Options", "StartupConnect", False
    End If
    
    DB.X "Saving Session Settings"
    
    SaveSetting App.EXEName, "Options", "SessionCrash", Abs(optSessionCD.value + optSessionCA.value * 2)
    SaveSetting App.EXEName, "Options", "SessionNormal", Abs(optSessionND.value + optSessionNA.value * 2)
    
    DB.X "Saving DCC/Connection Settings"
    
    SaveSetting App.EXEName, "Options", "DCC-Buddies", Abs(optDCCB(1).value + optDCCB(2).value * 2)
    SaveSetting App.EXEName, "Options", "DCC-Ignore", Abs(optDCCI(1).value + optDCCI(2).value * 2)
    SaveSetting App.EXEName, "Options", "DCC-Everyone", Abs(optDCCE(1).value + optDCCE(2).value * 2)
    
    SaveSetting App.EXEName, "Options", "AntiVirus", Options.DCCAntivirus
    SaveSetting App.EXEName, "Options", "Proxy/UseProxy", chkProxyEnable.value = vbChecked
      
    SaveSetting App.EXEName, "Options", "ConnectRetryDelay", txtRetryDelay.Text
    SaveSetting App.EXEName, "Options", "ConnectRetry", chkRetry.value
    
    DB.X "Saving Accessibility Settings"
    
    'Accessibility
    SaveSetting App.EXEName, "Options\Accessibility\Narration", "Enabled", chkNarration.value = vbChecked
    SaveSetting App.EXEName, "Options\Accessibility\Narration", "Interface", chkNarrationInterface.value = vbChecked
    SaveSetting App.EXEName, "Options\Accessibility\Narration", "Channels", chkNarrationChannels.value = vbChecked
    SaveSetting App.EXEName, "Options\Accessibility\Narration", "Privates", chkNarrationPrivates.value = vbChecked
    SaveSetting App.EXEName, "Options\Accessibility\Narration", "Status", chkNarrationStatus.value = vbChecked
    
    DB.X "Saving CTCP Settings"
    
    'CTCP
    SaveSetting App.EXEName, "Options", "CTCP/Ping/Reply", chkCTCPPing.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Ping/Ignore", chkCTCPPingIgnore.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Version/Reply", chkCTCPVersion.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Version/Ignore", chkCTCPVersionIgnore.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Version/Custom", chkCTCPVersionCustom.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Version/Message", txtCTCPVersionCustom.Text
    SaveSetting App.EXEName, "Options", "CTCP/Time/Reply", chkCTCPTime.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/Time/Ignore", chkCTCPTimeIgnore.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/FloodProtection", chkCTCPFlood.value = vbChecked
    SaveSetting App.EXEName, "Options", "CTCP/BounceToFlooders", chkCTCPFloodBounce.value = vbChecked

    DB.X "Saving TimeStamp Settings"
    
    'TimeStamp
    SaveSetting App.EXEName, "Options", "TimeStamp/Enable", chkTimeStamp.value = vbChecked
    SaveSetting App.EXEName, "Options", "TimeStamp/Channels", chkTimeStampChannels.value = vbChecked
    SaveSetting App.EXEName, "Options", "TimeStamp/Privates", chkTimeStampPrivates.value = vbChecked
    SaveSetting App.EXEName, "Options", "TimeStamp/Status", chkTimeStampStatus.value = vbChecked
    
    DB.X "Saving Misc2 and Proxy Settings"
    
    SaveSetting App.EXEName, "Options", "StartAWeb", chkStartPage.value = vbChecked
    SaveSetting App.EXEName, "Options", "NodeHome", txtStartPage.Text
    
    'Store proxy settings
    SaveSetting App.EXEName, "Options", "Proxy/Port", txtProxyPort.Text
    SaveSetting App.EXEName, "Options", "Proxy/Address", cboProxy.Text
    
    SaveSetting App.EXEName, "Options", "FocusJoined", chkFocusJoined.value = vbChecked
       
    DB.X "Saving Buddy Settings"
    
    'Here we save a "binary" sequence
    'each part stores a single value
    'we combine all the values and save
    'them in one registry key
    SaveSetting App.EXEName, "Options", "Buddys", _
        IIf(chkEnterMsg.value = vbChecked, "1", "0") & _
        IIf(chkEnterWin.value = vbChecked, "1", "0") & _
        IIf(chkLeaveMsg.value = vbChecked, "1", "0") & _
        IIf(chkLeaveWin.value = vbChecked, "1", "0")
    
    DB.X "Saving Manifest Settings"
    
    'if XP common controls aren't supposed to be available and there is a manifest file...
    If FS.FileExists(App.Path & "\node.exe.manifest") And chkXPCommonControls.value = Unchecked Then
        'if there's a no-manifest file
        If FS.FileExists(App.Path & "\node.exe.no-manifest") Then
            'delete the manifest file
            FS.DeleteFile App.Path & "\node.exe.manifest", True
        'if there's no no-manifest file
        Else
            'rename the manifest file to no-manifest
            Name App.Path & "\node.exe.manifest" As App.Path & "\node.exe.no-manifest"
        End If
    'if there's no manifest file there and the user wants to use XP Common Controls...
    ElseIf Not FS.FileExists(App.Path & "\node.exe.manifest") And chkXPCommonControls.value = Checked Then
        'if a no-manifest file exists...
        If FS.FileExists(App.Path & "\node.exe.no-manifest") Then
            'rename it to manifest
            Name App.Path & "\node.exe.no-manifest" As App.Path & "\node.exe.manifest"
        Else
            'bad luck
        End If
    End If
    
    DB.X "Saving Msg, Lang and Skin Settings"
    
    'save the quit message
    SaveSetting App.EXEName, "Options", "QuitMsg", txtQuit.Text
    SaveSetting App.EXEName, "Options", "QuitList", lblQuitMultiFile.Caption
    SaveSetting App.EXEName, "Options", "QuitMulti", optQuitMulti.value
    
    'save the selected language index and filename into the registry
    SaveSetting App.EXEName, "Options", "Language", icLanguage.SelectedItem.Index
    SaveSetting App.EXEName, "Options", "LanguageFile", icLanguage.SelectedItem.Tag

    'save away sys settings
    DB.X "Saving Away Sys Settings"
    
    SaveSetting App.EXEName, "Options", "Away", chkUseAway.value = vbChecked
    SaveSetting App.EXEName, "Options", "AwayMins", txtAwayMinutes.Text
    SaveSetting App.EXEName, "Options", "AwayNick", chkAwayChangeNick.value = vbChecked
    SaveSetting App.EXEName, "Options", "AwayNickStr", txtAwayNick.Text
    SaveSetting App.EXEName, "Options", "AwayPerform", chkAwayPerform.value = vbChecked
    SaveSetting App.EXEName, "Options", "AwayPerformStr", txtAwayPerform.Text
    
    SaveSetting App.EXEName, "Options", "AwayBackNick", chkAwayBackChangeNick.value = vbChecked
    SaveSetting App.EXEName, "Options", "AwayBackNickStr", txtAwayBackNick.Text
    SaveSetting App.EXEName, "Options", "AwayBackPerform", chkAwayBackPerform.value = vbChecked
    SaveSetting App.EXEName, "Options", "AwayBackPerformStr", txtAwayBackPerform.Text
    
    'save the filename of the skin being used
    SaveSetting App.EXEName, "Options", "Skin", Skins(cboSkin.ListIndex)
    
    DB.X "Saving Plugin Settings"
    
    'save the plugins (load on startup)
    If fgPlugins.rows = 2 Then
        If LenB(fgPlugins.TextMatrix(1, 0)) = 0 Then
            'no plugins available
            DB.X "There are no plugins installed!"
        Else
            GoTo SavePluginsSettings
        End If
    Else
SavePluginsSettings:
        For i = 1 To fgPlugins.rows - 1
            SaveSetting App.EXEName, "Plugins", fgPlugins.TextMatrix(i, 0), Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(i, 0) & ".dll")).boolLoadOnStartup
        Next i
    End If
    
    DB.X "Saving SoundScheme/SmileyPack"
    
    SaveSetting "Node", "Options", "SoundScheme", icSoundSchemes.Text
    SaveSetting "Node", "Options", "SmileyPack", icSmileyPacks.Text
   
    DB.Leave "SaveAll()"
End Sub
Public Function LoadAll()
    'load all settings
    Dim LogChannels As String 'hold logchannel boolean value
    Dim LogPrivates As String 'hold logprivates boolean value
    Dim i As Integer ' counter for looping through arrays
    Dim intFL As Integer 'free file index
    Dim strTempLine As String 'a string to hold the currently inputed line from a file
    Dim LanguageFile As File 'the file object which stores the current language file
    Dim strLanguageID As String 'the language name
    Dim strFlagFile As String 'the file storing the flag for the current language
    Dim strBuddys As String 'a temporary string variable used to store the buddys keys imported from the registry
    Dim intTemp As Integer
    
    DB.Enter "LoadAll()"
    
    DB.X "Loading info.dat"
    If Not FS.FileExists(App.Path & "\conf\info.dat") Then
        intFL = FreeFile
        DB.X "info.dat does not exist"
        DB.X "Creating info.dat"
        Open App.Path & "\conf\info.dat" For Output As intFL
        Print #intFL, "NodeUser"
        Print #intFL, "_NodeUser_"
        Print #intFL, "NodeUsr"
        Print #intFL, "node@sourceforge.net"
        Print #intFL, "node.sourceforge.net"
        Close intFL
    End If
    'get a free file index
    intFL = FreeFile
    'open the info file to get the user's information
    Open App.Path & "\conf\info.dat" For Input As #intFL
        'input five sequential line from the file
        'each line represents a property: nick, alt, alt2, email and real name.
        txtNickname.Text = xLineInput(intFL)
        txtAlt.Text = xLineInput(intFL)
        txtAltTwo.Text = xLineInput(intFL)
        txtEmail.Text = xLineInput(intFL)
        txtReal.Text = xLineInput(intFL)
    'close the file
    Close #intFL
        
    DB.X "Loading ignore_nick.dat"
    If Not FS.FileExists(App.Path & "\conf\ignore_nick.dat") Then
        intFL = FreeFile
        DB.X "ignore_nick.dat does not exist"
        DB.X "Creating ignore_nick.dat"
        Open App.Path & "\conf\ignore_nick.dat" For Output As intFL
        Close intFL
    End If
    
    'clear the ignore list in order to fill it again
    lstIgnore(0).Clear
    'get a free file
    intFL = FreeFile
    'open the ignore list file
    Open App.Path & "\conf\ignore_nick.dat" For Binary As #intFL
        'get all the lines from this file; fill the list with these items
        Do Until EOF(intFL)
            'read a line from the file and save it into strTempLine
            'create an error trap in the case an unexpected end-of-file occurs
            On Error GoTo Ignore_EOF
            Line Input #intFL, strTempLine
            'if it contains something, add the item to the ignore list
            If Len(strTempLine) > 0 Then
                lstIgnore(0).AddItem strTempLine
            End If
        Loop
Ignore_EOF:
        'close the file
    Close #intFL
    
    'TO DO: Load IP ignore list
    
    DB.X "Loading Logging Settings"
    'load values of logprivates and logchannels from reg
    'the value set is Checked if the stored value is True, or else it's Unchecked
    'store the option values at the variables frmMain.Logging_Channels and frmMain.Logging_Privates
    chkLogChannels.value = xLet(Options.LogChannels, IIf(CBool(GetSetting(App.EXEName, "Options", "LogChannels", True)), vbChecked, vbUnchecked))
    chkLogPrivates.value = xLet(Options.LogPrivates, IIf(CBool(GetSetting(App.EXEName, "Options", "LogPrivates", False)), vbChecked, vbUnchecked))
    chkLogByNet.value = xLet(Options.LogByNetwork, IIf(CBool(GetSetting(App.EXEName, "Options", "LogByNet", False)), vbChecked, vbUnchecked))
    chkLogRAW.value = xLet(Options.LogRAW, IIf(CBool(GetSetting(App.EXEName, "Options", "Debug\Raw", True)), vbChecked, vbUnchecked))
    chkTimeStampsLog.value = xLet(Options.TimeStampLogs, IIf(CBool(GetSetting(App.EXEName, "Options", "TimeStamp/Log", True)), vbChecked, vbUnchecked))

    DB.X "Loading Misc Settings"
    'use the same method to load some more settings
    chkLatest.value = xLet(Options.CheckLatest, IIf(CBool(GetSetting(App.EXEName, "Options", "CheckLatest", True)), vbChecked, vbUnchecked))
    chkFade.value = xLet(Options.FadeTransaction, IIf(CBool(GetSetting(App.EXEName, "Options", "Fade", False)), vbChecked, vbUnchecked))
    chkScripting.value = xLet(Options.EnableScripting, IIf(CBool(GetSetting(App.EXEName, "Options", "Scripting", True)), vbChecked, vbUnchecked))
    chkCodeBehind.value = xLet(Options.EnableCodeBehind, IIf(CBool(GetSetting(App.EXEName, "Options", "CodeBehind", True)), vbChecked, vbUnchecked))
    chkLoading.value = xLet(Options.HTMLLoading, IIf(CBool(GetSetting(App.EXEName, "Options", "HTMLLoading", True)), vbChecked, vbUnchecked))
    chkHTMLError.value = xLet(Options.HTMLError, IIf(CBool(GetSetting(App.EXEName, "Options", "HTMLError", True)), vbChecked, vbUnchecked))
    chkAutocomplete.value = xLet(Options.AutoComplete, IIf(CBool(GetSetting(App.EXEName, "Options", "AutoComplete", True)), vbChecked, vbUnchecked))
    chkTOD.value = xLet(Options.TOD, IIf(CBool(GetSetting(App.EXEName, "Options", "TOD", True)), vbChecked, vbUnchecked))
    chkInfoTips.value = xLet(Options.InfoTips, IIf(CBool(GetSetting(App.EXEName, "Options", "InfoTips", True)), vbChecked, vbUnchecked))
    chkTray.value = xLet(Options.KeepTrayRunning, IIf(CBool(GetSetting(App.EXEName, "Options", "Tray", True)), vbChecked, vbUnchecked))
    chkRestoreStatus.value = xLet(Options.RestoreStatus, IIf(CBool(GetSetting(App.EXEName, "Options", "RestoreStatus", True)), vbChecked, vbUnchecked))
    chkAutoNDC.value = xLet(Options.AutoNDC, IIf(CBool(GetSetting(App.EXEName, "Options", "AutoNDC", False)), vbChecked, vbUnchecked))
    TxtRangeL.Text = GetSetting(App.EXEName, "Options", "PortRangeLow", "6000")
    TxtRangeH.Text = GetSetting(App.EXEName, "Options", "PortRangeHigh", "7000")
    chkJoinPanel.value = xLet(Options.JoinPanel, IIf(CBool(GetSetting(App.EXEName, "Options", "JoinPanel", True)), vbChecked, vbUnchecked))
    chkProxyEnable.value = IIf(xLet(Options.UseProxy, CBool(GetSetting(App.EXEName, "Options", "Proxy/UseProxy", False))), vbChecked, vbUnchecked)
    chkProxyEnable_Click
    
    txtRetryDelay.Text = xLet(Options.ConnectRetryDelay, GetSetting(App.EXEName, "Options", "ConnectRetryDelay", 5))
    chkRetry.value = IIf(xLet(Options.ConnectRetry, CBool(GetSetting(App.EXEName, "Options", "ConnectRetry", False))), vbChecked, vbUnchecked)
    chkRetry_Click
        
    chkJoinOnInvite.value = IIf(xLet(Options.JoinOnKick, CBool(GetSetting(App.EXEName, "Options", "JoinOnInvite", True))), vbChecked, vbUnchecked)
    chkJoinOnKick.value = IIf(xLet(Options.JoinOnInvite, CBool(GetSetting(App.EXEName, "Options", "RejoinOnKick", False))), vbChecked, vbUnchecked)
    chkModesOnJoin.value = IIf(xLet(Options.ModesOnJoin, CBool(GetSetting(App.EXEName, "Options", "ModesOnJoin", True))), vbChecked, vbUnchecked)
    chkKeepChannelsOpen.value = IIf(xLet(Options.KeepChannelsOpen, CBool(GetSetting(App.EXEName, "Options", "KeepChannelsOpen", False))), vbChecked, vbUnchecked)
    
    'perform settings
    DB.X "Loading Perform"
    
    chkPerform.value = IIf(xLet(Options.EnablePerform, CBool(GetSetting(App.EXEName, "Options", "EnablePerform", False))), vbChecked, vbUnchecked)
    Options.PerformSingle = GetSetting(App.EXEName, "Options", "PerformSingle", True)
    
    LoadPerformServers
    
    'after setting opts values, the appropriate
    'events will be fired that will then load
    'the correct perform text inside the RTB
    'control
    If Options.PerformSingle Then
        optPerformSingle.value = True
        optPerformMultiple.value = False
    Else
        optPerformMultiple.value = True
        optPerformSingle.value = False
    End If

    'Browsing Settings
    DB.X "Loading Browsing Settings"
    Options.BrowseInternalBrowser = GetSetting(App.EXEName, "Options", "BrowseInternal", True)
    If Options.BrowseInternalBrowser Then
        optBrowseInternal.value = True
        optBrowseDefaultBrowser.value = False
    Else
        optBrowseInternal.value = False
        optBrowseDefaultBrowser.value = True
    End If
    
    chkBrowseParseLinks.value = IIf(xLet(Options.BrowseParseLinks, CBool(GetSetting(App.EXEName, "Options", "BrowseParseLinks", True))), vbChecked, vbUnchecked)
    
    'Display Settings
    DB.X "Loading Display Settings"
    Options.DisplayNormal = GetSetting(App.EXEName, "Options", "DisplayNormal", "&lt; %nick &gt;")
    txtDisplayNormal.Text = Options.DisplayNormal
    Options.DisplaySkinFontSize = GetSetting(App.EXEName, "Options", "DisplaySkinFontSize")
    txtDisplaySkinFontSize.Text = Options.DisplaySkinFontSize
    If LenB(Options.DisplaySkinFontSize) > 0 Then
        SetDisplaySkinFontSize
    End If
    
    'CTCP Settings
    DB.X "Loading CTCP Settings"
    chkCTCPPing.value = xLet(Options.CTCPPing, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Ping/Reply", True)), vbChecked, vbUnchecked))
    chkCTCPPingIgnore.value = xLet(Options.CTCPPingToIgnored, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Ping/Ignore", True)), vbChecked, vbUnchecked))
    chkCTCPPing_Click
    
    chkCTCPVersion.value = xLet(Options.CTCPVersion, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Version/Reply", True)), vbChecked, vbUnchecked))
    chkCTCPVersionIgnore.value = xLet(Options.CTCPVersionToIgnored, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Version/Ignore", False)), vbChecked, vbUnchecked))
    chkCTCPVersion_Click
        
    chkCTCPTime.value = xLet(Options.CTCPTime, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Time/Reply", True)), vbChecked, vbUnchecked))
    chkCTCPTimeIgnore.value = xLet(Options.CTCPTimeToIgnored, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Time/Ignore", False)), vbChecked, vbUnchecked))
    chkCTCPTime_Click
    
    chkCTCPFlood.value = xLet(Options.CTCPFloodProtect, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/FloodProtection", True)), vbChecked, vbUnchecked))
    chkCTCPFloodBounce.value = xLet(Options.CTCPFloodBounce, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/BounceToFlooders", False)), vbChecked, vbUnchecked))
    chkCTCPFlood_Click
    
    chkCTCPVersionCustom.value = xLet(Options.CTCPVersionCustomize, IIf(CBool(GetSetting(App.EXEName, "Options", "CTCP/Version/Custom", False)), vbChecked, vbUnchecked))
    
    txtCTCPVersionCustom.Text = GetSetting(App.EXEName, "Options", "CTCP/Version/Message", "I'm using Node " & VERSION_CODENAME & " " & App.Major & "." & App.Minor)
    If Options.CTCPVersionCustomize Then
        Options.CTCPVersionMessage = txtCTCPVersionCustom.Text
    Else
        Options.CTCPVersionMessage = "I'm using Node " & VERSION_CODENAME & " " & App.Major & "." & App.Minor
    End If
    chkCTCPVersion_Click
    
    DB.X "Loading Selection Copy/Clear Settings"
    chkSelectionCopy.value = xLet(Options.SelectionCopy, IIf(CBool(GetSetting(App.EXEName, "Options", "SelectCopy", True)), vbChecked, vbUnchecked))
    chkSelectionClear.value = xLet(Options.SelectionClear, IIf(CBool(GetSetting(App.EXEName, "Options", "SelectClear", True)), vbChecked, vbUnchecked))
    chkSelectionCopy_Click
    
    'Accessibility
    DB.X "Loading Accessibility Settings"
    chkNarration.value = xLet(Options.Narration, IIf(CBool(GetSetting(App.EXEName, "Options\Accessibility\Narration", "Enabled", False)), vbChecked, vbUnchecked))
    chkNarrationInterface.value = xLet(Options.NarrationInterface, IIf(CBool(GetSetting(App.EXEName, "Options\Accessibility\Narration", "Interface", False)), vbChecked, vbUnchecked))
    chkNarrationChannels.value = xLet(Options.NarrationChannels, IIf(CBool(GetSetting(App.EXEName, "Options\Accessibility\Narration", "Channels", False)), vbChecked, vbUnchecked))
    chkNarrationPrivates.value = xLet(Options.NarrationPrivates, IIf(CBool(GetSetting(App.EXEName, "Options\Accessibility\Narration", "Privates", False)), vbChecked, vbUnchecked))
    chkNarrationStatus.value = xLet(Options.NarrationStatus, IIf(CBool(GetSetting(App.EXEName, "Options\Accessibility\Narration", "Status", False)), vbChecked, vbUnchecked))
    chkNarration_Click
    
    'TimeStamp
    DB.X "Loading TimeStamp Settings"
    chkTimeStamp.value = xLet(Options.TimeStamp, IIf(CBool(GetSetting(App.EXEName, "Options", "TimeStamp/Enable", False)), vbChecked, vbUnchecked))
    chkTimeStampChannels.value = xLet(Options.TimeStampChannels, IIf(CBool(GetSetting(App.EXEName, "Options", "TimeStamp/Channels", True)), vbChecked, vbUnchecked))
    chkTimeStampPrivates.value = xLet(Options.TimeStampPrivates, IIf(CBool(GetSetting(App.EXEName, "Options", "TimeStamp/Privates", True)), vbChecked, vbUnchecked))
    chkTimeStampStatus.value = xLet(Options.TimeStampStatus, IIf(CBool(GetSetting(App.EXEName, "Options", "TimeStamp/Status", True)), vbChecked, vbUnchecked))
    'Careful: Options.TimeStamp may be Empty so do not use Not
    If Options.TimeStamp = False Then
        Options.TimeStampChannels = False
        Options.TimeStampPrivates = False
        Options.TimeStampStatus = False
    End If
    chkTimeStamp_Click
    
    DB.X "Loading Misc2 Settings"
    chkStartPage.value = xLet(Options.StartPage, IIf(CBool(GetSetting(App.EXEName, "Options", "StartAWeb", True)), vbChecked, vbUnchecked))
    txtStartPage.Text = xLet(Options.StartPageURL, GetSetting(App.EXEName, "Options", "NodeHome", "http://node.sourceforge.net/"))
    
    chkFocusJoined.value = xLet(Options.FocusJoined, IIf(CBool(GetSetting(App.EXEName, "Options", "FocusJoined", False)), vbChecked, vbUnchecked))
    
    chkNickLinkChan.value = xLet(Options.NickLinkChan, IIf(CBool(GetSetting(App.EXEName, "Options", "NickLinkChan", True)), vbChecked, vbUnchecked))
    chkNickLinkPriv.value = xLet(Options.NickLinkPriv, IIf(CBool(GetSetting(App.EXEName, "Options", "NickLinkPriv", False)), vbChecked, vbUnchecked))
    chkNickLinkMineChan.value = xLet(Options.NickLinkMineChan, IIf(CBool(GetSetting(App.EXEName, "Options", "NickLinkMineChan", False)), vbChecked, vbUnchecked))
    chkNickLinkMinePriv.value = xLet(Options.NickLinkMinePriv, IIf(CBool(GetSetting(App.EXEName, "Options", "NickLinkMinePriv", False)), vbChecked, vbUnchecked))
    
    'Load Proxy Settings
    DB.X "Loading Proxy Settings"
    txtProxyPort.Text = xLet(Options.ProxyPort, GetSetting(App.EXEName, "Options", "Proxy/Port", "1080"))
    cboProxy.Text = xLet(Options.ProxyIP, GetSetting(App.EXEName, "Options", "Proxy/Address", vbNullString))
    
    'first we set Options.DCCAntivirus
    'and then we assign vbChecked or vbUnchecked to chkAntivirus.Value
    'which will cause chkAntivirus_Click to fire
    'If it's vbUnchecked it will just set Options.DCCAntivirus to ""
    '(which is already done anyway)
    'Else, if it's vbChecked, Options.DCCAntivirus will be
    'already set, so it will avoid taking any actions
    DB.X "Loading Antivirus Setting"
    chkAntivirus.value = IIf(xLet(Options.DCCAntivirus, GetSetting(App.EXEName, "Options", "AntiVirus", vbNullString)) = vbNullString, vbUnchecked, vbChecked)
    
    'note: iif function is used differently in the particular situation bellow
    '(as xLet returns the result of the assignement which can be a different type than the value)
    DB.X "Loading Manifest Setting"
    chkXPCommonControls.value = IIf(xLet(Options.XPCommonControls, FS.FileExists(App.Path & "\" & App.EXEName & ".exe.manifest")), vbChecked, vbUnchecked)
    
    'load the buddy's setting
    'get the setting
    DB.X "Loading Buddy Settings"
    strBuddys = GetSetting(App.EXEName, "Options", "Buddys", "1111")
    'depending on the value of each digit change the setting
    chkEnterMsg.value = IIf(xLet(Options.BuddyEnterMSG, Strings.Mid$(strBuddys, 1, 1)), vbChecked, vbUnchecked)
    chkEnterWin.value = IIf(xLet(Options.BuddyEnterWIN, Strings.Mid$(strBuddys, 2, 1)), vbChecked, vbUnchecked)
    chkLeaveMsg.value = IIf(xLet(Options.BuddyLeaveMSG, Strings.Mid$(strBuddys, 3, 1)), vbChecked, vbUnchecked)
    chkLeaveWin.value = IIf(xLet(Options.BuddyLeaveWIN, Strings.Mid$(strBuddys, 4, 1)), vbChecked, vbUnchecked)
    
    Options.SessionC = GetSetting(App.EXEName, "Options", "SessionCrash", 2)
    optSessionCR.value = Options.SessionC = 0
    optSessionCD.value = Options.SessionC = 1
    optSessionCA.value = Options.SessionC = 2
    
    Options.SessionN = GetSetting(App.EXEName, "Options", "SessionNormal", 1)
    optSessionNR.value = Options.SessionN = 0
    optSessionND.value = Options.SessionN = 1
    optSessionNA.value = Options.SessionN = 2
    
    Options.DCCOptionsB = GetSetting(App.EXEName, "Options", "DCC-Buddies", 0)
    optDCCB(0).value = Options.DCCOptionsB = 0
    optDCCB(1).value = Options.DCCOptionsB = 1
    optDCCB(2).value = Options.DCCOptionsB = 2
    
    Options.DCCOptionsI = GetSetting(App.EXEName, "Options", "DCC-Ignore", 1)
    optDCCI(0).value = Options.DCCOptionsI = 0
    optDCCI(1).value = Options.DCCOptionsI = 1
    optDCCI(2).value = Options.DCCOptionsI = 2
    
    Options.DCCOptionsE = GetSetting(App.EXEName, "Options", "DCC-Everyone", 2)
    optDCCE(0).value = Options.DCCOptionsE = 0
    optDCCE(1).value = Options.DCCOptionsE = 1
    optDCCE(2).value = Options.DCCOptionsE = 2
    
    'buddys code by jnfoot
    Dim buddyname As String 'the name of the buddy
    Dim recnum As Integer 'the index of the buddy
    
    'clear the buddy and the Welcome Text lists
    lstBdyNk.Clear
    lstBdyWT.Clear
    'we start from buddy index zero
    recnum = 0
    If Not FS.FileExists(App.Path & "\conf\buddy.info") Then
        intFL = FreeFile
        DB.X "buddy.info does not exist"
        DB.X "Creating buddy.info"
        Open App.Path & "\conf\buddy.info" For Output As intFL
        Close intFL
    End If
    DB.X "Getting FreeFile to open buddy.info"
    'get a free file index
    intFL = FreeFile
    DB.X "Loading buddy.info"
    'open the buddys configuration file
    Open App.Path & "\conf\buddy.info" For Input As #intFL 'open the buddy list file
        DB.X "Opened File"
        'work until no more buddies
        Do Until EOF(intFL)
            'go to the next buddy's index
            recnum = recnum + 1
            'DB.X "Reading Record #" & recnum
            'get the buddy name
            Input #intFL, buddyname
            'DB.X "Record Text: " & buddyname
            'if its odd then
            If (recnum / 2) <> Int(recnum / 2) Then
                'add it to the names list
                lstBdyNk.AddItem buddyname
            'if its not odd then
            Else
                buddyname = Replace(buddyname, ChrW$(1), ",")
                'add it to the text list
                lstBdyWT.AddItem buddyname
            End If
        'go to the next line
        Loop
    'unload file
    Close #intFL
    DB.X "buddy.info loaded OK"
    
    If lstBdyNk.ListCount > 0 Then
        For i = 0 To lstBdyNk.ListCount - 1
            ReDim Preserve AllBuddies(i)
            Set AllBuddies(i) = New clsIdentity
            AllBuddies(i).Name = lstBdyNk.List(i)
            AllBuddies(i).isOnline = False
        Next i
    End If
    
    DB.X "Closed buddy.info"
    
    DB.X "Reading StartupConnect Setting..."
    chkStartupConnect.value = IIf(xLet(Options.StartupConnect, CBool(GetSetting(App.EXEName, "Options", "StartupConnect", False))), vbChecked, vbUnchecked)
    
    DB.X "Getting HostName for StartupConnect..."
    Options.StartupConnectHostname = GetSetting(App.EXEName, "Options", "StartupConnectServer", vbNullString)
    
    DB.X "Startup Connect HostName = " & Options.StartupConnectHostname
    If LenB(Options.StartupConnectHostname) > 0 Then
        DB.X "Startup Connect HostName exists"
        DB.X "Attemping to GetServerListItemIndexFromHostname()..."
        intTemp = GetServerListItemIndexFromHostname(Options.StartupConnectHostname)
        If intTemp > -1 Then
            cboConnectServer.ListIndex = intTemp
        End If
        On Error GoTo 0
        DB.X "Getting StartupConnect Port..."
        Options.StartupConnectPort = GetSetting(App.EXEName, "Options", "StartupConnectPort", "0")
    Else
        DB.X "Startup Connect HostName does not exist"
    End If
    
    chkStartupConnect_Click
    
    DB.X "Loading Languages List"
    icLanguage.ComboItems.Clear
    'load languages
    'go through the files inside \data\languages directory
    'if a file ends with .lang add it to the image combo.
    For Each LanguageFile In FS.GetFolder(App.Path & "\data\languages").Files
        If Strings.Right$(Strings.LCase$(LanguageFile.Name), Len(".lang")) = ".lang" Then
            'this is a language file; add it to the list
            'get a free file once again
            intFL = FreeFile
            'open the language file
            Open LanguageFile.Path For Input As #intFL
            'input the first two lines from the file
            'the first one is the language id(the name of the language)
            'the second is the filename of the flag file
            'relative to the directory \data\languages
            'create a new combo item with caption = LanguageID(language name)
            'set its tag to the actualy language file.
            icLanguage.ComboItems.Add(, , Mid$(xLineInput(intFL), 1)).Tag = LanguageFile.Path
            'create a new item in the image list containing the appropriate flag or other language symbol.
            ilLanguage.ListImages.Add , , LoadPicture(App.Path & "\data\languages\" & xLineInput(intFL))
            Close #intFL
        End If
        'go to the next language file
    Next LanguageFile
    'initialise the image list for the combo
    Set icLanguage.ImageList = ilLanguage
    'go through all languages
    For i = 1 To icLanguage.ComboItems.Count
        'set the language to the right flag item.
        icLanguage.ComboItems.Item(i).Image = i
        'if this is the langauge that should be selected...
        If Strings.LCase$(icLanguage.ComboItems.Item(i).Tag) = Strings.LCase$(GetSetting(App.EXEName, "Options", "LanguageFile", App.Path & "\data\languages\english.lang")) Then
            '...select it
            Set icLanguage.SelectedItem = icLanguage.ComboItems.Item(i)
        End If
    'go to the next language item
    Next i
    If icLanguage.SelectedItem Is Nothing Then
        Set icLanguage.SelectedItem = icLanguage.ComboItems.Item(1)
    End If
    
    'load server list
    'ReadServers
        
    DB.X "Loading Skins List"
    'load the skins list
    ReadSkins
        
    DB.X "Loading Smiley Pack"
    Options.UseSmileys = GetSetting("Node", "Options", "Smileys", True)
    chkEnableSmileys.value = IIf(Options.UseSmileys, vbChecked, vbUnchecked)
    chkEnableSmileys_Click
    
    If Options.UseSmileys Then
        Options.SmileyPack = GetSetting("Node", "Options", "SmileyPack", "phpbb")
        LoadSmileyPack ThisSmileyPack, App.Path & "/data/smileys/" & Options.SmileyPack & "/" & Options.SmileyPack & ".xml"
    End If
    
    
    'load the quit message from the registry; store it into Options.QuitMsg
    txtQuit.Text = xLet(Options.QuitMsg, GetSetting(App.EXEName, "Options", "QuitMsg", "I am using Node, http://node.sourceforge.net"))
    lblQuitMultiFile.Caption = xLet(Options.QuitFile, GetSetting(App.EXEName, "Options", "QuitList", App.Path & "\data\quit.lst"))
    optQuitMulti.value = xLet(Options.QuitMultiple, CBool(GetSetting(App.EXEName, "Options", "QuitMulti", False)))
    optQuitSingle.value = Not optQuitMulti.value
    
    'load the language file being used from the registry
    Options.LanguageFile = GetSetting(App.EXEName, "Options", "LanguageFile", App.Path & "\data\languages\english.lang")
    
    DB.X "Loading Plugins List"
    'build the plugins list
    ListPlugIns

    DB.X "Loading SoundSchemes List"
    
    'get the selected soundscheme
    Options.SoundScheme = GetSetting("Node", "Options", "SoundScheme", "blackalien")
    
    'build the sound schemes list
    ListSoundSchemes
    
    'build the smiley packs list
    ListSmileyPacks
    
    DB.X "Loading SoundScheme `" & Options.SoundScheme & "'"
    
    ThisSoundScheme = LoadSoundScheme(App.Path & "/data/sounds/" & Options.SoundScheme & "/" & Options.SoundScheme & ".xml")
    
    DB.X "Loading HotKeys"
    'read from the hot keys list
    LoadHotKeys

    'load away sys settings
    DB.X "Loading Away Sys Settings..."
    chkUseAway.value = IIf(xLet(Options.AwayEnabled, CBool(GetSetting(App.EXEName, "Options", "Away", False))), vbChecked, vbUnchecked)
    txtAwayMinutes.Text = xLet(Options.AwayMinutes, GetSetting(App.EXEName, "Options", "AwayMins", 10))
    chkUseAway_Click
    
    chkAwayChangeNick.value = IIf(xLet(Options.AwayNick, CBool(GetSetting(App.EXEName, "Options", "AwayNick", True))), vbChecked, vbUnchecked)
    txtAwayNick.Text = xLet(Options.AwayNickStr, GetSetting(App.EXEName, "Options", "AwayNickStr", "NodeUser_Away"))
    chkAwayPerform.value = IIf(xLet(Options.AwayPerform, CBool(GetSetting(App.EXEName, "Options", "AwayPerform", False))), vbChecked, vbUnchecked)
    txtAwayPerform.Text = xLet(Options.AwayPerformStr, GetSetting(App.EXEName, "Options", "AwayPerformStr", "/ame is now away"))
    
    chkAwayBackChangeNick.value = IIf(xLet(Options.AwayBackNick, CBool(GetSetting(App.EXEName, "Options", "AwayBackNick", True))), vbChecked, vbUnchecked)
    txtAwayBackNick.Text = xLet(Options.AwayBackNickStr, GetSetting(App.EXEName, "Options", "AwayBackNickStr", "NodeUser"))
    chkAwayBackPerform.value = IIf(xLet(Options.AwayBackPerform, CBool(GetSetting(App.EXEName, "Options", "AwayBackPerform", False))), vbChecked, vbUnchecked)
    txtAwayBackPerform.Text = xLet(Options.AwayBackPerformStr, GetSetting(App.EXEName, "Options", "AwayBackPerformStr", "/ame is now away"))

    If GetSetting(App.EXEName, "Options", "ParseMemoServ", False) Then
        DB.X "EXPERIMENTAL FEATURE: MemoServ Messages Parsing is ENABLED"
        Options.ParseMemoServ = True
    Else
        DB.X "MemoServ parsing is disabled"
        Options.ParseMemoServ = False
    End If
    
    DB.Leave "LoadAll()"
End Function
Private Sub lvList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewItem As ListItem
    If Button = 1 Then
        'display and reenable the list
        'so that items cannot be moved
        lvList.Enabled = False
        lvList.Enabled = True
        Set NewItem = lvList.HitTest(X, Y)
        If Not NewItem Is Nothing Then
            'select the clicked item
            Set lvList.SelectedItem = NewItem
            'give the focus to the list so
            'that the selected item is marked
            lvList.SetFocus
            'display the appropriate settings
            'category on the right
            'lvList_Click
        End If
    End If
End Sub
Private Sub mnuLoadPlugin_Click()
    If LoadPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll") Then
        fgPlugins.TextMatrix(fgPlugins.Row, 1) = Language(280) 'Loaded: Yes
    Else
        'failed
        'display warning
        MsgBox Replace(Language(528), "%1", Err.Description), vbCritical, Language(529)
    End If
End Sub
Private Sub mnuPlugInOptions_Click()
    If Not Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll")).objPlugIn Is Nothing Then
        Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll")).objPlugIn.PluginOptions
    Else
        MsgBox Language(572)
    End If
End Sub
Private Sub mnuPluginStartup_Click()
    Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll")).boolLoadOnStartup = Not Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll")).boolLoadOnStartup
    mnuPluginStartup.Checked = Plugins(NumToPlugIn("prjPlugIn" & fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll")).boolLoadOnStartup
End Sub
Private Sub mnuUnloadPlugin_Click()
    If UnloadPlugIn(fgPlugins.TextMatrix(fgPlugins.Row, 0) & ".dll") Then
        fgPlugins.TextMatrix(fgPlugins.Row, 1) = vbNullString 'Loaded: No
    End If
End Sub
Private Sub tvOptions_Collapse(ByVal Node As ComctlLib.Node)
    If Node.Key = "root" Then
        'the root node can not be collapsed
        Node.Expanded = True
    End If
End Sub
Public Sub tvOptions_NodeClick(ByVal Node As ComctlLib.Node)
    'The options category changed; display the appropriate frame
    Dim OptionCategory As Frame 'object variable used to store the current frame for the loop
    Dim ItemToSelect As Integer 'the index of the option category that should be selected; -1 if nothing is selected
    
    If (Left$(Node.Key, 1) = "s" Or Node.Key = "root") And Not Node.Child Is Nothing Then
        'this node is a parent of somethin else
        'display description html
        wbNav.Navigate2 App.Path & "/data/html/options.html"
        wbNav.Visible = True
        ItemToSelect = -1
    Else
        'find out which category is selected; store it into ItemToSelect variable
        ItemToSelect = Right$(Node.Key, Len(Node.Key) - 1)
        wbNav.Visible = False
    End If
    
    'go through all frames
    For Each OptionCategory In fraOptions
        'this frame should not be shown
        If OptionCategory.Index <> ItemToSelect Then
            'if it is visible
            If OptionCategory.Visible Then
                'hide it
                OptionCategory.Visible = False
            End If
        Else
            'this is the frame that should be shown(the selected category)
            'if it's not already visible
            If Not OptionCategory.Visible Then
                'show it
                OptionCategory.Visible = True
            End If
        End If
        'go to the next frame
    Next OptionCategory
End Sub
Private Sub txtAwayMinutes_Change()
    If Not IsNumeric(txtAwayMinutes.Text) Then
        txtAwayMinutes.Text = Val(txtAwayMinutes.Text)
    End If
End Sub

Private Sub txtProxyPort_LostFocus()
    txtProxyPort.Text = Int(Val(txtProxyPort.Text))
    If txtProxyPort.Text = 0 Or txtProxyPort.Text < 0 Then
        txtProxyPort.Text = 1080 'default port
    ElseIf txtProxyPort.Text > 65535 Then
        txtProxyPort.Text = 65535 'maximum value
    End If
End Sub
Private Sub wbNav_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim webdocNav As HTMLDocument
    Dim strTitle As String
    Dim strText As String

    Set webdocNav = wbNav.Document
    Select Case tvOptions.SelectedItem.Key
        Case "root"
            '@root = options
            strTitle = Language(4)
            strText = Language(605) & "<br>" & Language(606) & "<br>" & Language(607)
        Case "s01"
            '@root/s01 = irc
            strTitle = Language(608)
            strText = Language(609)
        Case "s02"
            '@root/s02 = personalize
            strTitle = Language(610)
            strText = Language(611)
        Case "s03"
            '@root/s03 = connections
            strTitle = Language(612)
            strText = Language(613)
        Case "s04"
            '@root/s04 = other options
            strTitle = Language(614)
            strText = Language(615)
        Case "s05"
            '@root/s05 = security
            strTitle = Language(622)
            strText = Language(898)
    End Select
    
    xNodeTag webdocNav, "lang_options_title", strTitle
    xNodeTag webdocNav, "lang_options", strText
End Sub
Private Sub FixPositions()
    'sub to fix the positions of the controls
    'on the form and the size of the form itself
    Dim OptionCategory As Frame
    'go through all frames and set their left, top, width and height
    'the same as the frame number 0. Make them invisible
    For Each OptionCategory In fraOptions
        OptionCategory.Left = fraOptions(0).Left
        OptionCategory.Top = fraOptions(0).Top
        OptionCategory.Width = fraOptions(0).Width
        OptionCategory.Height = fraOptions(0).Height
        OptionCategory.Visible = False
        OptionCategory.BorderStyle = vbBSNone
        'some items are not in the categories yet because the are under construction
        On Error Resume Next
        OptionCategory.Caption = lvList.ListItems.Item(OptionCategory.Index + 1)
    Next OptionCategory
    'set the height and width of the form to two constant values
    Me.Height = OPTIONS_HEIGHT
    Me.Width = OPTIONS_WIDTH
    
    fgHotKeys.ColWidth(0) = fgHotKeys.Width * (1 / 5)
    fgHotKeys.ColWidth(1) = fgHotKeys.Width * (4 / 5)

    fgPlugins.ColWidth(0) = fgPlugins.Width * (4 / 5)
    fgPlugins.ColWidth(1) = fgPlugins.Width * (1 / 5)

    wbNav.Left = tvOptions.Left + tvOptions.Width - 30
    wbNav.Top = -30
    
    wbNav.Width = Me.ScaleWidth - wbNav.Left + 60
    wbNav.Height = Me.ScaleHeight - wbNav.Top + 60
    
    tvOptions.Top = -30
    tvOptions.Height = Me.ScaleHeight
    tvOptions.ZOrder 0
End Sub
Private Sub ReadSkins()
    'routine used to fill the skins combo box
    Dim SkinFile As File 'store the current skin file as we loop through all the files in the directory
    Dim intFL As Integer 'file index variable
    Dim strSelected As String 'the filename of the selected skin
    Dim i As Integer 'a counter variable for the loop
    
    'clear the skins combo
    cboSkin.Clear
    'clear the skins filenames array
    ReDim Skins(0)
    
    'go through the files in the skins folder
    For Each SkinFile In FS.GetFolder(App.Path & "\data\skins").Files
        'if it's a skin file...
        If Strings.Right$(Strings.LCase$(SkinFile.Name), Len(".skin")) = ".skin" Then
            'this is a skin file; add it to the list
            'get a free file index
            intFL = FreeFile
            'open the skin file
            Open SkinFile.Path For Input As #intFL
            'input the line from the file
            'it is the name of the skin
            'create a new item in the combo box
            cboSkin.AddItem xLineInput(intFL)
            'close the file
            Close #intFL
            'store its filename in the filenames array
            Skins(UBound(Skins)) = SkinFile.Path
            'increase the size of the filenames array for the next skins
            ReDim Preserve Skins(UBound(Skins) + 1)
        End If
        'go to the next language file
    Next SkinFile
    
    'get the selected skin filename from the registry
    strSelected = GetSetting(App.EXEName, "Options", "Skin", App.Path & "\data\skins\default.skin")
    
    'go through the skins to find out which one is selected
    For i = 0 To cboSkin.ListCount - 1
        'if this is the one that should be selected...
        If Strings.LCase$(Skins(i)) = Strings.LCase$(strSelected) Then
            '...select it!
            cboSkin.ListIndex = i
        End If
    'go to the next skin
    Next i
    'if the setting was incorrect and no skin is selected...
    If cboSkin.ListIndex < 0 Then
        '...select the first skin in the list
        cboSkin.ListIndex = 0
    End If
End Sub
Private Sub ListPlugIns()
    Dim i As Integer
    Dim plugInFile As File
    'Lists all installed plugins
    
    'clear the plugins list
    For i = fgPlugins.rows - 1 To 2 Step -1
        fgPlugins.RemoveItem 1
    Next i
    
    'add all the valid plugins
    For i = 0 To NumToPlugIn.Count - 1
        'add to list(do not display the extension)
        fgPlugins.AddItem Right$(Plugins(i).strName, Len(Plugins(i).strName) - Len("prjPlugIn")), i + 2
        'if it is loaded show this at the matrix
        If Plugins(i).boolLoaded = True Then
            fgPlugins.TextMatrix(i + 2, 1) = Language(280) 'Loaded: Yes
        End If
    Next i
    
    'if there are some plugins installed
    If fgPlugins.rows > 2 Then
        'remove the first entry - it is empty
        fgPlugins.RemoveItem 1
    End If
End Sub
Private Sub ListSoundSchemes()
    Dim myFile As File
    Dim myFolder As Folder
    
    
    icSoundSchemes.ComboItems.Clear
    
    For Each myFolder In FS.GetFolder(App.Path & "/data/sounds").SubFolders
        For Each myFile In myFolder.Files
            If LCase$(Right$(myFile.Name, 4)) = ".xml" Then
                icSoundSchemes.ComboItems.Add , , UCase$(Left$(myFile.Name, 1)) & LCase$(Mid$(myFile.Name, 2, Len(myFile.Name) - 5))
                If LCase$(Left$(myFile.Name, Len(myFile.Name) - 4)) = LCase$(Options.SoundScheme) Then
                    Set icSoundSchemes.SelectedItem = icSoundSchemes.ComboItems.Item(icSoundSchemes.ComboItems.Count)
                End If
            End If
        Next myFile
    Next myFolder
End Sub
Private Sub ListSmileyPacks()
    Dim myFile As File
    Dim myFolder As Folder
    
    icSmileyPacks.ComboItems.Clear
    
    For Each myFolder In FS.GetFolder(App.Path & "/data/smileys").SubFolders
        For Each myFile In myFolder.Files
            If LCase$(Right$(myFile.Name, 4)) = ".xml" Then
                icSmileyPacks.ComboItems.Add , , UCase$(Left$(myFile.Name, 1)) & LCase$(Mid$(myFile.Name, 2, Len(myFile.Name) - 5))
                If LCase$(Left$(myFile.Name, Len(myFile.Name) - 4)) = LCase$(Options.SmileyPack) Then
                    Set icSmileyPacks.SelectedItem = icSmileyPacks.ComboItems.Item(icSmileyPacks.ComboItems.Count)
                End If
            End If
        Next myFile
    Next myFolder
End Sub
Public Sub Scripting_LoadImage(ByVal Index As Integer)
    'this routine is used by CodeBehind blocks of Skins
    'to load images that should appear in the options
    'form, as VBScript doesn't have a Load method.
    Load picCustom(Index)
End Sub
Private Sub OptionsLoadLanguage()
    Dim strBegin As String
    Dim strMid As String
    Dim strEnd As String

    DB.Enter "OptionsLoadLanguage"
    'Sub used to translate items of the options dialog interface
    'get the caption from the Language array and assign it to the
    'right item.
    Me.Caption = Language(71)
    lblNickname.Caption = Language(79)
    lblAlt.Caption = Language(80)
    lblAlt2.Caption = Language(81)
    lblEmail.Caption = Language(123)
    lblReal.Caption = Language(128)
    cmdAdd(0).Caption = Language(84)
    cmdRemove(0).Caption = Language(83)
    cmdClear(0).Caption = Language(93)
    cmdAdd(1).Caption = Language(84)
    cmdRemove(1).Caption = Language(83)
    cmdClear(1).Caption = Language(93)
    cmdEdit.Caption = Language(89)
    lblSelectLang.Caption = Language(95)
    lblScripts.Caption = Language(88)
    chkLogChannels.Caption = Language(92)
    chkLogPrivates.Caption = Language(91)
    cmdOK.Caption = Language(120)
    cmdCancel.Caption = Language(121)
    cmdApply.Caption = Language(122)
    lblNicklist.Caption = Language(144)
    lblWelcome.Caption = Language(143)
    cmdAddnick.Caption = Language(84)
    cmdREBdy.Caption = Language(145)
    chkLatest.Caption = Language(806)
    chkFade.Caption = Language(164)
    chkPerform.Caption = Language(167)
    chkXPCommonControls.Caption = Language(168)
    lblBuddyEnter.Caption = Language(192)
    lblBuddyLeave.Caption = Language(193)
    chkEnterMsg.Caption = Language(194)
    chkLeaveMsg.Caption = Language(194)
    chkEnterWin.Caption = Language(195)
    chkLeaveWin.Caption = Language(195)
    cmdDeleteScript.Caption = Language(197)
    lblSkin.Caption = Language(200)
    lblDownloadSkins.Caption = Language(221)
    chkHTMLError.Caption = Language(232)
    chkLoading.Caption = Language(233)
    cmdEditBdy.Caption = Language(89)
    chkCodeBehind.Caption = Language(238)
    chkScripting.Caption = Language(239)
    chkAutocomplete.Caption = Language(243)
    chkTOD.Caption = Language(248)
    lblViewLogs.Caption = Language(249)
    lblSessionNormal.Caption = Language(272)
    lblSessionCrash.Caption = Language(271)
    optSessionNR.Caption = Language(273)
    optSessionND.Caption = Language(274)
    optSessionNA.Caption = Language(275)
    optSessionCR.Caption = Language(273)
    optSessionCD.Caption = Language(274)
    optSessionCA.Caption = Language(275)
    chkLogByNet.Caption = Language(283)
    chkInfoTips.Caption = Language(286)
    lblBuddies.Caption = Language(310)
    lblIgnored.Caption = Language(311)
    lblEverybody.Caption = Language(312)
    lblOption(0).Caption = Language(313)
    lblOption(3).Caption = Language(313)
    lblOption(6).Caption = Language(313)
    lblOption(1).Caption = Language(77)
    lblOption(4).Caption = Language(77)
    lblOption(7).Caption = Language(77)
    lblOption(2).Caption = Language(314)
    lblOption(5).Caption = Language(314)
    lblOption(8).Caption = Language(314)
    lblOption(9).Caption = Language(273)
    lblOption(12).Caption = Language(273)
    lblOption(10).Caption = Language(274)
    lblOption(13).Caption = Language(274)
    lblOption(11).Caption = Language(275)
    lblOption(14).Caption = Language(275)
    chkTray.Caption = Language(317)
    chkAntivirus.Caption = Language(433)
    chkRestoreStatus.Caption = Language(441)
    chkAutoNDC.Caption = Language(442)
    lblIgnore(0).Caption = Language(623)
    lblIgnore(1).Caption = Language(624)
    lblViewDownloads.Caption = Language(445)
    chkJoinPanel.Caption = Language(455)
    lblDownloadLang.Caption = Language(477)
    cmdAddHotKey.Caption = Language(84)
    cmdRemoveHotKey.Caption = Language(83)
    fgHotKeys.TextMatrix(0, 0) = Language(490)
    fgHotKeys.TextMatrix(0, 1) = Language(491)
    lblInstalledPlugs.Caption = Language(508)
    lblDownloadPlugs.Caption = Language(509)
    fgPlugins.TextMatrix(0, 0) = Language(526)
    fgPlugins.TextMatrix(0, 1) = Language(527)
    mnuPlugInOptions.Caption = Language(4) & "..."
    mnuPluginStartup.Caption = Language(557)
    chkProxyEnable.Caption = Language(574)
    lblProxy.Caption = Language(588)
    lblProxyPort.Caption = Language(589)
    chkCTCPPing.Caption = Language(629)
    chkCTCPPingIgnore.Caption = Language(630)
    chkCTCPVersion.Caption = Language(639)
    chkCTCPVersionIgnore.Caption = Language(640)
    chkCTCPTime.Caption = Language(641)
    chkCTCPTimeIgnore.Caption = Language(642)
    chkCTCPFlood.Caption = Language(646)
    chkCTCPFloodBounce.Caption = Language(647)
    chkTimeStamp.Caption = Language(652)
    chkTimeStampChannels.Caption = Language(320)
    chkTimeStampPrivates.Caption = Language(655)
    chkTimeStampStatus.Caption = Language(70)
    chkStartPage.Caption = Language(658)
    chkTimeStampsLog.Caption = Language(659)
    chkFocusJoined.Caption = Language(670)
    chkLogRAW.Caption = Language(706)
    chkCTCPVersionCustom.Caption = Language(713)
    lblDCCReceive.Caption = Language(714)
    chkSelectionCopy.Caption = Language(736)
    chkSelectionClear.Caption = Language(737)
    lblInChannels.Caption = Language(741)
    lblInPrivates.Caption = Language(742)
    lblNickLinkChan.Caption = Language(739)
    lblNickLinkPriv.Caption = Language(739)
    lblNickLinkMineChan.Caption = Language(740)
    lblNickLinkMinePriv.Caption = Language(740)
    lblExportDescription.Caption = Language(745) '& vbnewline & Language(751)
    cmdExport.Caption = Language(746) & "..."
    cmdImport.Caption = Language(747) & "..."
    cdExport.DialogTitle = Language(748)
    cdImport.DialogTitle = Language(747)
    cdExport.Filter = Language(750) & " (*.xml)|*.xml"
    cdImport.Filter = Language(750) & " (*.xml)|*.xml"
    lblStartupConnect.Caption = Language(757)
    lblRetry.Caption = Language(777)
    lblRetryDelay.Caption = Language(778)
    lblRetryDelaySeconds.Caption = LCase$(Language(251))
    fraPortR.Caption = Language(794)
    lblRejoinOnKick.Caption = Language(804)
    lblJoinOnInvite.Caption = Language(805)
    lblModesOnJoin.Caption = Language(820)
    lblKeepChannelsOpen.Caption = Language(868)
    lblDownloadSoundSchemes.Caption = Language(878)
    lblSmileyPacks.Caption = Language(887)
    lblDownloadSmileyPacks.Caption = Language(888)
    lblSoundScheme.Caption = Language(866)
    lblUseAway.Caption = Language(901)
    lblAwayMinutes.Caption = Language(902)
    lblAwayChangeNick.Caption = Language(903)
    lblAwayPerform.Caption = Language(905)
    lblAwayBackChangeNick.Caption = Language(903)
    lblAwayBackPerform.Caption = Language(905)
    lblAwayGo.Caption = Language(931)
    lblAwayBackGo.Caption = Language(928)
    lblBrowseUsing.Caption = Language(909)
    lblBrowseParseLinks.Caption = Language(912)
    lblBrowseDefaultBrowser.Caption = Language(910)
    lblBrowseInternal.Caption = Language(911)
    lblQuitSingle.Caption = Language(934)
    lblQuitMulti.Caption = Language(935)
    lblParseMemoServ.Caption = Language(937) & " (" & Language(938) & ")"
    cdQuitList.DialogTitle = Language(936)
    chkEnableSmileys.Caption = Language(939)
    lblIdent(0).Caption = Language(942)
    lblIdent(1).Caption = Language(946)
    lblIdent(2).Caption = Language(947)
    lblIdent(3).Caption = Language(948)
    lblIdent(4).Caption = Language(943)
    lblIdent(5).Caption = Language(944) & " (UNIX)"
    lblIdent(6).Caption = Language(945)
    
    DB.Leave "OptionsLoadLanguage"
End Sub
Public Sub SetDisplaySkinFontSize()
    On Error Resume Next
    frmMain.webdocChanMain.body.Style.FontSize = Options.DisplaySkinFontSize
    frmMain.webdocPrivates.body.Style.FontSize = Options.DisplaySkinFontSize
    frmMain.webdocChanNicklist.body.Style.FontSize = Options.DisplaySkinFontSize
    frmMain.webdocChanFrameSet.body.Style.FontSize = Options.DisplaySkinFontSize
    frmMain.webdocDCCs.body.Style.FontSize = Options.DisplaySkinFontSize
    'frmMain.webdocSplit.body.Style.fontFamily = Options.DisplaySkinFont
End Sub

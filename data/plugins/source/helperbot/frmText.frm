VERSION 5.00
Begin VB.Form frmText 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtOnJoin 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter text that you want everyone to receive when they enter the channel:"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Me.Hide
    SaveSetting "Node.PlugIns", "HelperBot", "WelcomeText", txtOnJoin.Text
End Sub
Private Sub Form_Load()
    txtOnJoin.Text = GetSetting("Node.PlugIns", "HelperBot", "WelcomeText", "")
End Sub

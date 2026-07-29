VERSION 5.00
Begin VB.Form usermenu 
   BackColor       =   &H8000000D&
   Caption         =   "User Menu"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   120
   End
   Begin VB.CommandButton cmdCurrencies 
      Caption         =   "Change Currencies"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdChangePin 
      Caption         =   "Change Pin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "Deposit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdWithdraw 
      Caption         =   "Withdraw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "MartATM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "usermenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1


Private Sub cmdBack_Click()
Me.Hide
End Sub

Private Sub cmdChangePin_Click()
UserOperation = 4
NeedsHelp = False
changepin.Init
End Sub

Private Sub cmdCurrencies_Click()
currencyexchange.Init
End Sub

Private Sub cmdDeposit_Click()
UserOperation = 1
NeedsHelp = False
userinput.Init

End Sub

Private Sub cmdTransfer_Click()
transfer.Init
End Sub

Private Sub cmdWithdraw_Click()
UserOperation = 2
NeedsHelp = False
userinput.Init

End Sub

Private Sub Form_Load()
Me.Hide
End Sub

Public Sub Unload()
Me.Hide
Set Voice = Nothing
Me.Unload
End Sub

Public Sub Speak(text As String)
Voice.Speak text, 1 '1 = SVSFlagsAsync for non-blocking
End Sub

Public Sub Init()
    Me.Show
    Set Voice = New SpeechLib.SpVoice
    Voice.Volume = 100
    Voice.Rate = 0
    
    'call isFinished timer
    Call Timer1_Timer
    
    'resize to global resolution
    Me.Width = globalWidth
    Me.Height = globalHeight
    
    'greet the user
    Speak ("Please select an option.")
End Sub

Private Sub Timer1_Timer()
    If isFinished = False Then
        Timer1.Interval = 250
        Timer1.Enabled = True
    ElseIf isFinished = True Then
        Speak ("Thank you for using MartATM!")
        MsgBox "Thank you for using MartATM!", vbInformation
        Me.Hide
        isFinished = False
    End If
End Sub

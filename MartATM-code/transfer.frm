VERSION 5.00
Begin VB.Form transfer 
   BackColor       =   &H8000000D&
   Caption         =   "Transfer"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   1560
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
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
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtReceiving 
      Height          =   405
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtSending 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Receiving account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Sending account:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Transfer"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Private isCorrect As Boolean
Private sendingAccount, receivingAccount As Long

Private Sub cmdBack_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
txtReceiving.text = ""
txtSending.text = ""
End Sub

Public Sub Speak(text As String)
Voice.Speak text, 1 '1 = SVSFlagsAsync for non-blocking
End Sub

Public Sub Init()
    Me.Show
    Set Voice = New SpeechLib.SpVoice
    Voice.Volume = 100
    Voice.Rate = 0
    
    'hide sensitive information
    txtReceiving.PasswordChar = "*"
    txtSending.PasswordChar = "*"
    
    'greet the user
    Speak ("Input transfer and receiving accounts")
End Sub

Public Sub Unload()
Me.Hide
Set Voice = Nothing
Me.Unload
End Sub

Private Sub cmdNext_Click()

'checks for empty fields
If txtSending.text = "" Or txtReceiving.text = "" Then
    Speak ("Fields cannot be empty!")
    MsgBox "Fields cannot be empty!", vbExclamation
    isCorrect = False
Else
    isCorrect = True
End If

'checks correct length
If Len(txtSending) <> 28 Then
    Speak ("Incorrect sending account length")
    MsgBox "Incorrect sending account length", vbExclamation
    isCorrect = False
Else
    isCorrect = True
End If

If Len(txtReceiving) <> 28 Then
    Speak ("Incorrect receiving account length")
    MsgBox "Incorrect receiving account length", vbExclamation
    isCorrect = False
    isReceivingAccount = False
Else
    isCorrect = True
End If

'check for non-numerics
If HasNonNumeric(txtSending) Then
    Speak ("Improper characters in sending account!")
    MsgBox "Improper characters in sending account!", vbExclamation
    isCorrect = False
Else
    isCorrect = True
End If

If HasNonNumeric(txtReceiving) Then
    Speak ("Improper characters in sending account!")
    MsgBox "Improper characters in sending account!", vbExclamation
    isCorrect = False
Else
    isCorrect = True
End If

'check if bank accounts are the same account
If txtSending.text = txtReceiving.text Then
    Speak ("Accounts cannot be the same!")
    MsgBox "Accounts cannot be the same!", vbExclamation
    isCorrect = False ' do not proceed further with double account
Else
    isCorrect = True
End If

    'proceeds if isCorrect = True
    If isCorrect = True Then
    
        'put user input into variables
        sendingAccount = CLng(txtSending.text)
        receivingAccount = CLng(txtReceiving.text)
        
        'msg user
        Speak ("Please input card")
        MsgBox "Please input card.", vbInformation
        
        'set flag to false
        isFinished = False
        
        'open PIN verification
        Call pininput.Init(False, sendingAccount)
        
        'call timer checking for isFinished flag
        Call Timer1_Timer
        
    End If

End Sub

Private Sub Form_Load()
Me.Hide
End Sub

Private Sub Timer1_Timer()
    If isFinished = False Then
        Timer1.Interval = 250
        Timer1.Enabled = True
    ElseIf isFinished = True Then
        Me.Hide
        'clean up
        txtSending = ""
        txtReceiving = ""
        Timer1.Enabled = False
    End If
End Sub

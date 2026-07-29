VERSION 5.00
Begin VB.Form pininput 
   BackColor       =   &H8000000D&
   Caption         =   "Input PIN"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBankAccount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4335
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
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
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
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label lblBankAccount 
      BackColor       =   &H8000000D&
      Caption         =   "Input Bank Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Input PIN"
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
      Width           =   2295
   End
End
Attribute VB_Name = "pininput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Private BankAccountArg As String
Private ShowBankAccountArg, isCorrect As Boolean

Private Sub cmdBack_Click()
Me.Hide
End Sub

Private Sub cmdNext_Click()

    'check if PIN is empty
    If txtInput.text = "" Then
        Speak ("Fields cannot be empty!")
        MsgBox "Fields cannot be empty!", vbExclamation
        isCorrect = False
    Else
        isCorrect = True
    End If
    
    'check if bank account is empty
    If ShowBankAccountArg = True Then
        If txtBankAccount.text = "" Then
            Speak ("Fields cannot be empty!")
            MsgBox "Fields cannot be empty!", vbExclamation
            isCorrect = False
        Else
            isCorrect = True
        End If
    End If

    'checks for non-numeric - pin input
    If HasNonNumeric(txtInput.text) = True Then
        Speak ("PIN contains improper characters")
        MsgBox "PIN contains improper characters", vbInformation
        isCorrect = False
    Else
        isCorrect = True
    End If
    
    'checks for non-numeric - bank account input
    If ShowBankAccountArg = True Then
        If HasNonNumeric(txtBankAccount.text) = True Then
            Speak ("Bank account contains improper characters!")
            MsgBox "Bank account contains improper characters!", vbExclamation
            isCorrect = False
        Else
            isCorrect = True
        End If
    End If
    
    'checks pin length
    If Len(txtInput) <> 4 Then
        Speak ("Incorrect PIN length")
        MsgBox "Incorrect PIN length", vbExclamation
        isCorrect = False
    Else
        isCorrect = True
    End If
    
    'checks bank account length
    If ShowBankAccountArg = True Then
        If Len(txtBankAccount) <> 29 Then
            Speak ("Incorrect bank account length")
            MsgBox "Incorrect bank account length", vbExclamation
            isCorrect = False
        Else
            isCorrect = True
        End If
    End If
    
    'proceed if isCorrect = True
    If isCorrect = True Then
        MsgBox "COBOL PIN and bank account verification here" & vbNewLine & "PIN is bound to bank account (debit card)", vbInformation
        txtInput = ""
        txtBankAccount = ""
                
        Me.Hide
        'set flag for transfer to hide
        isFinished = True
        isAllowed = True
    End If
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

Public Sub Init(ShowBankAccount As Boolean, ByVal BankAccount As String)
    Me.Show
    Set Voice = New SpeechLib.SpVoice
    Voice.Volume = 100
    Voice.Rate = 0
    
    'cover PIN on screen
    txtInput.PasswordChar = "*"
    
    
    'greet the user
    Speak ("Input PIN to verify transaction")
    
    'set the variables for next to use
    ShowBankAccountArg = ShowBankAccount
    BankAccountArg = BankAccount
    
    'check whether the program should show bank account text box
    If ShowBankAccount = True Then
        lblBankAccount.Visible = True
        txtBankAccount.Visible = True
        txtBankAccount.PasswordChar = "*" 'hide bank account digits
    ElseIf ShowBankAccount = False Then
        lblBankAccount.Visible = False
        txtBankAccount.Visible = False
    End If
    
    
End Sub


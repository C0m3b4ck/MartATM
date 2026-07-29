VERSION 5.00
Begin VB.Form userinput 
   BackColor       =   &H8000000D&
   Caption         =   "Input Data"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboOptions 
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtAmmount 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtBankAccount 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
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
      Left            =   1080
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtCardNumber 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   2655
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
      TabIndex        =   2
      Top             =   3120
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
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Selection:"
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
      Left            =   0
      TabIndex        =   11
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Ammount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Bank account number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Credit card number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Input Data"
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
      Width           =   2415
   End
End
Attribute VB_Name = "userinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Private isCorrect As Boolean

Private Sub cmdBack_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
txtCardNumber = ""
txtBankAccount = ""
txtAmmount = ""
End Sub

Private Sub cmdNext_Click()

    ' checks for empty fields
    If txtCardNumber.text = "" Or txtBankAccount.text = "" Or txtAmmount.text = "" Then
        isCorrect = False
        Speak ("Fields cannot be empty!")
        MsgBox "Fields cannot be empty!", vbExclamation
    Else
        isCorrect = True
    End If
    
    'checks correct length
    If Len(txtCardNumber) <> 19 Then
        Speak ("Incorrect card number length")
        MsgBox "Incorrect card number length", vbExclamation, "Incorrect number length"
        isCorrect = False
    Else
        isCorrect = True
    End If
        
    If Len(txtBankAccount) <> 28 Then
        Speak ("Incorrect bank account length")
        MsgBox "Incorrect bank account length", vbExclamation, "Incorrect bank account length"
        isCorrect = False
    Else
        isCorrect = True
    End If
        
    If Len(txtAmmount) > 12 Then
        Speak ("Maximum ammount exceeded")
        MsgBox "Maximum ammount exceeded", vbExclamation, "Maximum ammount exceeded"
        isCorrect = False
    Else
        isCorrect = True
    End If
    
    ' proceeds only if isCorrect = True
        'checks to see exact user operation
      If isCorrect = True Then
        MsgBox "Similar templates will be made as COBOL functionality comes", vbInformation
            
        'exit operation
        Me.Hide
        isFinished = True
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

Public Sub Init()
    Me.Show
    Set Voice = New SpeechLib.SpVoice
    Voice.Volume = 100
    Voice.Rate = 0
    
    'wipe from previous input
    txtBankAccount.text = ""
    txtCardNumber.text = ""
    txtAmmount.text = ""
    cboOptions = ""
    
    'protect sensitive info from being viewed
    txtCardNumber.PasswordChar = "*"
    txtBankAccount.PasswordChar = "*"
    
    'greet the user
    Speak ("Please input your transaction details to proceed.")
    
    'check for user action type 1 - deposit, 2 - withdrawal, then populates combo box
    cboOptions = "" 'reset before setting again
    If UserOperation = 1 Then
        cboOptions.AddItem "Cash to bank account"
        cboOptions.AddItem "Credit card to bank account"
    ElseIf UserOperation = 2 Then
        cboOptions.AddItem "Withdraw cash"
        cboOptions.AddItem "Withdraw check"
    Else
        MsgBox "Error: invalid user operation: " & UserOperation, vbOKOnly
    End If
End Sub


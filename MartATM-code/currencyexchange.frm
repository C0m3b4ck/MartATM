VERSION 5.00
Begin VB.Form currencyexchange 
   BackColor       =   &H8000000D&
   Caption         =   "Exchange Currencies"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3960
      Top             =   2640
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
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
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
      Left            =   2160
      TabIndex        =   9
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtAmmount 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.ComboBox cboDesiredCurrency 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox cboSourceSelection 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox cboUserCurrency 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Select desired currency ammount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Exchange currency to:"
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
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Select source:"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Select your currency:"
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
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Currency Exchange"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "currencyexchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1

Private ammount As Single
Private userCurrency As String
Private source As String
Private desiredCurrency As String
Private proceed, isUserCurrency, isDesiredCurrency, isSource, isAmmount, isAmmountEmpty, isCorrect As Boolean 'flags
Private ammountSingle As Single
Private fakenone As String

Private Sub cmdBack_Click()
Me.Hide
txtAmmount.text = ""
cboUserCurrency = ""
cboSourceSelection = ""
cboDesiredCurrency = ""
End Sub

Private Sub Command1_Click()

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
    
    'create disabled 250 ms timer
    'Timer1.Interval = 250
    'Timer1.Enabled = False
    
    'populate cboBoxes
    
    'cboUserCurrency
    cboUserCurrency.AddItem "USD"
    cboUserCurrency.AddItem "PLN"
    cboUserCurrency.AddItem "EUR"
    
    'cboSourceSelection
    cboSourceSelection.AddItem "Debit Card"
    cboSourceSelection.AddItem "Cash"
    cboSourceSelection.AddItem "Bank Account"
    
    'cboDesiredCurrency
    cboDesiredCurrency.AddItem "USD"
    cboDesiredCurrency.AddItem "PLN"
    cboDesiredCurrency.AddItem "EUR"
    
    'reset old fields from previous input
    cboUserCurrency = ""
    cboSourceSelection = ""
    cboDesiredCurrency = ""
    txtAmmount = ""
    
    'greet the user
    Speak ("Fill out data before exchanging currencies.")
End Sub

Private Sub cmdNext_Click()

    ' set isAmmountEmpty flag to false (default)
    isAmmountEmpty = False
    
    ' checks for empty fields
    If txtAmmount.text = "" Then
        Speak ("Please input ammount!")
        MsgBox "Please input ammount!", vbExclamation
        isAmmountEmpty = True
    End If
    
    'check for non-alphanumeric characters
    If HasNonNumeric(txtAmmount.text) Then
        Speak ("Improper characters in ammount!")
        MsgBox "Improper characters in ammount!", vbExclamation
        isAmmount = False
    Else
        isAmmount = True
    End If
    
    ' check ammount math size, not length
    If isAmmount = True Then
        ammountSingle = CSng(txtAmmount.text)
        If ammountSingle > 3001 Then
            Speak ("Please input smaller ammount!")
            MsgBox "Please input smaller ammount!", vbExclamation
            isAmmount = False
            isCorrect = False
        Else
            isAmmount = True
        If cboUserCurrency = "" Then
            Speak ("Please input currency")
            MsgBox "Please input currency!", vbExclamation
            isUserCurrency = False
            isCorrect = False
        Else
            isUserCurrency = True
        If cboDesiredCurrency = "" Then
            Speak ("Please input desired currency")
            MsgBox "Please input desired currency!", vbExclamation
            isDesiredCurrency = False
            isCorrect = False
        Else
            isDesiredCurrency = True
        If cboSourceSelection = "" Then
            Speak ("Please select source!")
            MsgBox "Please select source!", vbExclamation
            isSource = False
            isCorrect = False
        Else
            isSource = True
        End If
        End If
        End If
        End If
    
    'check if user filled out all forms
    If isAmmount = False Or isUserCurrency = False Or isDesiredCurrency = False Or isSource = False Then
        Speak ("Make sure all of the data fields are filled out!")
        MsgBox "Make sure all of the data fields are filled out!", vbInformation
        isCorrect = False
    End If
    
    'check if user and desired currency are the same
    If cboUserCurrency = cboDesiredCurrency Then
        Speak ("Cannot convert to same currency!")
        MsgBox "Cannot convert to same currency!", vbExclamation
        isCorrect = False
        isDesiredCurrency = False
        isUserCurrency = False
    Else
        isCorrect = True
    End If
    
    'proceed if everything is correct
    If isCorrect = True Then
        'fill variables with user input
        source = cboSourceSelection
        desiredCurrency = cboDesiredCurrency
        userCurrency = cboUserCurrency
        ammount = CSng(txtAmmount.text)
        
        'call pininput for authentication
        If cboSourceSelection = "Debit Card" Or cboSourceSelection = "Bank Account" Then
            fakenone = ""
            Call pininput.Init(True, fakenone)
        ElseIf cboSourceSelection = "Cash" Then
            Speak ("Input cash into ATM. ")
            MsgBox "Input cash into ATM. ", vbInformation
        End If
        
        'checks for authentication returned by pininput - timer
        If cboSourceSelection = "Debit Card" Or cboSourceSelection = "Bank Account" Then
            Call Timer1_Timer
        Else
            Speak ("Please take out cash.")
            MsgBox "Please take out cash.", vbInformation
            
            'set isFinished flag to true
            Me.Hide
            isFinished = True
            
            'upon exiting - reset fields
            cboUserCurrency = ""
            cboSourceSelection = ""
            cboDesiredCurrency = ""
            txtAmmount = ""
            
        End If

    End If
    End If
End Sub

Private Sub OnceAuthenticated()
'reset old fields from previous input
    cboUserCurrency = ""
    cboSourceSelection = ""
    cboDesiredCurrency = ""
    txtAmmount = ""
    
    MsgBox "Its COBOL time!", vbInformation

'If (cboOptions = "Withdraw check" And isAmmount = True And isBankAccount = True) Then
            'MsgBox "Similar templates will be made as COBOL functionality comes", vbInformation
        'Else
            'MsgBox "Different case will also be handled with a lot of COBOL" & vbNewLine & cboOptions & isAmmount & isBankAccount, vbOKOnly
        'End If
End Sub
    
Private Sub Timer1_Timer()
    If isAllowed = False Then
        Timer1.Interval = 250
        Timer1.Enabled = True
    ElseIf isAllowed = True Then
        Timer1.Enabled = False
        Speak ("Authenticated! Proceeding!")
        MsgBox "Authenticated! Proceeding!", vbInformation
        Call OnceAuthenticated ' runs function checking exact user operation
    End If
End Sub


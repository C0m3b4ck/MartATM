VERSION 5.00
Begin VB.Form changepin 
   BackColor       =   &H8000000D&
   Caption         =   "Change Pin"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   9
      Top             =   2160
      Width           =   2295
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
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtNew2 
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
      Left            =   2160
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtNew1 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtOld 
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
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Confirm new PIN:"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "New PIN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
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
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Old PIN:"
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
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Change Pin"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "changepin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1

Private OldPin, NewPin1, NewPin2 As Long

Private isCorrect As Boolean

Private Sub cmdBack_Click()
Me.Hide
End Sub

Private Sub cmdClear_Click()
txtOld.text = ""
txtNew1.text = ""
txtNew2.text = ""
End Sub

Public Sub Unload()
Me.Hide
Set Voice = Nothing
Me.Unload
End Sub

Private Sub cmdSubmit_Click()
    
    'check for non-numerics
    If HasNonNumeric(txtOld.text) Then
        Speak ("Incorrect PIN characters in old pin!")
        MsgBox "Incorrect PIN characters in old pin!", vbExclamation
        isCorrect = False
    End If
    
    If HasNonNumeric(txtNew1.text) Then
        Speak ("Incorrect PIN characters in new PIN!")
        MsgBox "Incorrect PIN characters in new PIN!", vbExclamation
        isCorrect = False
    End If
    
    If HasNonNumeric(txtNew2.text) Then
        Speak ("Incorrect PIN characters in PIN confirmation!")
        MsgBox "Incorrect PIN characters in PIN confirmation!", vbExclamation
        isCorrect = False
    End If
    
    'check if fields are empty
    If txtOld.text = "" Or txtNew1.text = "" Or txtNew2.text = "" Then
        Speak ("Fields cannot be empty!")
        MsgBox "Fields cannot be empty!", vbExclamation
        isCorrect = False
    End If
    
    'compare both new pins for difference
    If Not txtNew1.text = txtNew2.text Then
        Speak ("New PINs are not the same!")
        MsgBox "New PINs are not the same!", vbExclamation
        isCorrect = False
    Else
        isCorrect = True
    End If
    
    'if all goes well - isCorrect = true and proceed
    If isCorrect = True Then
        OldPin = CLng(txtOld.text)
        NewPin1 = CLng(txtNew1.text)
        
        'COBOL logic here
        
        Speak ("PIN changed succesfully! Please take out card!")
        MsgBox "PIN changed succesfully! Please take out card!", vbInformation
        
        'exit procedure
        isFinished = True
        Me.Hide
    End If

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
    txtOld.PasswordChar = "*"
    txtNew1.PasswordChar = "*"
    txtNew2.PasswordChar = "*"
    
    'wipe from previous input
    txtOld.text = ""
    txtNew1.text = ""
    txtNew2.text = ""


Me.Show
Speak ("Please input card")
MsgBox "Please input card", vbInformation, "Card input required"
Speak ("Input old PIN, then new PIN twice")
End Sub

Private Sub Form_Load()
Me.Hide
End Sub

VERSION 5.00
Begin VB.Form frontpage 
   BackColor       =   &H8000000D&
   Caption         =   "Front Page"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox lang8 
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox lang7 
      Height          =   495
      Left            =   2400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox lang6 
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox lang5 
      Height          =   495
      Left            =   1200
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox lang4 
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox lang3 
      Height          =   495
      Left            =   3480
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox lang2 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox lang1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox logo 
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1155
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Select language"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Welcome to MartATM!"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frontpage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Private waitingForKey As Boolean

Private Sub Form_Load()
Call LoadAllForms
Me.Init
End Sub

Public Sub Unload()
Me.Hide
Set Voice = Nothing
Me.Unload
Call UnloadAllForms
End Sub

Public Sub Speak(text As String)
Voice.Speak text, 1 '1 = SVSFlagsAsync for non-blocking
End Sub

Public Sub Init()
    Me.Show
    Set Voice = New SpeechLib.SpVoice
    Voice.Volume = 100
    Voice.Rate = 0
    
    'resize to global resolution
    Me.Width = globalWidth
    Me.Height = globalHeight
    
    'load images from images folder - change the loading names before use
    logo.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    lang1.Picture = LoadPicture(App.Path & "\icons\eus.ico")
    lang2.Picture = LoadPicture(App.Path & "\icons\pol.ico")
    lang3.Picture = LoadPicture(App.Path & "\icons\ukr.ico")
    'lang4.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    'lang5.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    'lang6.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    'lang7.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    'lang8.Picture = LoadPicture(App.Path & "\icons\logo.ico")
    
    
    'greet the user
    Speak ("Welcome to Mart Automatic Teller Machine!")
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 111 Then
        MsgBox "o", vbOKOnly
        Call lang1_Click
    End If
End Sub

Private Sub Label3_Click()

End Sub

Private Sub lang1_Click()
usermenu.Init
language = "EnglishUS"
End Sub

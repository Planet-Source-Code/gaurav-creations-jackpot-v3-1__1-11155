VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H008080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Players"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   Icon            =   "player.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      BackColor       =   &H008080FF&
      Caption         =   "Option3"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   7
      Text            =   "PLAYER 2"
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "PLAYER 1"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "Option2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H008080FF&
      Caption         =   "Option1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      Caption         =   "Two Players(LIMITED)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   1320
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H008080FF&
      Caption         =   "Two Players(FIRST LUCK)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      Caption         =   "Player2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H008080FF&
      Caption         =   "Player1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H008080FF&
      Caption         =   "One Player"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "How many Players ?"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
player1 = Text1.Text
player2 = Text2.Text
If Option1.Value = True Then
Unload Me
Form1.Show
End If
If Option2.Value = True Then
Unload Me
Form4.Show
End If
If Option3.Value = True Then
Unload Me
Form7.Show
End If
End Sub


Private Sub Option1_Click()
Form6.Height = 3345
End Sub

Private Sub Option2_Click()
Form6.Height = 4440
End Sub

Private Sub Option3_Click()
Form6.Height = 4440
End Sub

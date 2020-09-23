VERSION 5.00
Begin VB.Form startup 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   3285
   ClientTop       =   8805
   ClientWidth     =   3435
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "www.question.chatbook.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "   Gaurav CREATIONS"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
RememberScreenRes
ChangeScreenSettings 1024, 768, 16
End Sub

Private Sub Timer1_Timer()
Label2.Visible = False
Call sndPlaySound(App.Path & "\near_mis.wav", &H1)
For i = 500 To (startup.Height / 2) Step 30
        Sleep 15
      startup.Circle (startup.Width / 2, startup.Height / 2), i, RGB(4, 255, 4)
Next i
Timer1.Enabled = False
Sleep 20
Label1.Visible = True
startup.Circle (startup.Width / 2, startup.Height / 2), i, RGB(4, 255, 4)
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Sleep 2000
Call sndPlaySound(App.Path & "\headcras.wav", &H1)
For i = 0 To (startup.Height) Step 20
        Sleep 15
   FillColor = RGB(i, 0, 0)
   startup.Circle (startup.Width / 2, startup.Height / 2), i
Next i
Timer2.Enabled = False
Label1.Visible = False
startup.BackColor = vbRed
Label2.Visible = True
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Sleep 3000
Timer3.Enabled = False
Unload Me
Form6.Show
End Sub

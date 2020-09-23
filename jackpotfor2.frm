VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleMode       =   0  'User
   ScaleWidth      =   17076.92
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C000&
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5760
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LEAVE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5760
      Width           =   4215
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   720
      Top             =   5520
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1080
      Top             =   5520
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   11
      Top             =   3960
      Width           =   1935
      Begin VB.CommandButton Command2 
         Caption         =   "WIN"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H80000012&
      Height          =   495
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   1755
      TabIndex        =   10
      Top             =   3960
      Width           =   1815
      Begin VB.CommandButton Command1 
         Caption         =   "WIN"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   7440
      Picture         =   "jackpotfor2.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   6240
      Picture         =   "jackpotfor2.frx":1A2E
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   5040
      Picture         =   "jackpotfor2.frx":3500
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   480
      Picture         =   "jackpotfor2.frx":4D82
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1680
      Picture         =   "jackpotfor2.frx":67B0
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2880
      Picture         =   "jackpotfor2.frx":8282
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   5520
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   360
      Top             =   5520
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "First Luck"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "WINNER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   "WINNER"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   720
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000009&
      Height          =   3495
      Left            =   4680
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000009&
      Height          =   3495
      Left            =   120
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4680
      TabIndex        =   15
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PLAYER 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5520
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PLAYER 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   840
      TabIndex        =   8
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      Height          =   1575
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      Height          =   1575
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "JACKPOT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "JACKPOT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first As Integer
Dim second As Integer
Dim third As Integer
Dim try As Integer
Dim try1 As Integer


' Player 1

Private Sub Command1_Click()
Label7.Visible = False
Label8.Visible = False

Picture1.Picture = LoadPicture("")
Picture2.Picture = LoadPicture("")
Picture3.Picture = LoadPicture("")
For comm = 0 To 1090
Timer2.Enabled = True
Command1.Visible = True
Command1.Move comm
Next comm
Timer2.Enabled = False
For comm = 1090 To 0 Step -1
Timer2.Enabled = True
Command1.Visible = True
Command1.Move comm
Next comm
Timer2.Enabled = False
Randomize
For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 30
Picture1.Picture = LoadPicture(App.Path & "\i.bmp")
Timer1.Interval = y
Picture1.Picture = LoadPicture(App.Path & "\b.bmp")
Timer1.Interval = y + 1
Picture1.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture1.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture1.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture1.Picture = LoadPicture(App.Path & "\s.bmp")
End If
first = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 30
Picture2.Picture = LoadPicture(App.Path & "\i.bmp")
Timer1.Interval = y
Picture2.Picture = LoadPicture(App.Path & "\b.bmp")
Timer1.Interval = y + 1
Picture2.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture2.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture2.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture2.Picture = LoadPicture(App.Path & "\s.bmp")
End If
second = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 30
Picture3.Picture = LoadPicture(App.Path & "\i.bmp")
Timer1.Interval = y
Picture3.Picture = LoadPicture(App.Path & "\b.bmp")
Timer1.Interval = y + 1
Picture3.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture3.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture3.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture3.Picture = LoadPicture(App.Path & "\s.bmp")
End If
Next x
third = myvalue
Call sndPlaySound(App.Path & "\bell.wav", &H1)

try = try + 1
Label5.Caption = try
If first = second Then
compare = first
End If
If compare = third Then
Call sndPlaySound(App.Path & "\chimes2.wav", &H1)
Command1.Enabled = False
Command2.Enabled = False
try = 0
try1 = 0
Label5.Caption = try
Sleep 500
Label7.Visible = True

Label6.Caption = try
Label5.Caption = try
winner1 = winner1 + 1
End If
Command1.Enabled = False
Command2.Enabled = True
Command2.SetFocus
Shape4.FillColor = vbRed
Shape5.FillColor = vbGreen
End Sub

' Player 2

Private Sub Command2_Click()
Label7.Visible = False
Label8.Visible = False
Picture4.Picture = LoadPicture("")
Picture5.Picture = LoadPicture("")
Picture6.Picture = LoadPicture("")
For comm = 0 To 1090
Timer4.Enabled = True
Command2.Visible = True
Command2.Move comm
Next comm
Timer4.Enabled = False
For comm = 1090 To 0 Step -1
Timer4.Enabled = True
Command2.Visible = True
Command2.Move comm
Next comm
Timer4.Enabled = False
Randomize
For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 30
Picture4.Picture = LoadPicture(App.Path & "\i.bmp")
Timer3.Interval = y
Picture4.Picture = LoadPicture(App.Path & "\b.bmp")
Timer3.Interval = y + 1
Picture4.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer3.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture4.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture4.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture4.Picture = LoadPicture(App.Path & "\s.bmp")
End If
first = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer3.Enabled = True
For y = 1 To 30
Picture5.Picture = LoadPicture(App.Path & "\i.bmp")
Timer3.Interval = y
Picture5.Picture = LoadPicture(App.Path & "\b.bmp")
Timer3.Interval = y + 1
Picture5.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture5.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture5.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture5.Picture = LoadPicture(App.Path & "\s.bmp")
End If
second = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 30
Picture6.Picture = LoadPicture(App.Path & "\i.bmp")
Timer1.Interval = y
Picture6.Picture = LoadPicture(App.Path & "\b.bmp")
Timer1.Interval = y + 1
Picture6.Picture = LoadPicture(App.Path & "\s.bmp")
Next y
Timer3.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture6.Picture = LoadPicture(App.Path & "\i.bmp")
End If
If myvalue = 2 Then
Picture6.Picture = LoadPicture(App.Path & "\b.bmp")
End If
If myvalue = 3 Then
Picture6.Picture = LoadPicture(App.Path & "\s.bmp")
End If
Next x
third = myvalue
Call sndPlaySound(App.Path & "\bell.wav", &H1)

try1 = try1 + 1
Label6.Caption = try1
If first = second Then
compare = first
End If
If compare = third Then
Call sndPlaySound(App.Path & "\chimes2.wav", &H1)
Command1.Enabled = False
Command2.Enabled = False
try1 = 0
try = 0
Label6.Caption = try1
Sleep 500
Label8.Visible = True
winner2 = winner2 + 1
Label6.Caption = try
Label5.Caption = try

End If
Command2.Enabled = False
Command1.Enabled = True
Command1.SetFocus
Shape5.FillColor = vbRed
Shape4.FillColor = vbGreen
End Sub

Private Sub Command3_Click()
Unload Me
Form8.Show
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub

Private Sub Form_Load()
Label2.Caption = player1
Label4.Caption = player2
winner1 = 0
winner2 = 0
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Jackpot"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8880
   Icon            =   "jackpot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270.142
   ScaleMode       =   0  'User
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7560
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "WIN"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4950
      Width           =   735
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000008&
      Height          =   615
      Left            =   3240
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   6
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   240
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -120
      Top             =   7200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   840
      TabIndex        =   3
      Top             =   2640
      Width           =   495
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   5280
      ScaleHeight     =   1875
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3720
      ScaleHeight     =   1875
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   2160
      ScaleHeight     =   1875
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   24
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   100
      Left            =   5400
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   99
      Left            =   5160
      Shape           =   3  'Circle
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   98
      Left            =   7200
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   97
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   96
      Left            =   7680
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   95
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   94
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   93
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   92
      Left            =   7440
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   91
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   90
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   89
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   88
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   87
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   86
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   85
      Left            =   5520
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   84
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   83
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   82
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   81
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   80
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   79
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   78
      Left            =   6240
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   77
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   76
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   75
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   74
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   73
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   72
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   71
      Left            =   6720
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   70
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   69
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   68
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   67
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   66
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   65
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   64
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   63
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   62
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   61
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   60
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   59
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   58
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   57
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   56
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   55
      Left            =   4680
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   54
      Left            =   4800
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   53
      Left            =   5280
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   52
      Left            =   960
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   51
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   50
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   49
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   48
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   47
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   46
      Left            =   960
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   45
      Left            =   1440
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   44
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   43
      Left            =   1680
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   42
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   41
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   40
      Left            =   2160
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   39
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   38
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   37
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   6600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   36
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   6840
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   35
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   34
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   33
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   7080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   32
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   7320
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   31
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   30
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   6360
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   29
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   26
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   25
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   23
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   22
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   21
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   600
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   20
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   19
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   18
      Left            =   3840
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   17
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   16
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   15
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   14
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   13
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   480
      Width           =   255
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   7440
      Shape           =   1  'Square
      Top             =   4320
      Width           =   735
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   7440
      Shape           =   1  'Square
      Top             =   2400
      Width           =   735
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   7200
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   12
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   11
      Left            =   120
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   10
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   9
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   8
      Left            =   840
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   7
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   8520
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF00FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   8400
      Shape           =   3  'Circle
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   8160
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   7800
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   120
      Shape           =   3  'Circle
      Top             =   2520
      Width           =   255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   480
      Shape           =   2  'Oval
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   720
      Shape           =   1  'Square
      Top             =   2400
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   720
      Shape           =   1  'Square
      Top             =   4320
      Width           =   735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      Height          =   2415
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   5055
   End
   Begin VB.Label Label1 
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
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   2280
      Shape           =   2  'Oval
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   3480
      TabIndex        =   4
      Top             =   5400
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first As Integer
Dim second As Integer
Dim third As Integer
Dim try As Integer

Private Sub Command1_Click()
If Form1.Height = 7650 Then
Form1.Height = Form1.Height - 1350
End If
Picture1.Picture = LoadPicture("")
Picture2.Picture = LoadPicture("")
Picture3.Picture = LoadPicture("")
For comm = 3360 To 5060
Timer2.Enabled = True
Command1.Visible = True
Command1.Move comm
Next comm
Timer2.Enabled = False
For comm = 5060 To 3360 Step -1
Timer2.Enabled = True
Command1.Visible = True
Command1.Move comm
Next comm
Timer2.Enabled = False
Randomize
For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 20
Picture1.Picture = LoadPicture(App.Path & "\m.bmp")
Timer1.Interval = y
Picture1.Picture = LoadPicture(App.Path & "\n.bmp")
Timer1.Interval = y + 1
Picture1.Picture = LoadPicture(App.Path & "\z.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture1.Picture = LoadPicture(App.Path & "\m.bmp")
End If
If myvalue = 2 Then
Picture1.Picture = LoadPicture(App.Path & "\n.bmp")
End If
If myvalue = 3 Then
Picture1.Picture = LoadPicture(App.Path & "\z.bmp")
End If
first = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 20
Picture2.Picture = LoadPicture(App.Path & "\m.bmp")
Timer1.Interval = y
Picture2.Picture = LoadPicture(App.Path & "\n.bmp")
Timer1.Interval = y + 1
Picture2.Picture = LoadPicture(App.Path & "\z.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture2.Picture = LoadPicture(App.Path & "\m.bmp")
End If
If myvalue = 2 Then
Picture2.Picture = LoadPicture(App.Path & "\n.bmp")
End If
If myvalue = 3 Then
Picture2.Picture = LoadPicture(App.Path & "\z.bmp")
End If
second = myvalue
Next x
Call sndPlaySound(App.Path & "\bell.wav", &H1)


For x = 0 To 1
Timer1.Enabled = True
For y = 1 To 20
Picture3.Picture = LoadPicture(App.Path & "\m.bmp")
Timer1.Interval = y
Picture3.Picture = LoadPicture(App.Path & "\n.bmp")
Timer1.Interval = y + 1
Picture3.Picture = LoadPicture(App.Path & "\z.bmp")
Next y
Timer1.Enabled = False
myvalue = Int((3 * Rnd) + 1)
If myvalue = 1 Then
Picture3.Picture = LoadPicture(App.Path & "\m.bmp")
End If
If myvalue = 2 Then
Picture3.Picture = LoadPicture(App.Path & "\n.bmp")
End If
If myvalue = 3 Then
Picture3.Picture = LoadPicture(App.Path & "\z.bmp")
End If
Next x
third = myvalue
Call sndPlaySound(App.Path & "\bell.wav", &H1)

try = try + 1
Label2.Caption = try
If first = second Then
compare = first
End If
If compare = third Then
Command1.Enabled = False
Form1.Height = Form1.Height + 1350
try = 0
Label2.Caption = try
Sleep 500
Form2.Show
End If
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
Select Case iWidth 'put screen res back to normal
Case 640
ChangeScreenSettings 640, 480, 16
Case 800
ChangeScreenSettings 800, 600, 16
Case 1024
ChangeScreenSettings 1024, 768, 16
End Select

End
End Sub


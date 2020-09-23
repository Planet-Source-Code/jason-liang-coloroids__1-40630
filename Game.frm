VERSION 5.00
Begin VB.Form Coloriods 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colors"
   ClientHeight    =   5775
   ClientLeft      =   3195
   ClientTop       =   2475
   ClientWidth     =   8055
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8055
   Begin VB.CommandButton NewG 
      Caption         =   "New Game"
      Height          =   375
      Left            =   360
      TabIndex        =   106
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Timer Everythingelse 
      Interval        =   1
      Left            =   0
      Top             =   2640
   End
   Begin VB.CommandButton About 
      Caption         =   "About"
      Height          =   375
      Left            =   1080
      TabIndex        =   104
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Help 
      Caption         =   "Help"
      Height          =   375
      Left            =   360
      TabIndex        =   103
      Top             =   1920
      Width           =   735
   End
   Begin VB.OptionButton Fast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   360
      TabIndex        =   102
      Top             =   3120
      Width           =   1455
   End
   Begin VB.OptionButton Normal 
      Caption         =   "Normal"
      Height          =   255
      Left            =   360
      TabIndex        =   101
      Top             =   2760
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Faster 
      Caption         =   "Faster"
      Height          =   255
      Left            =   360
      TabIndex        =   100
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Timer Everything 
      Left            =   0
      Top             =   2520
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   95
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   99
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   94
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   98
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   93
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   97
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   92
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   96
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   91
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   95
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   90
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   94
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   89
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   93
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   88
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   92
      Top             =   5280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   87
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   91
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   86
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   90
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   85
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   89
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   84
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   88
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   83
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   87
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   82
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   86
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   81
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   85
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   80
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   84
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   79
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   83
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   78
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   82
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   77
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   81
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   76
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   80
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   75
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   79
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   74
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   78
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   73
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   77
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   72
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   76
      Top             =   4320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   71
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   75
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   70
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   74
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   69
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   73
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   68
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   72
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   67
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   71
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   66
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   70
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   65
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   69
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   64
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   68
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   63
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   67
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   62
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   66
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   61
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   65
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   60
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   64
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   59
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   63
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   58
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   62
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   57
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   61
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   56
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   60
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   55
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   59
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   54
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   58
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   53
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   57
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   52
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   56
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   51
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   50
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   54
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   49
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   53
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   48
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   52
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   47
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   51
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   46
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   45
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   49
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   44
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   43
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   47
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   42
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   41
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   45
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   40
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   44
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   39
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   43
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   38
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   37
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   41
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   36
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   40
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   35
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   34
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   33
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   32
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   36
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   31
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   35
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   30
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   34
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   29
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   28
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   32
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   27
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   31
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   26
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   25
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   29
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   24
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   28
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   23
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   22
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   21
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   20
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   24
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   19
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   18
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   17
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   19
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   18
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   17
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   7320
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   6600
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   5880
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   5160
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3720
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   3000
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Box 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   2280
      ScaleHeight     =   465
      ScaleWidth      =   705
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Start 
      Caption         =   "Start"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Timer Timer 
      Left            =   0
      Top             =   2400
   End
   Begin VB.Timer Goingon 
      Left            =   0
      Top             =   2280
   End
   Begin VB.Label Hi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   105
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   5760
      Y2              =   0
   End
   Begin VB.Label Time 
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Score 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Color 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Color"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Shape Current 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   360
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   8040
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "Coloriods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 by Jason Liang
'Written/Coded by Jason Liang
'This is my first game... and it's updated!
'I'm currently trying to make games for no apparent reason.
'I'm 14.
'Website: www.angelfire.com/games5/gzone/
'The Website is me and my friend's...
'I just started to make programs so you shouldn't find much stuff there.
'Please VOTE for my thing.
'Now with useful little sound module!

Dim Colors As Integer   'The colors... DUH!!
Dim Clickcolor As Integer   'The current clicking color
Dim Block As Integer    'The block number
Dim timing As Integer   'The timing
Dim Scores As Integer   'Hmmm... I wonder what this is...
Dim GameOver As Boolean 'Hmmm... I wonder what this is too...
Dim V As Integer    'Number of blocks all together
Dim bonus As Integer    'The bonus (depending on how fast)
Dim Highscore As Integer    'Highscore
Dim Times As Integer    'I have no clue what this is...
Dim NewGame As Boolean  'NewGame?

'What's this about?
Private Sub About_Click()
    MsgBox "Coloroids is made by Jason Liang. Copyright 2002 by Jason Liang. Version 3.0!"
End Sub

'When you click on a box....
Private Sub Box_Click(Index As Integer)
    Select Case Clickcolor
    Case 0:
        If Box(Index).BackColor = RGB(10, 136, 239) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 30 + bonus
        Call PlaySnd(ArwHt)
        End If
    Case 1:
        If Box(Index).BackColor = RGB(255, 0, 0) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 20 + bonus
        Call PlaySnd(Chop)
        End If
    Case 2:
        If Box(Index).BackColor = RGB(0, 255, 0) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 40 + bonus
        Call PlaySnd(Drop)
        End If
    Case 3:
        If Box(Index).BackColor = RGB(0, 0, 255) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 50 + bonus
        Call PlaySnd(Glsbrk)
        End If
    Case 4:
        If Box(Index).BackColor = RGB(255, 255, 255) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 60 + bonus
        Call PlaySnd(Slap)
        End If
    Case 5:
        If Box(Index).BackColor = RGB(35, 35, 35) Then
        Box(Index).Visible = False
        Clickcolor = Rnd * 5
        Scores = Scores + 70 + bonus
        Call PlaySnd(Whip)
        End If
    End Select
End Sub

'This times EVERYTHING!! (hence the name)
Private Sub Everything_Timer()
    Score.Caption = "SCORE: " & Scores
    Select Case Clickcolor      'make a random color
    Case 0:
        Current.FillColor = RGB(10, 136, 239)
    Case 1:
        Current.FillColor = RGB(255, 0, 0)
    Case 2:
        Current.FillColor = RGB(0, 255, 0)
    Case 3:
        Current.FillColor = RGB(0, 0, 255)
    Case 4:
        Current.FillColor = RGB(255, 255, 255)
    Case 5:
        Current.FillColor = RGB(35, 35, 35)
    End Select
    If GameOver = True Then
    MsgBox "GAME OVER!!! SCORE: " & Scores
    If Scores > Highscore Then
        Highscore = Scores
    MsgBox "YOU GOT A HIGHSCORE!!"
    Open "C:\do.dll" For Binary As #1
        Put #1, , Highscore
    Close #1
    End If
    End
    End If
End Sub

'Times everythingelse (hehe name again)
Private Sub Everythingelse_Timer()
    If NewGame = True Then      'Sets values
        Everything.Interval = 0
        Scores = 0
        Score.Caption = "SCORE: "
        Time.Caption = "Time: "
        Start.Enabled = True
        Current.Visible = False
        Normal.Value = True
        For i = 0 To 95         'makes all boxes invisible!!
            Box(i).Visible = False
        Next i
        NewG.Enabled = False
        NewGame = False
        Goingon.Interval = 0
        Timer.Interval = 0
        Times = 0       'I still don't know what this is...
    End If
End Sub

'Load a form
Private Sub Form_Load()
    Open "c:\do.dll" For Binary As #1   'Opens the data file...
        Get #1, , Highscore
    Close #1
    NewG.Enabled = False
    Hi.Caption = "Highscore: " & Highscore    'Writes the highscore
End Sub

'End a form
Private Sub Form_Terminate()
    Open "C:\do.dll" For Binary As #1
        Put #1, , Highscore
    Close #1
End Sub

'Unloading a form
Private Sub Form_Unload(Cancel As Integer)
    Open "C:\do.dll" For Binary As #1
        Put #1, , Highscore
    Close #1
End Sub

'What's going on?
Private Sub Goingon_Timer()
    On Error GoTo 2
    V = 0
    Block = Rnd * 95
1
    V = V + 1
    If V = 95 Then GameOver = True
    Block = Block + 1
    If Block > 95 Then Block = 0
    Colors = Rnd * 5
    If Box(Block).Visible = True Then GoTo 1
    Select Case Colors      'Randomly colors boxes
    Case 0:
        Box(Block).BackColor = RGB(10, 136, 239)
        Box(Block).Visible = True
    Case 1:
        Box(Block).BackColor = RGB(255, 0, 0)
        Box(Block).Visible = True
    Case 2:
        Box(Block).BackColor = RGB(0, 255, 0)
        Box(Block).Visible = True
    Case 3:
        Box(Block).BackColor = RGB(0, 0, 255)
        Box(Block).Visible = True
    Case 4:
        Box(Block).BackColor = RGB(255, 255, 255)
        Box(Block).Visible = True
    Case 5:
        Box(Block).BackColor = RGB(35, 35, 35)
        Box(Block).Visible = True
    End Select
    GoTo 3
2
    GameOver = True     'Is the game over??
3
End Sub

'HELP!!
Private Sub Help_Click()
    MsgBox "Look on the bottom right of the screen. There is a thing that says current color. When the game starts, it will turn a different color. On the left of the screen, boxes will begin to appear. Click on the box of the current color. The current color will change. Keep doing it!"
End Sub

Private Sub NewG_Click()
    NewGame = True
End Sub

'Starts the game
Private Sub Start_Click()
    timing = 0
    Everything.Interval = 1
    If Faster.Value = True Then     'The speeds
        Current.Visible = True
        Clickcolor = Rnd * 5
        Goingon.Interval = 250
        Everything.Interval = 1
        bonus = 100
        Timer.Interval = 1000
    Select Case Clickcolor      'Make first clickcolor
    Case 0:
        Current.FillColor = RGB(10, 136, 239)
    Case 1:
        Current.FillColor = RGB(255, 0, 0)
    Case 2:
        Current.FillColor = RGB(0, 255, 0)
    Case 3:
        Current.FillColor = RGB(0, 0, 255)
    Case 4:
        Current.FillColor = RGB(255, 255, 255)
    Case 5:
        Current.FillColor = RGB(35, 35, 35)
    End Select
    End If
    If Normal.Value = True Then
        Goingon.Interval = 750
        Current.Visible = True
        Clickcolor = Rnd * 5
        timing = 0
        Everything.Interval = 1
        bonus = 0
        Timer.Interval = 1000
    Select Case Clickcolor
    Case 0:
        Current.FillColor = RGB(10, 136, 239)
    Case 1:
        Current.FillColor = RGB(255, 0, 0)
    Case 2:
        Current.FillColor = RGB(0, 255, 0)
    Case 3:
        Current.FillColor = RGB(0, 0, 255)
    Case 4:
        Current.FillColor = RGB(255, 255, 255)
    Case 5:
        Current.FillColor = RGB(35, 35, 35)
    End Select
    End If
    If Fast.Value = True Then
    Goingon.Interval = 500
        timing = 0
        Everything.Interval = 1
        Timer.Interval = 1000
        Clickcolor = Rnd * 5
        Current.Visible = True
        bonus = 50
        Timer.Interval = 1000
        Select Case Clickcolor
    Case 0:
        Current.FillColor = RGB(10, 136, 239)
    Case 1:
        Current.FillColor = RGB(255, 0, 0)
    Case 2:
        Current.FillColor = RGB(0, 255, 0)
    Case 3:
        Current.FillColor = RGB(0, 0, 255)
    Case 4:
        Current.FillColor = RGB(255, 255, 255)
    Case 5:
        Current.FillColor = RGB(35, 35, 35)
    End Select
    End If
    Start.Enabled = False
    NewG.Enabled = True
End Sub

'A timer... that times :)
Private Sub Timer_Timer()
    timing = timing + 1
    Time.Caption = "Time: " & timing
End Sub

'I HOPE THIS WAS A GOOD HELP FOR A BEGINNER GUY!!
'Please VOTE for my thing at http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=40630&lngWId=1
'=)

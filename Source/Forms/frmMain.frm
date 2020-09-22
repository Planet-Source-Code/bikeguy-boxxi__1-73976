VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   9210
   ScaleWidth      =   9915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNoMoves 
      BackColor       =   &H00FFFFC0&
      Caption         =   "No More Moves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   2280
      Width           =   9615
   End
   Begin VB.CommandButton cmdNextLevel 
      BackColor       =   &H0000FFFF&
      Caption         =   "Next Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   3000
      Width           =   9615
   End
   Begin VB.TextBox txtBlocks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   98
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtBlocks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   97
      Text            =   "0"
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtBlocks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   96
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtBlocks 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   95
      Text            =   "0"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2400
      TabIndex        =   94
      Text            =   "0"
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox picDead 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   3720
      Width           =   9615
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   90
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   92
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   89
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   91
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   88
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   90
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   87
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   89
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   86
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   88
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   85
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   87
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   84
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   86
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   83
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   85
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   82
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   84
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   81
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   83
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   80
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   82
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   79
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   81
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   78
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   80
         Top             =   4440
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   77
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   79
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   76
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   78
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   75
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   77
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   74
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   76
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   73
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   75
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   72
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   74
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   71
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   73
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   70
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   72
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   69
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   71
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   68
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   70
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   67
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   69
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   66
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   68
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   65
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   67
         Top             =   3720
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   64
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   66
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   63
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   65
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   62
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   64
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   61
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   63
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   60
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   62
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   59
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   61
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   58
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   60
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   57
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   59
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   56
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   58
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   55
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   57
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   54
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   56
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   53
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   55
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   52
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   54
         Top             =   3000
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   51
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   53
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   50
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   52
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   49
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   51
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   48
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   50
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   47
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   49
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   46
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   48
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   45
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   47
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   44
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   46
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   43
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   45
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   42
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   44
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   41
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   43
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   40
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   42
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   39
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   41
         Top             =   2280
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   38
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   40
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   37
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   39
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   36
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   38
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   35
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   37
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   34
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   36
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   33
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   35
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   32
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   34
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   31
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   33
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   30
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   32
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   29
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   31
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   28
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   30
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   27
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   29
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   26
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   28
         Top             =   1560
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   25
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   27
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   24
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   26
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   23
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   25
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   22
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   24
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   21
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   23
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   20
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   22
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   19
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   21
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   18
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   20
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   17
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   19
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   16
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   18
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   15
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   17
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   14
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   16
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   13
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   15
         Top             =   840
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   12
         Left            =   8760
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   14
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   11
         Left            =   8040
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   13
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   10
         Left            =   7320
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   12
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   9
         Left            =   6600
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   11
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   8
         Left            =   5880
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   10
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   7
         Left            =   5160
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   9
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   6
         Left            =   4440
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   8
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   5
         Left            =   3720
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   7
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   4
         Left            =   3000
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   6
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   3
         Left            =   2280
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   5
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   2
         Left            =   1560
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   4
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   1
         Left            =   840
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   3
         Top             =   120
         Width           =   705
      End
      Begin VB.PictureBox blockPics 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   720
         Index           =   0
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   705
         TabIndex        =   2
         Top             =   120
         Width           =   705
      End
   End
   Begin MSComctlLib.ImageList blocks 
      Left            =   960
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   43
      ImageHeight     =   44
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0779
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E8A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label labHighestScore 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   107
      Top             =   1440
      Width           =   6735
   End
   Begin VB.Label labSelectionPts 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   106
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label labCurrentScore 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   105
      Top             =   240
      Width           =   6615
   End
   Begin VB.Label labColors 
      Caption         =   "Yellow:"
      Height          =   255
      Index           =   4
      Left            =   1680
      TabIndex        =   104
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label labColors 
      Caption         =   "Red:"
      Height          =   255
      Index           =   3
      Left            =   1680
      TabIndex        =   103
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label labColors 
      Caption         =   "Green:"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   102
      Top             =   960
      Width           =   615
   End
   Begin VB.Label labColors 
      Caption         =   "Blue:"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   101
      Top             =   600
      Width           =   615
   End
   Begin VB.Label labLevel 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   100
      Top             =   360
      Width           =   615
   End
   Begin VB.Shape shpLevel 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   120
      Shape           =   3  'Circle
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label labTot 
      Caption         =   "Total Blocks:"
      Height          =   255
      Left            =   1320
      TabIndex        =   93
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mView 
      Caption         =   "View"
      Begin VB.Menu mViewHighScore 
         Caption         =   "High Score"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub blockPics_Click(Index As Integer)
    If blockData(Index).status <> BLOCK_DEAD Then
        If blockData(Index).status = BLOCK_ON Then
            ' The selected block is on, and the user is indicating they want to accept
            ' this move
            
            removeBlocks
            moveBlocksDown
            moveBlocksLeft
            runInfo.curPoints = runInfo.curPoints + runInfo.selPoints
            runInfo.selPoints = 0
            If runInfo.curPoints >= runInfo.nextLevelPoints Then
                Me.labLevel.BackColor = &HFF00&
                Me.shpLevel.FillColor = &HFF00&
                Me.cmdNextLevel.BackColor = &HFF00&
                
            End If
            
                
            
            
            
        Else
            ' The selected block is off, so make sure all blocks are turned off
            turnBlocksOff
            
            ' Now, turn on the blocks associated to the current block
            
            doCheck Index, BLOCK_ON
            
            
        End If
        
        ' Redraw the blocks
        drawMap
    End If
    refreshScores
    
    
    
    
    
End Sub


Private Sub refreshScores()
    Dim maxColor As Integer
    Dim maxNdx As Integer
    Dim ctr As Integer
    
    updStats
    ' Determine which color of block is predominant
    
    maxColor = -1
    For ctr = 1 To 4
        If runInfo.blockCount(ctr) > maxColor Then
            maxNdx = ctr
            maxColor = runInfo.blockCount(ctr)
        End If
        
        labColors(ctr).BorderStyle = 0
        labColors(ctr).BackColor = vbWhite
        
    Next
    labColors(maxNdx).BorderStyle = 1
    labColors(maxNdx).BackColor = &HFF00&
        
    
    For ctr = 1 To 4
        Me.txtBlocks(ctr).Text = Format(runInfo.blockCount(ctr))
    Next
    
    Me.txtTotal.Text = Format(runInfo.totalBlocks)
    
    Me.labSelectionPts.Caption = "Current selection is worth " & Format(runInfo.selPoints, "##,##0") & " points"
    Me.labCurrentScore.Caption = "Your current score is " & Format(runInfo.curPoints, "##,##0") & " points"
    If runInfo.curPoints > runInfo.highScore Then
        Me.labHighestScore.Caption = "Highest score ever is " & Format(runInfo.curPoints, "##,##0") & " points"
    Else
        Me.labHighestScore.Caption = "Highest score ever is " & Format(runInfo.highScore, "##,##0") & " points"
    End If
    Me.cmdNoMoves.Caption = "No More Moves (Will cost " & Format(runInfo.nmmPoints) & " points)"
    
    
End Sub

Private Sub cmdNextLevel_Click()
    If runInfo.curPoints >= runInfo.nextLevelPoints Then
    Else
        If runInfo.curPoints > runInfo.highScore Then
            runInfo.highScore = runInfo.curPoints
            putNewHighScore
        End If
        
        MsgBox "Sorry - not enough points...", vbOKOnly, "Game Over"
        
        initRunInfo
        
        
    End If
    buildMap
    refreshScores
    frmMain.labLevel.Caption = Format(runInfo.level)
    Me.labLevel.BackColor = &H80FFFF
    Me.shpLevel.FillColor = &H80FFFF
    Me.cmdNextLevel.BackColor = &H80FFFF
    drawMap
    
End Sub

Private Sub cmdNoMoves_Click()
    noMoreMoves
    refreshScores
    drawMap
    
End Sub

Private Sub Form_Load()
    runInfo.level = 0
    Me.labLevel.BackColor = &H80FFFF
    Me.shpLevel.FillColor = &H80FFFF
    Me.cmdNextLevel.BackColor = &H80FFFF
    Me.Icon = LoadPicture(App.Path & "\Face03.ico")
    Me.cmdNoMoves.Caption = "No More Moves (Will cost " & Format(runInfo.nmmPoints) & " points"
    
    
    
    
    
    
    'Randomize Timer
    
    initRunInfo
        
    
    ' Build the blocks
    
    buildMap
    refreshScores
    
    
    
    ' Draw the blocks
    
    drawMap
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    End
    
End Sub

Private Sub labHighestScore_Click()
    frmHighScore.Show vbModal, Me
    Me.labHighestScore.Caption = "Highest score ever is " & Format(runInfo.highScore, "##,##0") & " points"
    
End Sub

Private Sub mAbout_Click()
    frmAbout.Show vbModal, Me
    
End Sub

Private Sub mExit_Click()
    Unload Me
    End
    
End Sub

Private Sub mViewHighScore_Click()
    frmHighScore.Show vbModal, Me
    Me.labHighestScore.Caption = "Highest score ever is " & Format(runInfo.highScore, "##,##0") & " points"

End Sub

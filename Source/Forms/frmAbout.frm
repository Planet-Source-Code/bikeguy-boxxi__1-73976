VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmMovePic 
      Interval        =   1000
      Left            =   840
      Top             =   2880
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00FF8080&
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   810
      TabIndex        =   2
      Top             =   1125
      Width           =   4725
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FF8080&
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   810
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FF8080&
      Caption         =   "Version"
      Height          =   225
      Left            =   810
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Dim iLine As String
    
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Me.picIcon.Picture = LoadPicture(App.Path & "\face03.ico")
    iLine = "A pretty darn spectacular game of Boxxi..." & Chr(13) & Chr(13)
    iLine = iLine & "Highest score ever is " & Format(runInfo.highScore, "##,##0") & " made by " & runInfo.highScoreWho & " on " & runInfo.highScoreDate
    
    
    Me.lblDescription.Caption = iLine
    
    
    
End Sub

Private Sub tmMovePic_Timer()
    Dim maxLeft As Long
    Dim maxtop As Long
    
    maxLeft = Me.ScaleWidth - 300
    maxtop = Me.ScaleHeight - 300
    
    Me.picIcon.Left = Int(Rnd * maxLeft)
    Me.picIcon.Top = Int(Rnd * maxtop)
    
    
    
End Sub

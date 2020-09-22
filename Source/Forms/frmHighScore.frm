VERSION 5.00
Begin VB.Form frmHighScore 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Boxxi High Score"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label labHighScore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    Me.labHighScore.Caption = "Highest score ever is " & Format(runInfo.highScore, "##,##0") & " points, attained by " & runInfo.highScoreWho & " on " & Format(runInfo.highScoreDate) & Chr(13) & "Click to reset"
    
End Sub

Private Sub labHighScore_Click()
    If MsgBox("Reset high score - Are you sure?", vbYesNo, "Confirm Reset?") = vbYes Then
        runInfo.highScore = 0
        runInfo.highScoreDate = "Never"
        runInfo.highScoreWho = "Nobody"
        Me.labHighScore.Caption = "Highest score ever is " & Format(runInfo.highScore) & " points, attained by " & runInfo.highScoreWho & " on " & Format(runInfo.highScoreDate) & Chr(13) & "(Click to reset)"
        
        
        If FileExists(App.Path & "\HighScore.txt") Then
            Kill App.Path & "\HighScore.txt"
        End If
        
        
        
    End If

End Sub

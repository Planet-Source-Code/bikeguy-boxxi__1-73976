Attribute VB_Name = "modUtils"
'---------------------------------------------------------------------------------------
' Module    : modUtils
' Author    : JW
' Date      : 6/22/2011
' Purpose   : General utilities
'---------------------------------------------------------------------------------------
Public Sub putNewHighScore()
    Dim iHand As Integer
    Dim iFile As String
    Dim iLine As String
    
    Dim iWho As String
    
    iWho = "Enter your name"
    iWho = InputBox("Enter your name:", "New High Score!", iWho)
    
    If iWho = "Enter your name" Then
        iWho = "N/A"
    End If
    
    iLine = Format(runInfo.highScore) & "#" & Format(Date, "short date") & "#" & iWho
    
    iFile = App.Path & "\HighScore.txt"
    
    
    iHand = FreeFile()
    Open iFile For Output As #iHand
    Print #iHand, iLine
    
    Close #iHand
    
        

End Sub
Public Sub initRunInfo()
    Dim iHand As Integer
    Dim iFile As String
    Dim iData() As String
    Dim ctr As Integer
    
    runInfo.highScore = 0
    iFile = App.Path & "\HighScore.txt"
    If FileExists(iFile) Then
        iHand = FreeFile()
        Open iFile For Input As #iHand
        Line Input #iHand, iLine
        iData = Split(iLine, "#")
        
        
        runInfo.highScore = Val(iData(0))
        runInfo.highScoreDate = (iData(1))
        runInfo.highScoreWho = (iData(2))
        Close #iHand
    Else
        runInfo.highScore = 0
        runInfo.highScoreDate = "Never"
        runInfo.highScoreWho = "Nobody"
        
    End If
        
    For ctr = 1 To 4
        runInfo.blockCount(ctr) = 0
    Next
    
    runInfo.level = 0
    runInfo.curPoints = 0
    Randomize Timer
    
    
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FileExists
' Author    : JW
' Date      : 6/22/2011
' Purpose   : Returns true if the filename sent in as a parameters exists, else false.
'---------------------------------------------------------------------------------------
'
Public Function FileExists(FullFileName As String) As Boolean
    Dim iHand As Integer
    On Error GoTo MakeF
    iHand = FreeFile()
        'If file does Not exist, there will be an Error
        Open FullFileName For Input As #iHand
        Close #iHand
        'no error, file exists
        FileExists = True
    Exit Function
MakeF:
        'error, file does Not exist
        FileExists = False
    Exit Function
End Function

Public Function RandomNumber(ByVal MaxValue As Long, Optional _
ByVal MinValue As Long = 0)

  On Error Resume Next
  
  RandomNumber = Int((MaxValue - MinValue + 1) * Rnd) + MinValue

End Function


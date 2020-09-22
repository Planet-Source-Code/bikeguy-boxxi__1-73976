Attribute VB_Name = "modMain"
Dim ndxList() As Integer        ' Used to determine linked blocks upon click
Dim numNdxList As Integer

Public Sub main()
    frmMain.Show
    
End Sub
Private Function addNdxList(iNdx As Integer) As Boolean
    Dim ctr As Integer
    Dim rVal As Boolean
    rVal = True
    For ctr = 0 To numNdxList - 1
        If ndxList(ctr) = iNdx Then
            rVal = False
            Exit For
        End If
    Next
    If rVal Then
        numNdxList = numNdxList + 1
        ReDim Preserve ndxList(numNdxList)
        ndxList(numNdxList - 1) = iNdx
    End If
    
    addNdxList = rVal
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : turnBlocksOff
' Author    : nie1jlw
' Date      : 6/21/2011
' Purpose   : Turns all the blocks off
'---------------------------------------------------------------------------------------
'
Public Sub turnBlocksOff()
    Dim ctr As Long
    For ctr = 0 To 90
        If blockData(ctr).status = BLOCK_ON Then
            blockData(ctr).status = BLOCK_OFF
        End If
    Next

End Sub
'---------------------------------------------------------------------------------------
' Procedure : removeBlocks
' Author    : nie1jlw
' Date      : 6/21/2011
' Purpose   : Blocks that are on at this point need to be killed.  This is called after
'             the player has clicked on a set to turn them on.
'---------------------------------------------------------------------------------------
'
Public Sub removeBlocks()
    Dim ctr As Long
    For ctr = 0 To 90
        If blockData(ctr).status = BLOCK_ON Then
            blockData(ctr).status = BLOCK_DEAD
        End If
    Next
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : doCheck
' Author    : nie1jlw
' Date      : 6/21/2011
' Purpose   : This procedure checks for linked groups of blocks
'             iNdx is the index of the block clicked on by the user
'             iStat is the new status to change the group of blocks too
'---------------------------------------------------------------------------------------
'
Public Sub doCheck(iNdx As Integer, iStat As Long)
    Dim nAdded As Integer
    Dim ctr As Integer
    Dim iAdded As Boolean
    Dim ndx As Integer
    
    
    ' Reset the ndxlist of blocks to be turned on
    Erase ndxList
    numNdxList = 0
    
    ' Add the index of the clicked on box to the ndxList
    iAdded = addNdxList(iNdx)
    
    
    
    ' Iteratively move through the list, checking blocks that are in the set to determine
    ' if they have neighbors which need to be added.
    nAdded = 1
    Do While nAdded > 0
        nAdded = 0
        For ctr = 0 To numNdxList - 1
            ndx = ndxList(ctr)
            
        
            If blockData(ndx).status <> BLOCK_DEAD Then
                iNdx = checkAbove(ndx)
                If iNdx >= 0 Then

                    
                    If addNdxList(iNdx) Then

                    
                        nAdded = nAdded + 1
                        
                    End If
                End If
                    
                iNdx = checkBelow(ndx)
                If iNdx >= 0 Then
                    If addNdxList(iNdx) Then

                    
                        nAdded = nAdded + 1
                        
                    End If
                End If
                iNdx = checkright(ndx)
                If iNdx >= 0 Then
                    If addNdxList(iNdx) Then

                    
                        nAdded = nAdded + 1
                        
                    End If
                End If
                iNdx = checkLeft(ndx)
                If iNdx >= 0 Then
                    If addNdxList(iNdx) Then

                    
                        nAdded = nAdded + 1
                        
                    End If
                End If
            End If
        Next
    Loop
    ' You have to have a least two blocks in the set to do anything.
    ' If we have at least two blocks in the set, change their status to the
    ' status parameter
    If numNdxList >= 2 Then
        For ctr = 0 To numNdxList - 1
            blockData(ndxList(ctr)).status = iStat
        Next
    End If
    
                        
                    
                    
        

End Sub
'---------------------------------------------------------------------------------------
' Procedure : moveBlocksLeft
' Author    : nie1jlw
' Date      : 6/21/2011
' Purpose   : Checks for dead columns, and if it finds one, it moves everything to the left.
'---------------------------------------------------------------------------------------
'
Public Sub moveBlocksLeft()
    Dim cCtr As Long
    Dim ctr As Long
    Dim colOk As Boolean
    Dim killCol As Long
    Dim numMoved As Integer
    numMoved = 1
    Do While numMoved > 0
        numMoved = 0
    
        killCol = -1
        For cCtr = 0 To 12  ' Column
            colOk = False
            
            For ctr = 0 To 90
                If blockData(ctr).status <> BLOCK_DEAD Then
                    If blockData(ctr).colVal = cCtr Then
                        colOk = True
                        Exit For
                    End If
                End If
            Next
            If Not colOk Then
                killCol = cCtr
                Exit For
            End If
            
        Next
        If killCol >= 0 Then
            For ctr = 0 To 90
                If blockData(ctr).colVal >= killCol Then
                    If blockData(ctr).leftNdx >= 0 Then
                        If blockData(blockData(ctr).leftNdx).status = BLOCK_DEAD Then
                            blockData(blockData(ctr).leftNdx).status = BLOCK_OFF
                            blockData(blockData(ctr).leftNdx).value = blockData(ctr).value
                            blockData(ctr).status = BLOCK_DEAD
                            numMoved = numMoved + 1
                        End If
                    End If
                End If
            Next
            
            
        End If
    Loop
End Sub
Public Function getNewLeftNdx(iNdx As Integer) As Integer
    Dim goodNdx As Boolean
    Dim rNdx As Long
    rNdx = -1
    goodNdx = True
    Do While goodNdx
        If blockData(iNdx).status = BLOCK_DEAD Then
            goodNdx = False
            rNdx = iNdx
        Else
            If blockData(iNdx).leftNdx >= 0 Then
                iNdx = blockData(iNdx).leftNdx
            Else
                goodNdx = False
            End If
        End If
    Loop
    getNewLeftNdx = rNdx
    
            
End Function
'---------------------------------------------------------------------------------------
' Procedure : moveBlocksDown
' Author    : nie1jlw
' Date      : 6/21/2011
' Purpose   : After killing blocks, this procedure moves blocks down where appropriate
'---------------------------------------------------------------------------------------
'
Public Sub moveBlocksDown()
    Dim nMoved As Long
    Dim ctr As Long
    
    nMoved = 1
    Do While nMoved > 0
        nMoved = 0
        For ctr = 0 To 90
            If blockData(ctr).belowNdx >= 0 And blockData(ctr).status <> BLOCK_DEAD Then
                If blockData(blockData(ctr).belowNdx).status = BLOCK_DEAD Then
                
                    blockData(blockData(ctr).belowNdx).status = blockData(ctr).status
                    blockData(blockData(ctr).belowNdx).value = blockData(ctr).value
                    
                    blockData(ctr).status = BLOCK_DEAD
                    nMoved = nMoved + 1
                End If
            End If
        Next
    Loop
    
                    
            
        
End Sub
Private Function checkBelow(ndx As Integer) As Integer
    Dim rVal As Integer
    
    
    rVal = -1
    If blockData(ndx).belowNdx >= 0 Then
        If blockData(blockData(ndx).belowNdx).status <> BLOCK_DEAD Then
            If blockData(ndx).value = blockData(blockData(ndx).belowNdx).value Then
                rVal = blockData(ndx).belowNdx
            End If
        End If
    
    End If
    
    checkBelow = rVal
    
            
End Function
Private Function checkright(ndx As Integer) As Integer
    Dim rVal As Integer
    
    
    
    rVal = -1
    If blockData(ndx).rightNdx >= 0 Then
        If blockData(blockData(ndx).rightNdx).status <> BLOCK_DEAD Then
            If blockData(ndx).value = blockData(blockData(ndx).rightNdx).value Then
                rVal = blockData(ndx).rightNdx
            End If
        End If
    End If
    checkright = rVal
    
            
End Function

Private Function checkLeft(ndx As Integer) As Integer
    Dim rVal As Integer
    
    
    
    rVal = -1
    If blockData(ndx).leftNdx >= 0 Then
        If blockData(blockData(ndx).leftNdx).status <> BLOCK_DEAD Then
            If blockData(ndx).value = blockData(blockData(ndx).leftNdx).value Then
                rVal = blockData(ndx).leftNdx
        
            End If
        End If
    End If
    checkLeft = rVal
    
            
End Function

Private Function checkAbove(ndx As Integer) As Integer
    Dim rVal As Integer
    
    
    
    rVal = -1
    If blockData(ndx).aboveNdx >= 0 Then
        If blockData(blockData(ndx).aboveNdx).status <> BLOCK_DEAD Then
            If blockData(ndx).value = blockData(blockData(ndx).aboveNdx).value Then
                rVal = blockData(ndx).aboveNdx
                
                
            End If
        End If
        
    End If
    
    checkAbove = rVal
    
            
End Function

Public Sub buildMap()
    Dim rCtr As Long
    Dim cCtr As Long
    Dim rndVal As Integer
    
    runInfo.blueBlocks = 0
    runInfo.redBlocks = 0
    runInfo.greenBlocks = 0
    runInfo.yellowBlocks = 0
    'Randomize 14
    
    For rCtr = 0 To 6   ' Row
        For cCtr = 0 To 12  ' Column
            numblockData = numblockData + 1
            ReDim Preserve blockData(numblockData)
            blockData(numblockData - 1).status = BLOCK_OFF
            blockData(numblockData - 1).rowVal = rCtr
            blockData(numblockData - 1).colVal = cCtr
            
            blockData(numblockData - 1).aboveNdx = calcAboveNdx(rCtr, numblockData)
            blockData(numblockData - 1).belowNdx = calcBelowNdx(rCtr, numblockData)
            blockData(numblockData - 1).leftNdx = calcLeftNdx(cCtr, numblockData)
            blockData(numblockData - 1).rightNdx = calcRightNdx(cCtr, numblockData)
            
            rndVal = RandomNumber(0, 5)
            blockData(numblockData - 1).value = rndVal
            Select Case rndVal
                Case Is = 1
                    
                    runInfo.blueBlocks = runInfo.blueBlocks + 1
                Case Is = 2
                    
                    runInfo.greenBlocks = runInfo.greenBlocks + 1
                Case Is = 3
                    
                    runInfo.redBlocks = runInfo.redBlocks + 1
                Case Is = 4
                    
                    runInfo.yellowBlocks = runInfo.yellowBlocks + 1
            End Select
        Next
    Next
    runInfo.totalBlocks = runInfo.redBlocks + runInfo.blueBlocks + runInfo.greenBlocks + runInfo.yellowBlocks
    runInfo.level = 1

End Sub
Public Function calcAboveNdx(rowNdx As Long, itemNdx As Long) As Integer
    calcAboveNdx = -1
    If rowNdx >= 1 Then
        calcAboveNdx = itemNdx - 14
    End If
End Function
Public Function calcBelowNdx(rowNdx As Long, itemNdx As Long) As Integer
    calcBelowNdx = -1
    If rowNdx < 6 Then
        calcBelowNdx = itemNdx + 12
    End If

End Function
Public Function calcLeftNdx(colNdx As Long, itemNdx As Long) As Integer
    calcLeftNdx = -1
    If colNdx > 1 Then
        calcLeftNdx = itemNdx - 2
    End If

End Function
Public Function calcRightNdx(colNdx As Long, itemNdx As Long) As Integer
    calcRightNdx = -1
    If colNdx < 13 Then
        calcRightNdx = itemNdx
    End If

End Function

            

Public Sub drawmap()
    Dim ctr As Long
    Dim ndx As Long
    
    For ctr = 0 To 90
        ndx = -1
        
        Select Case blockData(ctr).status
        
            Case Is = BLOCK_DEAD
                frmMain.blockPics(ctr).Picture = frmMain.picDead
                ndx = -1
            Case Is = BLOCK_ON
                ndx = blockData(ctr).value + 4
            Case Is = BLOCK_OFF
                ndx = blockData(ctr).value
        End Select
        If ndx >= 0 Then
        
            Set frmMain.blockPics(ctr).Picture = frmMain.blocks.ListImages(ndx).Picture
        End If
    Next
    
End Sub

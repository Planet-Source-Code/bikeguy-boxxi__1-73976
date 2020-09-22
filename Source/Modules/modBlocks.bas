Attribute VB_Name = "modBlocks"
Dim ndxList() As Integer        ' Used to determine linked blocks upon click
Dim numNdxList As Integer
Public Sub updStats()
    Dim ctr As Long
    For ctr = 1 To 4
        runInfo.blockCount(ctr) = 0
    Next
    runInfo.totalBlocks = 0
    For ctr = 0 To 90
        If blockData(ctr).status <> BLOCK_DEAD Then
            runInfo.totalBlocks = runInfo.totalBlocks + 1
            runInfo.blockCount(blockData(ctr).value) = runInfo.blockCount(blockData(ctr).value) + 1
        End If
    Next
            
End Sub
Private Function addNdxList(iNdx As Integer) As Integer
    Dim ctr As Integer
    Dim rVal As Integer
    rVal = 1
    For ctr = 0 To numNdxList - 1
        If ndxList(ctr) = iNdx Then
            rVal = 0
            Exit For
        End If
    Next
    If rVal = 1 Then
        numNdxList = numNdxList + 1
        ReDim Preserve ndxList(numNdxList)
        ndxList(numNdxList - 1) = iNdx
    End If
    
    addNdxList = rVal
    
End Function
Public Sub noMoreMoves()
    Dim ctr As Long
    Dim rndVal As Integer
    For ctr = 1 To 4
        runInfo.blockCount(ctr) = 0
    Next
    runInfo.totalBlocks = 0
    
    runInfo.curPoints = runInfo.curPoints - runInfo.nmmPoints
    runInfo.nmmPoints = runInfo.nmmPoints * 2
    For ctr = 0 To 90
        If blockData(ctr).status <> BLOCK_DEAD Then
            rndVal = RandomNumber(0, 5)
            blockData(ctr).value = rndVal
            runInfo.blockCount(rndVal) = runInfo.blockCount(rndVal) + 1
            runInfo.totalBlocks = runInfo.totalBlocks + 1
        End If
    Next
    
            
End Sub
'---------------------------------------------------------------------------------------
' Procedure : turnBlocksOff
' Author    : JW
' Date      : 6/21/2011
' Purpose   : Turns all the blocks off
'---------------------------------------------------------------------------------------
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
' Author    : JW
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
' Author    : JW
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
    nAdded = addNdxList(iNdx)
    
    
    
    ' Iteratively move through the list, checking blocks that are in the set to determine
    ' if they have neighbors which need to be added.
    nAdded = 1
    Do While nAdded > 0
        nAdded = 0
        For ctr = 0 To numNdxList - 1
            ndx = ndxList(ctr)
            
        
            If blockData(ndx).status <> BLOCK_DEAD Then
                iNdx = checkDir(ndx, blockData(ndx).aboveNdx)
                

                If iNdx >= 0 Then

                    
                    nAdded = nAdded + addNdxList(iNdx)

                    
                        
                        
                    
                End If
                iNdx = checkDir(ndx, blockData(ndx).belowNdx)
                

                If iNdx >= 0 Then
                    nAdded = nAdded + addNdxList(iNdx)
                End If
                iNdx = checkDir(ndx, blockData(ndx).rightNdx)

                If iNdx >= 0 Then
                    nAdded = nAdded + addNdxList(iNdx)
                End If
                iNdx = checkDir(ndx, blockData(ndx).leftNdx)

                If iNdx >= 0 Then
                    nAdded = nAdded + addNdxList(iNdx)
                End If
            End If
        Next
    Loop
    ' You have to have a least two blocks in the set to do anything.
    ' If we have at least two blocks in the set, change their status to the
    ' status parameter
    runInfo.selPoints = 0
    If numNdxList >= 2 Then
        ' This sets the score of the potential selection to display to the user.
    
        runInfo.selPoints = numNdxList ^ 2 * 500
        For ctr = 0 To numNdxList - 1
            blockData(ndxList(ctr)).status = iStat
        Next
    End If
    
                        
                    
                    
        

End Sub
'---------------------------------------------------------------------------------------
' Procedure : moveBlocksLeft
' Author    : JW
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
                If blockData(ctr).status <> BLOCK_DEAD Then
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
                End If
                
            Next
            
            
        End If
    Loop
End Sub
'---------------------------------------------------------------------------------------
' Procedure : moveBlocksDown
' Author    : JW
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
'---------------------------------------------------------------------------------------
' Procedure : checkDir
' Author    : nie1jlw
' Date      : 6/27/2011
' Purpose   : The current block index is sent in, along with the index to be checked.
'             If the block to be check is not dead, and the values of the two blocks match,
'             then return the dirndx, else, return -1.
'---------------------------------------------------------------------------------------
'
Private Function checkDir(ndx As Integer, dirNdx As Integer) As Integer
    Dim rVal As Integer
    
    
    rVal = -1
    If dirNdx >= 0 Then
        If blockData(dirNdx).status <> BLOCK_DEAD Then
            If blockData(ndx).value = blockData(dirNdx).value Then
                rVal = dirNdx
            End If
        End If
    End If
    
    
    checkDir = rVal

End Function

Public Sub buildMap()
    Dim rCtr As Long
    Dim cCtr As Long
    Dim rndVal As Integer
    runInfo.nmmPoints = 10000
    
    'Randomize 14
    Erase blockData
    numBlockData = 0
    
    For rCtr = 0 To 6   ' Row
        For cCtr = 0 To 12  ' Column
            
            numBlockData = numBlockData + 1
            ReDim Preserve blockData(numBlockData)
            blockData(numBlockData - 1).status = BLOCK_OFF
            blockData(numBlockData - 1).rowVal = rCtr
            blockData(numBlockData - 1).colVal = cCtr
            ' Given where the current block is located, calculate the directional indexes
            ' For example, the bottom row will have a below index of -1, and the
            ' top row will have an above index of -1.
            
            blockData(numBlockData - 1).aboveNdx = calcAboveNdx(rCtr, numBlockData)
            blockData(numBlockData - 1).belowNdx = calcBelowNdx(rCtr, numBlockData)
            blockData(numBlockData - 1).leftNdx = calcLeftNdx(cCtr, numBlockData)
            blockData(numBlockData - 1).rightNdx = calcRightNdx(cCtr, numBlockData)
            blockNdx(rCtr, cCtr) = numBlockData - 1
            
            rndVal = RandomNumber(0, 5)
            blockData(numBlockData - 1).value = rndVal
            runInfo.blockCount(rndVal) = runInfo.blockCount(rndVal) + 1
            runInfo.totalBlocks = runInfo.totalBlocks + 1
        Next
    Next

    runInfo.level = runInfo.level + 1
    
    If runInfo.level < 3 Then
        runInfo.nextLevelPoints = ((runInfo.level + 1) ^ 2 * 50000)
    ElseIf runInfo.level > 3 And runInfo.level < 5 Then
        runInfo.nextLevelPoints = ((runInfo.level + 1) ^ 2 * 40000)
    Else
        runInfo.nextLevelPoints = ((runInfo.level + 1) ^ 2 * 35000)
    End If
    frmMain.Caption = "Boxxi -> Points Needed for next level: " & Format(runInfo.nextLevelPoints)


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
    If colNdx > 0 Then
        calcLeftNdx = itemNdx - 2
    End If

End Function
Public Function calcRightNdx(colNdx As Long, itemNdx As Long) As Integer
    calcRightNdx = -1
    If colNdx < 12 Then
        calcRightNdx = itemNdx
    End If

End Function
Public Sub drawMap()
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


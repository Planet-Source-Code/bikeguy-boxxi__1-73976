Attribute VB_Name = "modGlobals"
' This structure holds all of the general scoring items for the system.

Public Type scoreType
    blockCount(1 To 4) As Integer  ' Holds the count of different block colors
    
    totalBlocks As Integer
    level As Long               ' Game level
    
    highScore As Double         ' Highest score ever attained
    highScoreDate As String     ' Date of highest score ever
    highScoreWho As String      ' Name of highest scorer
    
    nmmPoints As Double         ' Cost of pressing the no more moves button

    selPoints As Double         ' Point value of current selection
    curPoints As Double         ' Current game points
    nextLevelPoints As Double   ' Points required to get to next level
End Type
Public runInfo As scoreType

' Block statuses
Public Const BLOCK_OFF = 0
Public Const BLOCK_ON = 1
Public Const BLOCK_DEAD = 2

Public Type blockType
    colVal As Integer       ' Column of this block
    rowVal As Integer       ' Row of this block
    status As Integer       ' 1 = on, 0 = off
    value As Integer        ' BLUE_LOW through YELLOW_HIGH [See constants below]
    leftNdx As Integer      ' Index of block to the left
    rightNdx As Integer     ' Index of the block to the right
    aboveNdx As Integer     ' Index of the block above
    belowNdx As Integer     ' Index of the block below
    
End Type
Public blockData() As blockType
Public numBlockData As Long
Public blockNdx(7, 13) As Long

' These constants are indexes into the image list on frmMain.  The current setting for
' a block is stored in the "value" property
Public Const BLUE_LOW = 1
Public Const GREEN_LOW = 2
Public Const RED_LOW = 3
Public Const YELLOW_LOW = 4
Public Const BLUE_HIGH = 5
Public Const GREEN_HIGH = 6
Public Const RED_HIGH = 7
Public Const YELLOW_HIGH = 8




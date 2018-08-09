Public pBoardRangeRow As Integer
Public pBoardRangeCol As Integer
Public BoardRangeRow As Integer
Public BoardRangeCol As Integer
Public BoardRange As Range
Public BoardRangeBorder As Range

Private Sub Class_Initialize()
    pBoardRangeRow = 0
    pBoardRangeCol = 0
End Sub

Public Sub InitializeWithValues(ByVal BoardRangeRow As Integer, ByVal BoardRangeCol As Integer)
    pBoardRangeRow = BoardRangeRow
    pBoardRangeCol = BoardRangeCol
End Sub

Public Function BoardGameRange()
    Set BoardRangeBorder = Range(Cells(1, 1), Cells(pBoardRangeRow, pBoardRangeCol))
    Set BoardRange = Range(Cells(2, 2), Cells(pBoardRangeRow - 1, pBoardRangeCol - 1))
End Function
        
Public Sub GenerateGameBoard()
    Cells.ClearFormats
    Call BoardGameRange
    BoardRangeBorder.Interior.ColorIndex = 1
    BoardRange.Interior.ColorIndex = 2
End Sub

Public Function RandomRow() As Integer
    RandomRow = Int(((pBoardRangeRow) - 3 + 1) * Rnd + 1)
End Function

Public Function RandomCol() As Integer
    RandomCol = Int(((pBoardRangeCol) - 3 + 1) * Rnd + 2)
End Function

Dim TimerActive As Boolean
Sub StartTimer()
    Start_Timer
End Sub
Private Sub Stop_Timer()
    TimerActive = False
End Sub

Private Sub Timer()
    If TimerActive Then
        ActiveSheet.Cells(1, 1).Value = Time
        Application.OnTime Now() + TimeValue("00:00:03"), "Timer"
    End If
End Sub

Function RunEveryOneMinutes()
Dim ActiveCellRange As Range
Set ActiveCellRange = ActiveCell
Set RunEveryOneMinutes = ActiveCellRange
End Function

Sub RunEveryTwoMinutes()
Call RunEveryOneMinutes
Dim CellRand As Range
Dim MyClass As New BoardAssets
Application.ScreenUpdating = False
MyClass.InitializeWithValues 30, 20
MyClass.GenerateGameBoard

Set CellRand = MyClass.BoardRange.EntireRow.Cells(MyClass.RandomRow, MyClass.RandomCol)

With ActiveWorkbook.ActiveSheet
    For Each Cell In CellRand
        Set ActiveCellRange = RunEveryOneMinutes
        Cell.Interior.ColorIndex = 3
        ActiveCellRange.Interior.ColorIndex = 4
        'Application.Wait (Now + TimeValue("00:00:02"))
        If ActiveCellRange.Address = CellRand.Address Then
            MsgBox ("You caught it!")
            Exit Sub
        End If
        Application.OnTime Now + TimeValue("00:00:01"), "RunEveryTwoMinutes"
    Next Cell
End With
'To add a score keeper and a time limit
End Sub



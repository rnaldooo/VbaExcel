Sub CreatePermutation()

Dim FirstCell As Integer
Dim SecondCell As Integer
Dim NumRows As Integer
Dim OutputRow As Long

    ' Get the total number of rows in column A
    NumRows = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row()

    ' You want to start outputting in row 1
    OutputRow = 1

    For FirstCell = 1 To NumRows - 1
        For SecondCell = FirstCell + 1 To NumRows

            ' Put in the data from the first cell into columnc C & D
            Cells(OutputRow, 3).Value = Cells(FirstCell, 1).Value
            Cells(OutputRow, 4).Value = Cells(FirstCell, 2).Value

            ' Put in the data from the second cell into column E & F
            Cells(OutputRow, 5).Value = Cells(SecondCell, 1).Value
            Cells(OutputRow, 6).Value = Cells(SecondCell, 2).Value

            ' Move to the next row to output
            OutputRow = OutputRow + 1

        Next SecondCell
    Next FirstCell
End Sub

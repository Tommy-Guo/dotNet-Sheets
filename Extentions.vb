Public Module Extensions
    <Runtime.CompilerServices.Extension>
    Public Function LocationFromCell(x As Integer, y As Integer, cellSize As Size) As Point
        Return New Point(cellSize.Width + cellSize.Width * x, cellSize.Height + cellSize.Height * y)
    End Function

    <Runtime.CompilerServices.Extension>
    Public Function LocationFromCell(cell As Cell, cellSize As Size) As Point
        Return New Point(cellSize.Width + (cell.Location.X * cellSize.Width), cellSize.Height + (cell.Location.Y * cellSize.Height))
    End Function

    <Runtime.CompilerServices.Extension>
    Public Function CellFromLocation(x As Integer, y As Integer, cells As Cell(,)) As Cell
        Return cells(x, y)
    End Function

    <Runtime.CompilerServices.Extension>
    Public Function CellFromLocation(location As Point, cells As Cell(,), cellSize As Size, gridSize As Size) As Cell
        For x As Integer = 0 To gridSize.Width - 1
            For y As Integer = 0 To gridSize.Height - 1
                If cells(x, y).Rectangle.Contains(location) Then
                    Return cells(x, y)
                End If
            Next
        Next
        Return Nothing
    End Function
End Module
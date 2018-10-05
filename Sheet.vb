Public Class Sheet
    Inherits Control

    Dim Running As Boolean

    Private _gridSize As New Size(10, 10)
    Public Property GridSize As Size
        Get
            Return _gridSize
        End Get
        Set(value As Size)
            _gridSize = value
            ReDim Cells(_gridSize.Width - 1, _gridSize.Height - 1)
            InitializeCells()
            Me.Location = New Point(0, 0)
            Me.Size = New Size(Cells(GridSize.Width - 1, 0).Rectangle.X + CellSize.Width + 1, Cells(0, GridSize.Height - 1).Rectangle.Y + CellSize.Height + 1)
            Invalidate()
        End Set
    End Property

    Private _cellSize As New Size(80, 30)
    Public Property CellSize As Size
        Get
            Return _cellSize
        End Get
        Set(value As Size)
            _cellSize = value
            Invalidate()
        End Set
    End Property

    Private _activeCell As Cell
    Public Property ActiveCell As Cell
        Get
            Return _activeCell
        End Get
        Set(value As Cell)
            _activeCell = value
        End Set
    End Property

    Public Cells(GridSize.Width - 1, GridSize.Height - 1) As Cell

    Private _formulaBox As TextBox
    Public Property FormulaBox As TextBox
        Get
            Return _formulaBox
        End Get
        Set(value As TextBox)
            _formulaBox = value
            RemoveHandler FormulaBox.KeyDown, AddressOf equation_KeyDown
            AddHandler FormulaBox.KeyDown, AddressOf equation_KeyDown
        End Set
    End Property

    Private _statLabel As ToolStripLabel
    Public Property StatusLabel As ToolStripLabel
        Get
            Return _statLabel
        End Get
        Set(value As ToolStripLabel)
            _statLabel = value
        End Set
    End Property


    Public FormulaMode As Boolean = False

    Public SelectedBox As TextBox = New TextBox() With {.Multiline = True, .Size = New Size(CellSize.Width - 1, CellSize.Height - 1), .BorderStyle = BorderStyle.None, .Visible = False}

    Sub New()
        InitializeCells()
        Controls.Add(SelectedBox)
        DoubleBuffered = True

        AddHandler SelectedBox.KeyDown, AddressOf cell_Keydown
    End Sub

    Private Sub InitializeCells()
        For x As Integer = 0 To GridSize.Width - 1
            For y As Integer = 0 To GridSize.Height - 1
                Cells(x, y) = New Cell() With {.Location = New Point(x, y), .Rectangle = New Rectangle(LocationFromCell(x, y, CellSize), CellSize), .Sheet = Me}
            Next
        Next
        Invalidate()
    End Sub 

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e) 
        Dim g As Graphics = e.Graphics : g.SmoothingMode = Drawing2D.SmoothingMode.HighSpeed : g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

        g.Clear(Color.White)

        '/ -- Draw column and row margins -- /'                                      
        g.FillRectangle(New SolidBrush(Color.FromArgb(242, 242, 242)), 0, 0, Width, CellSize.Height)
        g.FillRectangle(New SolidBrush(Color.FromArgb(242, 242, 242)), 0, 0, CellSize.Width, Height)

        If Not ActiveCell Is Nothing Then
            g.FillRectangle(New SolidBrush(Color.FromArgb(220, 220, 220)), New Rectangle(0, LocationFromCell(ActiveCell, CellSize).Y, CellSize.Width, CellSize.Height))
            g.FillRectangle(New SolidBrush(Color.FromArgb(220, 220, 220)), New Rectangle(LocationFromCell(ActiveCell, CellSize).X, 0, CellSize.Width, CellSize.Height))
        End If

        ' Draw margin text 
        For column As Integer = 1 To GridSize.Width
            g.DrawString(column.ToString, New Font("Verdana", 8.25), Brushes.Black, New Rectangle(column * CellSize.Width, 0, CellSize.Width + 1, CellSize.Height + 1), New StringFormat() With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
        Next

        For row As Integer = 1 To GridSize.Height
            g.DrawString(row.ToString, New Font("Verdana", 8.25), Brushes.Black, New Rectangle(0, row * CellSize.Height, CellSize.Width + 1, CellSize.Height + 1), New StringFormat() With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center})
        Next

        g.FillPolygon(New SolidBrush(Color.FromArgb(187, 187, 187)), New Point(2) {New Point(CellSize.Width - 3, CellSize.Height - 3), New Point(CellSize.Width - 3, CellSize.Height - 16), New Point(CellSize.Width - 16, CellSize.Height - 3)})

        '/ -- Draw basic grid -- /'
        For x As Integer = 0 To GridSize.Width + 1
            g.DrawLine(Pens.Silver, New Point(x * CellSize.Width, 0), New Point(x * CellSize.Width, Height))
        Next
        For y As Integer = 0 To GridSize.Height + 1
            g.DrawLine(Pens.Silver, New Point(0, y * CellSize.Height), New Point(Width, y * CellSize.Height))
        Next

        DrawCells(g)
        DrawSelection(g)

        If _statLabel IsNot Nothing Then
            Dim statusText As String = ""
            statusText &= String.Format("Grid Size: {0}   |   ", "{" & GridSize.Width & ", " & GridSize.Height & "}")
            If ActiveCell IsNot Nothing Then
                statusText &= "Selected Cell: {" & Me.ActiveCell.Location.X + 1 & ", " & Me.ActiveCell.Location.Y + 1 & "}"
            Else
                statusText &= "Selected Cell: None"
            End If

            StatusLabel.Text = statusText
        End If
    End Sub 

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If New Rectangle(CellSize.Width, CellSize.Height, Width - CellSize.Width, Height - CellSize.Height).Contains(e.Location) Then
            If FormulaMode Then
                Dim selectionS As Integer = FormulaBox.SelectionStart
                Dim insertCord As String = "{" & CellFromLocation(e.Location, Cells, CellSize, GridSize).Location.X + 1 & ", " & CellFromLocation(e.Location, Cells, CellSize, GridSize).Location.Y + 1 & "}"
                FormulaBox.Text = FormulaBox.Text.Insert(selectionS, insertCord)
                FormulaBox.SelectionStart = selectionS + insertCord.Length
            Else
                If Not SelectedBox.Text.StartsWith("=") Then
                    Try
                        cell_Keydown(SelectedBox, New KeyEventArgs(Keys.Enter)) 
                    Catch : End Try
                End If
                SelectCell(CellFromLocation(e.Location, Cells, CellSize, GridSize))
                FormulaBox.Text = ActiveCell.Equation.ToString
                Invalidate()
            End If
        End If
    End Sub

    Private Sub SelectCell(cell As Cell)
        ActiveCell = cell
        SelectedBox.Location = New Point(CellSize.Width * cell.Location.X + 1 + CellSize.Width, CellSize.Height * cell.Location.Y + 1 + CellSize.Height)
        SelectedBox.Text = cell.Text
        SelectedBox.ForeColor = cell.Forecolor
        SelectedBox.BackColor = cell.Fill
        SelectedBox.Visible = True
        SelectedBox.TextAlign = HorizontalAlignment.Center
        SelectedBox.Select()
    End Sub

    Private Sub DrawCells(g As Graphics)
        For x As Integer = 0 To GridSize.Width - 1
            For y As Integer = 0 To GridSize.Height - 1
                Cells(x, y).Draw(g)
            Next
        Next
    End Sub

    Private Sub DrawSelection(g As Graphics)
        If (Me.ActiveCell IsNot Nothing) Then
            g.DrawRectangle(New Pen(Color.DodgerBlue, 2), New Rectangle(New Point(CellSize.Width * ActiveCell.Location.X + CellSize.Width, CellSize.Height * ActiveCell.Location.Y + CellSize.Height), New Size(CellSize.Width + 1, CellSize.Height + 1)))
        End If
    End Sub

    Private Sub cell_Keydown(sender As Object, e As KeyEventArgs)
        If FormulaMode Then
            equation_KeyDown(sender, e)
            e.SuppressKeyPress = True
            e.Handled = True
        Else 
                If e.KeyCode = Keys.Enter Then
                    ActiveCell.Text = CType(sender, TextBox).Text  
                    SelectedBox.Visible = False
                ElseIf e.KeyData = 187 Then
                    FormulaMode = True
                    FormulaBox.Text = "="
                FormulaBox.Select()
                    FormulaBox.Select(FormulaBox.Text.Length, 0)
                End If
            End If
    End Sub

    Private Sub equation_KeyDown(sender As Object, e As KeyEventArgs)
        If (FormulaMode = True) And (e.KeyCode = Keys.Enter) Then
            FormulaMode = False
            ActiveCell.Text = CType(sender, TextBox).Text
            SelectedBox.Visible = False
            Invalidate()
            e.SuppressKeyPress = True
        End If
    End Sub 
End Class
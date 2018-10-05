Imports System.Text.RegularExpressions

Public Class Cell
    Private _text As String = ""
    Public Property Text As String
        Get
            Return _text
        End Get
        Set(value As String)
            If value.StartsWith("=") Then
                Me.Equation = value
                _text = value
            Else
                _text = value
            End If
        End Set
    End Property

    Public Sheet As Sheet
    Public Property Equation As String = ""
    Public Property Location As Point
    Public Property Fill As Color = Color.White
    Public Property Border As Color = Color.Silver
    Public Property Forecolor As Color = Color.Black
    Public Property Font As Font = New Font("Microsoft Sans Serif", 8.25)
    Public Property TextAlignment As New StringFormat() With {.LineAlignment = StringAlignment.Center, .Alignment = StringAlignment.Center}
    Public Property Rectangle As Rectangle

    Sub Draw(g As Graphics)
        g.FillRectangle(New SolidBrush(Fill), Me.Rectangle)
        g.DrawRectangle(New Pen(Me.Border), Me.Rectangle)
        If Not String.IsNullOrEmpty(Equation) Then
            Try
                Dim result As Double = New MathParserNet.Parser().SimplifyDouble(parseEquation(Me.Equation, Sheet.Cells))
                Me.Text = result.ToString
                g.DrawString(result, Me.Font, New SolidBrush(Me.Forecolor), Me.Rectangle, Me.TextAlignment)
            Catch ex As Exception
                g.DrawString("#ERROR " & ex.Message.ToString, Me.Font, New SolidBrush(Me.Forecolor), Me.Rectangle, Me.TextAlignment)
            End Try
        Else
            g.DrawString(Text, Font, New SolidBrush(Forecolor), Me.Rectangle, Me.TextAlignment)
        End If
    End Sub

    Public Function parseEquation(input As String, cells As Cell(,)) As String
        Dim regexMatches As MatchCollection = New Regex("\{(.*?)\}").Matches(input)
        Dim output As String = input.Replace("[SCC]", regexMatches.Count).Replace("[PI]", "(3.14159265358979323846)").Replace("[HALF_PI]", "1.57079632679489661923").Replace("[QUARTER_PI]", "0.7853982").Replace("[TAU]", "6.28318530717958647693").Replace("[TWO_PI]", "6.28318530717958647693").Replace("[]", "").Replace("[]", "")
        Dim cellCount As Integer = 0
        For Each cellReplacement As Match In regexMatches
            Dim splitCordinates As String() = cellReplacement.Groups(1).Value.Replace(" ", "").Replace("{", "").Replace("}", "").Split(","c)
            output = output.Replace(cellReplacement.Groups(1).Value, CellFromLocation(Convert.ToInt32(splitCordinates(0)) - 1, Convert.ToInt32(splitCordinates(1)) - 1, cells).Text)
        Next
        Return output.Replace("{", "(").Replace("}", ")").Remove(0, 1)
        Return Nothing
    End Function 
End Class
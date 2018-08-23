Imports System.Drawing

Namespace Features.Formatting.Borders
    Public Class AddBordersToCell
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeB2 As IRange = worksheet.Range!B2
            'set left, top, right, bottom borders together.

            rangeB2.Borders.LineStyle = BorderLineStyle.DashDot
            rangeB2.Borders.Color = Color.Green

            'set top border individually.
            rangeB2.Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Medium
            rangeB2.Borders(BordersIndex.EdgeTop).Color = Color.Red

            'set diagonal down border individually.
            rangeB2.Borders(BordersIndex.DiagonalDown).LineStyle = BorderLineStyle.Hair
            rangeB2.Borders(BordersIndex.DiagonalDown).Color = Color.Blue

            'set diagonal up border individually.
            rangeB2.Borders(BordersIndex.DiagonalUp).LineStyle = BorderLineStyle.Dotted
            rangeB2.Borders(BordersIndex.DiagonalUp).Color = Color.Blue
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

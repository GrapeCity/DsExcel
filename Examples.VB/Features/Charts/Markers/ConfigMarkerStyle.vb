Namespace Features.Charts.Markers
    Public Class ConfigMarkerStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.LineMarkers, 250, 20, 360, 230)
            worksheet.Range("A1:B6").Value = New Object(,)  {
                {Nothing, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 40}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B6"), RowCol.Columns, True, True)
            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)

            'config line markers style
            series1.MarkerStyle = MarkerStyle.Square
            series1.MarkerSize = 10
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

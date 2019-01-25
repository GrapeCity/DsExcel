Namespace Features.Charts.ChartLines
    Public Class CreateTrendline
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            series1.Trendlines.Add()
            series1.Trendlines(0).Type = TrendlineType.Linear
            series1.Trendlines(0).Forward = 1
            series1.Trendlines(0).Backward = 0.5
            series1.Trendlines(0).Intercept = 2.5
            series1.Trendlines(0).DisplayEquation = True
            series1.Trendlines(0).DisplayRSquared = True
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

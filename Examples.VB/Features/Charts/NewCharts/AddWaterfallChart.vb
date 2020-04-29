Namespace Features.Charts.NewCharts
    Public Class AddWaterfallChart
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:B8").Value = New Object(,) {
                {"Starting Amt", 130},
                {"Measurement 1", 25},
                {"Measurement 2", -75},
                {"Subtotal", 80},
                {"Measurement 3", 45},
                {"Measurement 4", -65},
                {"Measurement 5", 80},
                {"Total", 140}
            }
            worksheet.Range("A:A").Columns.AutoFit()

            'Create a waterfall chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Waterfall, 300, 20, 300, 250)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B8"))

            'Set subtotal&total points.
            Dim points As IPoints = shape.Chart.SeriesCollection(0).Points
            points(3).IsTotal = True
            points(7).IsTotal = True

            'Connector lines are not shown.
            Dim series As ISeries = shape.Chart.SeriesCollection(0)
            series.ShowConnectorLines = False

            'Modify the fill color of the first legend entry.
            Dim LegendEntries As ILegendEntries = shape.Chart.Legend.LegendEntries
            LegendEntries(0).Format.Fill.Color.ObjectThemeColor = ThemeColor.Accent6
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

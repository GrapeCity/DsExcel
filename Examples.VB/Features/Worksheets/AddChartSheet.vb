Namespace Features.Worksheets
    Public Class AddChartSheet
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:D6").Value = New Object(,) {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", 51, 36, 27},
                {"Item3", 52, 85, 30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }

            'Add a chart sheet.
            Dim chartSheet As IWorksheet = workbook.Worksheets.Add(SheetType.Chart)

            'Add the main chart for the chart sheet.
            Dim mainChart As IShape = chartSheet.Shapes.AddChart(ChartType.ColumnClustered, 100, 100, 200, 200)
            mainChart.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"))

            'Make the chart sheet the active sheet.
            chartSheet.Activate()
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

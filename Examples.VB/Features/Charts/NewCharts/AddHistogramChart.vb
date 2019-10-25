Namespace Features.Charts.NewCharts
    Public Class AddHistogramChart
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:B11").Value = New Object(,) {
                {"Complaint", "Count"},
                {"Too noisy", 27},
                {"Overpriced", 789},
                {"Food is tasteless", 65},
                {"Food is not fresh", 9},
                {"Food is too salty", 15},
                {"Not clean", 30},
                {"Unfriendly staff", 12},
                {"Wait time", 109},
                {"No atmosphere", 45},
                {"Small portions", 621}
            }
            worksheet.Range("A:A").Columns.AutoFit()

            'Create a histogram chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Histogram, 300, 20, 300, 200)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B11"))

            With shape.Chart.ChartGroups(0)
                'Sets bins type by count.
                .BinsType = BinsType.BinsTypeBinCount
                .BinsCountValue = 3

                'Set overflow bin value
                .BinsOverflowEnabled = True
                .BinsOverflowValue = 500
            End With
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

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

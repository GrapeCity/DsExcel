Namespace Features.Charts.NewCharts
    Public Class AddFunnelChart
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:B7").Value = New Object(,) {
                {"Stage", "Amount"},
                {"Prospects", 500},
                {"Qualified prospects", 425},
                {"Needs analysis", 200},
                {"Price quotes", 150},
                {"Negotiations", 100},
                {"Closed sales", 90}
            }
            worksheet.Range("A:A").Columns.AutoFit()

            'Create a funnel chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Funnel, 300, 20, 300, 200)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B7"))

            'Set the axis invisible.
            Dim axis As IAxis = shape.Chart.Axes.Item(AxisType.Category, AxisGroup.Primary)
            axis.Visible = False
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

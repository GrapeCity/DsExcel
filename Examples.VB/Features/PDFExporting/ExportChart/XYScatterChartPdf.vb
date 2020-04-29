Namespace Features.PDFExporting.ExportChart
    Public Class XYScatterChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.XYScatter, 20, 20, 360, 230)
            worksheet.Range("A20:D25").Value = New Object(,) {
                {55, 964, 67, 475},
                {20, 825, 10, 163},
                {77, 840, 87, 224},
                {182, 596, 46, 196},
                {190, 384, 100, 377},
                {140, 503, 92, 47}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A20:B25"), RowCol.Columns)
            shape.Chart.SeriesCollection.Add(worksheet.Range("C20:D25"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Scatter Chart"
            'config markers style
            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            Dim series2 As ISeries = shape.Chart.SeriesCollection(1)
            series1.MarkerSize = 10
            series2.MarkerSize = 10
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
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

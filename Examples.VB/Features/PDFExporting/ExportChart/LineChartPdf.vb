Namespace Features.PDFExporting.ExportChart
    Public Class LineChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Line, 20, 20, 360, 230)
            worksheet.Range("A20:C26").Value = New Object(,) {
                {0, 59.18, 27.14},
                {44.64, 52.22, 25.08},
                {45.21, 49.8, 57.99},
                {24.32, 37.3, 42.73},
                {58.34, 34.43, 28.34},
                {31.89, 69.78, 46.88},
                {41.79, 63.94, 56.24}
            }
            With shape.Chart
                .SeriesCollection.Add(worksheet.Range("A20:C26"), RowCol.Columns)
                'set series lines style
                .SeriesCollection(0).Format.Line.Weight = 2.25
                .SeriesCollection(1).Format.Line.Weight = 2.25
                .SeriesCollection(2).Format.Line.Weight = 2.25
                .ChartTitle.Text = "Line Chart"
            End With
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

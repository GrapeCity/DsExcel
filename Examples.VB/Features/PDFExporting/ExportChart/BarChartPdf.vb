Namespace Features.PDFExporting.ExportChart
    Public Class BarChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.BarClustered, 20, 20, 360, 230)
            worksheet.Range("A20:D21").Value = New Object(,) {
                {100, 200, 300, 400},
                {100, 200, 300, 400}
            }
            With shape.Chart
                .SeriesCollection.Add(worksheet.Range("A20:D21"), RowCol.Columns)
                .ChartTitle.Text = "Bar Clustered Chart"
                .Legend.Position = LegendPosition.Left
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

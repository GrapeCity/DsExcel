Namespace Features.PDFExporting.ExportChart
    Public Class PieChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Pie, 20, 20, 360, 230)
            worksheet.Range("A20:B23").Value = New Object(,) {
                {"Blue", 1}, {"Red", 2}, {"Green", 3}, {"Purple", 4}
            }

            With shape.Chart
                .SeriesCollection.Add(worksheet.Range("A20:B23"), RowCol.Columns)
                .ChartTitle.Text = "Pie Chart"
                .Legend.Position = LegendPosition.Right
            End With
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

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

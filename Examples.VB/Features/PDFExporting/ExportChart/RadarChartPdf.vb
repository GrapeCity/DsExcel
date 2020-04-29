Namespace Features.PDFExporting.ExportChart
    Public Class RadarChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Radar, 20, 20, 360, 230)
            worksheet.Range("B20:C20").Value = {
                {"S1", "S2"}
            }
            worksheet.Range("A21:A25").Value = {
                {"A"}, {"B"}, {"C"}, {"D"}, {"E"}
            }
            worksheet.Range("B21:C25").Value = {
                {10.0, 25}, {51, 36}, {52, 85}, {22, 65}, {23, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A20:C25"))
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

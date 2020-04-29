Imports System.Drawing

Namespace Features.PDFExporting.ExportChart
    Public Class ErrorbarPdf
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Line, 20, 200, 360, 230)
            worksheet.Range("A1:D4").Value = New Object(,) {
                {Nothing, "Q1", "Q2", "Q3"},
                {"Mobile Phones", 1330, 2330, 3330},
                {"Laptops", 4032, 5632, 6197},
                {"Tablets", 6233, 7233, 8233}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D4"), RowCol.Rows)

            'get first series
            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)

            'set HasErrorBars as true
            series1.HasErrorBars = True

            'config style of first series' error bar
            With series1.YErrorBar.Format.Line
                .Color.RGB = Color.Red
                .Weight = 2
            End With
        End Sub
    End Class
End Namespace

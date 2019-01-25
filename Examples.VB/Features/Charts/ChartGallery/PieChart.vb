Namespace Features.Charts.ChartGallery
    Public Class PieChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Pie, 250, 20, 360, 230)
            worksheet.Range("A1:B4").Value = New Object(,)  {
                {"Blue", 1},
                {"Red", 2},
                {"Green", 3},
                {"Purple", 4}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B4"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Pie Chart"
            shape.Chart.Legend.Position = LegendPosition.Right
        End Sub
    End Class
End Namespace

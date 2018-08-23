Namespace Features.Charts.ChartGallery
    Public Class BarStackedChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.BarStacked, 250, 20, 360, 230)
            worksheet.Range("A1:C4").Value = {
                {103, 121, 109},
                {56, 94, 115},
                {116, 89, 99},
                {55, 93, 70}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C4"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Bar Stacked Chart"
            shape.Chart.Legend.Position = LegendPosition.Left
        End Sub
    End Class
End Namespace

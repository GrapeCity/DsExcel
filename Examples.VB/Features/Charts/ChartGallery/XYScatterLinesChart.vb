Namespace Features.Charts.ChartGallery
    Public Class XYScatterLinesChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.XYScatterLines, 250, 20, 360, 230)
            worksheet.Range("A1:B8").Value = New Object(,)  {
                {75, 250},
                {50, 125},
                {25, 375},
                {75, 250},
                {50, 875},
                {25, 625},
                {75, 750},
                {125, 500}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B8"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Scatter with Straight Lines and Markers Chart"
        End Sub
    End Class
End Namespace

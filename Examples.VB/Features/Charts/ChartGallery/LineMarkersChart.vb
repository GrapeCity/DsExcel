Namespace Features.Charts.ChartGallery
    Public Class LineMarkersChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.LineMarkers, 250, 20, 360, 230)
            worksheet.Range("A1:B8").Value = {
                {6, 55},
                {45, 25},
                {35, 45},
                {25, 65},
                {65, 15},
                {45, 75},
                {75, 55},
                {65, 35}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B8"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Line with Markers"
            shape.Chart.SeriesCollection(0).MarkerStyle = MarkerStyle.Square
        End Sub
    End Class
End Namespace

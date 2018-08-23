Namespace Features.Charts.ChartGallery
    Public Class XYScatterSmoothWithMarkers
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.XYScatterSmooth, 250, 20, 360, 230)
            worksheet.Range("A1:B5").Value = {
                {4, 2},
                {6, 1},
                {1, 2},
                {7, 4},
                {4, 4}
            }
            worksheet.Range("A7:B11").Value = {
                {9, 5},
                {7, 8},
                {9, 8},
                {5, 9},
                {2, 4}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B5"), RowCol.Columns)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A7:B11"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Scatter with Smooth Lines and Markers"
        End Sub
    End Class
End Namespace

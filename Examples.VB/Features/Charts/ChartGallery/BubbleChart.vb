Namespace Features.Charts.ChartGallery
    Public Class BubbleChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Bubble, 250, 20, 360, 230)
            worksheet.Range("A1:C10").Value = New Object(,)  {
                {"Blue", Nothing, Nothing},
                {125, 750, 3},
                {25, 625, 7},
                {75, 875, 5},
                {175, 625, 6},
                {"Red", Nothing, Nothing},
                {125, 500, 10},
                {25, 250, 1},
                {75, 125, 5},
                {175, 250, 8}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C5"), RowCol.Columns)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A6:C10"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Bubble Chart"
        End Sub
    End Class
End Namespace

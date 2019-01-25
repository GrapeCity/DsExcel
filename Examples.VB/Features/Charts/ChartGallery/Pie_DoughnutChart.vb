Namespace Features.Charts.ChartGallery
    Public Class Pie_DoughnutChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Doughnut, 250, 20, 360, 230)
            worksheet.Range("A1:B6").Value = New Object(,)  {
                {"S1", "S2"},
                {10, 25},
                {51, 36},
                {52, 85},
                {22, 65},
                {23, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B6"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Area Chart"
            shape.Chart.ChartGroups(0).DoughnutHoleSize = 50
            shape.Chart.SeriesCollection(0).HasDataLabels = True
            shape.Chart.SeriesCollection(1).HasDataLabels = True
            shape.Chart.SeriesCollection(1).Explosion = 2
        End Sub
    End Class
End Namespace

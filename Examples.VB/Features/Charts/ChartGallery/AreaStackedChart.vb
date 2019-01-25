Namespace Features.Charts.ChartGallery
    Public Class AreaStackedChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.AreaStacked, 250, 20, 360, 230)
            worksheet.Range("A1:C13").Value = New Object(,)  {
                {0, 59.18, 27.14},
                {44.64, 52.22, 25.08},
                {45.21, 49.8, 57.99},
                {24.32, 37.3, 42.73},
                {58.34, 34.43, 28.34},
                {31.89, 69.78, 46.88},
                {41.79, 63.94, 56.24},
                {67.94, 57.4, 27.78},
                {49.87, 48.26, 52.06},
                {62.39, 67.43, 33.33},
                {54.76, 22.95, 50.36},
                {28.33, 36.6, 36.61},
                {22.77, 55.65, 65.64},
                {20.34, 49.35, 45.6},
                {32.1, 47.6, 20.62},
                {26.37, 63.0, 53.97},
                {35, 75, 60}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C13"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Area Stacked Chart"
        End Sub
    End Class
End Namespace

Namespace Features.Charts.ChartGallery
    Public Class CombinationChart1
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:B13").Value = {
                {"Blue Column", "Red Line"},
                {75, 20},
                {149, 50},
                {105, 30},
                {55, 80},
                {121, 40},
                {76, 110},
                {128, 50},
                {114, 140},
                {75, 60},
                {105, 170},
                {145, 70},
                {110, 100}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B13"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Combination Chart"

            'change series type
            Dim series2 As ISeries = shape.Chart.SeriesCollection(1)
            series2.ChartType = ChartType.LineMarkers
        End Sub
    End Class
End Namespace

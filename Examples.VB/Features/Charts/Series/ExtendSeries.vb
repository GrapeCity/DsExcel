Namespace Features.Charts.Series
    Public Class ExtendSeries
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D4").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 50},
                {"Item2", 15, -36, 40},
                {"Item3", 52, 40, -30}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D4"), RowCol.Columns, True, True)
            worksheet.Range("A12:D13").Value = {
                {"Item5", 10, 20, -30},
                {"Item6", 20, 40, 80}
            }

            'add new data point to existing series.
            shape.Chart.SeriesCollection.Extend(worksheet.Range("A12:D13"), RowCol.Columns, True)
        End Sub
    End Class
End Namespace

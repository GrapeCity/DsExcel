Namespace Features.Charts.Series
    Public Class DeleteSeries
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            'delete chart series, make series count to 1.
            shape.Chart.SeriesCollection(2).Delete()
            shape.Chart.SeriesCollection(1).Delete()

            'Chart's series count is 1.
            Dim count = shape.Chart.SeriesCollection.Count
        End Sub
    End Class
End Namespace

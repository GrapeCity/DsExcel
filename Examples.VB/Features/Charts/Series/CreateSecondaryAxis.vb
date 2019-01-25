Namespace Features.Charts.Series
    Public Class CreateSecondaryAxis
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:C6").Value = New Object(,)  {
                {Nothing, "S1", "S2"},
                {"Item1", 10, 25},
                {"Item2", -51, -36},
                {"Item3", 32, 64},
                {"Item4", 44, 80},
                {"Item5", 60, 100}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)
            Dim series2 As ISeries = shape.Chart.SeriesCollection(1)

            'add a secondary axis
            series2.AxisGroup = AxisGroup.Secondary
            series2.ChartType = ChartType.Line
        End Sub
    End Class
End Namespace

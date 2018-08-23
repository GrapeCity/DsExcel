Namespace Features.Charts.Axes
    Public Class ConfigValueAxisUnits
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)

            worksheet.Range("A1:D6").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 90, -30},
                {"Item4", 22, 65, 50},
                {"Item5", 23, 55, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim category_axis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
            Dim value_axis As IAxis = shape.Chart.Axes.Item(AxisType.Value)

            value_axis.MaximumScale = 100
            value_axis.MinimumScale = -100
            value_axis.MajorUnit = 30
            value_axis.MinorUnit = 6
        End Sub
    End Class
End Namespace

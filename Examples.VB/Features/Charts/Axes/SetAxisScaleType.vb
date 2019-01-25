Namespace Features.Charts.Axes
    Public Class SetAxisScaleType
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D5").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 4, 25, 7},
                {"Item2", 15, -10, 18},
                {"Item3", 45, 90, 20},
                {"Item4", 8, 20, 11}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D5"), RowCol.Columns, True, True)

            Dim value_axis As IAxis = shape.Chart.Axes.Item(AxisType.Value)
            value_axis.ScaleType = ScaleType.Logarithmic
            value_axis.LogBase = 5
        End Sub
    End Class
End Namespace

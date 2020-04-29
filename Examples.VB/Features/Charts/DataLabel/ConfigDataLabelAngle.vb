Namespace Features.Charts.DataLabel
    Public Class ConfigDataLabelAngle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:B5").Value = New Object(,) {
                {Nothing, "S1"},
                {"Item1", -20},
                {"Item2", 30},
                {"Item3", 50},
                {"Item3", 40}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B5"), RowCol.Columns, True, True)

            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            series1.HasDataLabels = True

            ' set series1's all data labels' angle.
            series1.DataLabels.Orientation = 45

            ' set series1's specific data label's angle.
            series1.DataLabels(2).Orientation = -45
        End Sub
    End Class
End Namespace

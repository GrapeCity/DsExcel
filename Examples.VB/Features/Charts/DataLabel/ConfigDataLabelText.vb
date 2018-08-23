Namespace Features.Charts.DataLabel
    Public Class ConfigDataLabelText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:B5").Value = {
                {Nothing, "S1", "S2"},
                {"Item1", -20, 10},
                {"Item2", 30, 60},
                {"Item3", 50, 80},
                {"Item3", 40, 70}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B5"), RowCol.Columns, True, True)

            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            series1.HasDataLabels = True

            'customize data lables' text.
            series1.DataLabels.ShowCategoryName = True
            series1.DataLabels.ShowSeriesName = True
            series1.DataLabels.ShowLegendKey = True
        End Sub
    End Class
End Namespace

Imports System.Drawing

Namespace Features.Charts.Legend
    Public Class ConfigLegendPositionAndLayout
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -20, 36, 27},
                {"Item3", 52, 70, 30},
                {"Item4", 22, 33, -20},
                {"Item5", 23, 30, 30}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)
            shape.Chart.HasLegend = True
            Dim legend As ILegend = shape.Chart.Legend

            'position.
            legend.Position = LegendPosition.Left

            'font.
            legend.Font.Color.RGB = Color.Red
            legend.Font.Italic = True
        End Sub
    End Class
End Namespace

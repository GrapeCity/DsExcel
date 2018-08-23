Imports System.Drawing

Namespace Features.Charts.Legend
    Public Class ConfigLegendFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)
            shape.Chart.HasLegend = True

            'config legend font style
            Dim legend As ILegend = shape.Chart.Legend
            legend.Font.Size = 12
            legend.Font.Name = "Cooper Black"

            'config legend format
            legend.Format.Fill.Color.RGB = Color.LightGray
            legend.Format.Line.Color.RGB = Color.Gray
        End Sub
    End Class
End Namespace

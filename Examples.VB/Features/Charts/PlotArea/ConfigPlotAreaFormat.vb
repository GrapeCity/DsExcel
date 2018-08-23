Imports System.Drawing

Namespace Features.Charts.PlotArea
    Public Class ConfigPlotAreaFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 50, -30},
                {"Item4", 22, 65, 30},
                {"Item5", 23, 40, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim plotarea As IPlotArea = shape.Chart.PlotArea
            plotarea.Format.Fill.Color.RGB = Color.LightGray
            plotarea.Format.Line.Color.RGB = Color.Gray
            plotarea.Format.Line.Weight = 1
        End Sub
    End Class
End Namespace

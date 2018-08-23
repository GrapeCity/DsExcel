Imports System.Drawing

Namespace Features.Charts.ChartArea
    Public Class ConfigChartAreaFont
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 60, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim chartarea As IChartArea = shape.Chart.ChartArea

            'Font
            chartarea.Font.Color.RGB = Color.MediumSeaGreen
            chartarea.Font.Name = "Times New Roman"
            chartarea.Font.Size = 12

            'rounded corners.
            chartarea.RoundedCorners = True
        End Sub
    End Class
End Namespace

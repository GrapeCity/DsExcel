Imports System.Drawing

Namespace Features.Charts.ChartTitle
    Public Class ConfigChartTitleStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -20, 36, 27},
                {"Item3", 62, 70, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 50, 50}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            'config chart title style
            shape.Chart.HasTitle = True
            shape.Chart.ChartTitle.Text = "MyChartTitle"
            shape.Chart.ChartTitle.Font.Size = 15
            shape.Chart.ChartTitle.Font.Name = "Bradley Hand ITC"
            shape.Chart.ChartTitle.Format.Fill.Color.RGB = Color.DarkOrange
            shape.Chart.ChartTitle.Format.Line.Color.RGB = Color.CornflowerBlue
            shape.Chart.ChartTitle.Format.Line.Weight = 1.5
        End Sub
    End Class
End Namespace

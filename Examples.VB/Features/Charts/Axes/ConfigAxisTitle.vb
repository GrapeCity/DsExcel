Imports System.Drawing

Namespace Features.Charts.Axes
    Public Class ConfigAxisTitle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)

            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", 51, 36, 27},
                {"Item3", 52, 85, 30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim category_axis As IAxis = shape.Chart.Axes.Item(AxisType.Category)

            'config axis title
            category_axis.HasTitle = True
            category_axis.AxisTitle.Text = "CategoryAxisTitle"
            category_axis.AxisTitle.Font.Size = 18
            category_axis.AxisTitle.Font.Color.RGB = Color.Orange
        End Sub
    End Class
End Namespace

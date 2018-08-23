﻿Namespace Features.Charts.ChartGallery
    Public Class ColumnStacked100Chart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnStacked100, 250, 20, 360, 230)
            worksheet.Range("A1:B6").Value = {
                {1, 5},
                {2, 4},
                {3, 3},
                {4, 2},
                {5, 1},
                {5, 3}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B6"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Column Stacked 100 Chart"
        End Sub
    End Class
End Namespace

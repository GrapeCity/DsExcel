﻿Namespace Features.Charts.ChartGallery
    Public Class LineMarkerStacked100Chart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.LineMarkersStacked100, 250, 20, 360, 230)
            worksheet.Range("A1:C5").Value = {
                {12, 22, 27},
                {45, 52, 25},
                {58, 35, 58},
                {21, 37, 43},
                {44, 45, 28}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C5"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Line Marker Stacked 100 Chart"
        End Sub
    End Class
End Namespace

Imports System.Drawing

Namespace Features.Charts.Markers
    Public Class ConfigMarkersFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.LineMarkers, 250, 20, 360, 230)
            worksheet.Range("A1:B6").Value = {
                {Nothing, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 40}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B6"), RowCol.Columns, True, True)
            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)

            'config line markers style
            series1.MarkerFormat.Fill.Color.RGB = Color.CornflowerBlue
            series1.MarkerFormat.Line.Style = LineStyle.ThickThin
            series1.MarkerFormat.Line.Color.RGB = Color.LightGreen
            series1.MarkerFormat.Line.Weight = 3
        End Sub
    End Class
End Namespace

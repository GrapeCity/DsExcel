Imports System.Drawing

Namespace Features.Charts.DataLabel
    Public Class ConfigDataLabelFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:B5").Value = {
                {Nothing, "S1"},
                {"Item1", -20},
                {"Item2", 30},
                {"Item3", 50},
                {"Item3", 40}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B5"), RowCol.Columns, True, True)

            Dim series1 As ISeries = shape.Chart.SeriesCollection(0)
            series1.HasDataLabels = True
            series1.DataLabels.ShowSeriesName = True

            'set series1's all data label's format.
            series1.DataLabels.Format.Fill.Color.RGB = Color.Pink
            series1.DataLabels.Format.Line.Color.RGB = Color.Green
            series1.DataLabels.Format.Line.Weight = 1

            'set series1's specific data label's format.
            series1.DataLabels(2).Format.Fill.Color.RGB = Color.LightGreen
            series1.Points(2).DataLabel.Format.Line.Color.RGB = Color.Gray
            series1.Points(2).DataLabel.Format.Line.Weight = 2
        End Sub
    End Class
End Namespace

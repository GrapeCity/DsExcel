Namespace Features.Charts.ChartGallery
    Public Class BarStacked100Chart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.BarStacked100, 250, 20, 360, 230)
            worksheet.Range("A1:B5").Value = New Object(,)  {
                {1, 5},
                {2, 4},
                {3, 3},
                {4, 2},
                {4, 1}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:B5"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Bar Stacked 100 Chart"
            shape.Chart.Legend.Position = LegendPosition.Left
        End Sub
    End Class
End Namespace

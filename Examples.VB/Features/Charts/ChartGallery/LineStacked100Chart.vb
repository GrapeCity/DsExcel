Namespace Features.Charts.ChartGallery
    Public Class LineStacked100Chart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.LineStacked100, 250, 20, 360, 230)
            worksheet.Range("A1:C5").Value = New Object(,)  {
                {12, 22, 27},
                {45, 52, 25},
                {58, 35, 58},
                {21, 37, 43},
                {44, 45, 28}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C5"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Line Stacked 100 Chart"
            shape.Chart.SeriesCollection(0).Format.Line.Weight = 2.25
            shape.Chart.SeriesCollection(1).Format.Line.Weight = 2.25
            shape.Chart.SeriesCollection(2).Format.Line.Weight = 2.25
        End Sub
    End Class
End Namespace

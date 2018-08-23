Namespace Features.Charts.ChartGallery
    Public Class ColumnStackedChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnStacked, 250, 20, 360, 230)
            worksheet.Range("A1:C6").Value = {
                {103, 121, 109},
                {56, 94, 115},
                {116, 89, 99},
                {55, 93, 70},
                {114, 114, 83},
                {125, 138, 136}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C6"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Column Stacked Chart"
        End Sub
    End Class
End Namespace

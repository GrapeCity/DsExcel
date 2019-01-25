Namespace Features.Charts.ChartGallery
    Public Class ColumnClusteredChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            worksheet.Range("A1:D2").Value = New Object(,)  {
                {100, 200, 300, 400},
                {100, 200, 300, 400}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D2"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Column Clustered Chart"
        End Sub
    End Class
End Namespace

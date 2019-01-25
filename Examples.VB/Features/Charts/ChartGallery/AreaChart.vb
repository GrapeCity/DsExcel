Namespace Features.Charts.ChartGallery
    Public Class AreaChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Area, 250, 20, 360, 230)
            worksheet.Range("A1:C13").Value = New Object(,)  {
                {Nothing, "Blue Series", "Orange Series"},
                {"Jan", 0, 59.1883603948205},
                {"Feb", 44.6420211591501, 52.2280901938606},
                {"Mar", 45.2174930051225, 49.8093056416248},
                {"Apr", 62, 37.3065749226828},
                {"May", 53, 34.4312192530766},
                {"Jun", 31.8933622049831, 69.7834561753736},
                {"Jul", 41.7930895085093, 63.9418103906982},
                {"Aug", 73, 57.4049534494926},
                {"Sep", 49.8773891668518, 33},
                {"Oct", 50, 74},
                {"Nov", 54.7658428630216, 22.9587876597096},
                {"Dec", 32, 54}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:C13"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Area Chart"
        End Sub
    End Class
End Namespace

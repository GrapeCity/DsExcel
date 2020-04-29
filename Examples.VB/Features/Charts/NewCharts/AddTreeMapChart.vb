Namespace Features.Charts.NewCharts
    Public Class AddTreeMapChart
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:D16").Value = New Object(,) {
                {"Quarter", "Month", "Week", "Output"},
                {"1st", "Jan", Nothing, 3.5},
                {Nothing, "Feb", "Week1", 1.2},
                {Nothing, Nothing, "Week2", 0.8},
                {Nothing, Nothing, "Week3", 0.6},
                {Nothing, Nothing, "Week4", 0.5},
                {Nothing, "Mar", Nothing, 1.7},
                {"2st", "Apr", Nothing, 1.1},
                {Nothing, "May", Nothing, 0.8},
                {Nothing, "Jun", Nothing, 0.3},
                {"3st", "July", Nothing, 0.7},
                {Nothing, "Aug", Nothing, 0.6},
                {Nothing, "Sept", Nothing, 0.1},
                {"4st", "Oct", Nothing, 0.5},
                {Nothing, "Nov", Nothing, 0.4},
                {Nothing, "Dec", Nothing, 0.3}
            }

            'Create a treemap chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Treemap, 300, 20, 300, 200)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D16"))

            'Set the parent data labels are displayed as banners.
            shape.Chart.SeriesCollection(0).ParentDataLabelOption = Drawing.ParentDataLabelOptions.Banner

            'Modify chart title text.
            shape.Chart.ChartTitle.Text = "Annual Report"
        End Sub
    End Class
End Namespace

Namespace Features.Charts.Axes
    Public Class ConfigCategoryAxisUnits
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A2:A6").NumberFormat = "m/d/yyyy"
            worksheet.Range("A1:D6").Value = New Object(,) {
                {Nothing, "S1", "S2", "S3"},
                {
                    #10/7/2015#, 10,
                    25,
                    25
                },
                {
                    #10/24/2015#,
                    51,
                    36,
                    27
                },
                {
                    #11/8/2015#,
                    52,
                    85,
                    30
                },
                {
                    #11/25/2015#,
                    22,
                    65,
                    65
                },
                {
                    #12/10/2015#,
                    23,
                    69,
                    69
                }
            }

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.ColumnClustered, 250, 20, 360, 230)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim category_axis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
            category_axis.MaximumScale = (#12/20/2015#).ToOADate()
            category_axis.MinimumScale = (#10/1/2015#).ToOADate()
            category_axis.BaseUnit = TimeUnit.Months
            category_axis.MajorUnitScale = TimeUnit.Months
            category_axis.MajorUnit = 1
            category_axis.MinorUnitScale = TimeUnit.Days
            category_axis.MinorUnit = 15
        End Sub
    End Class
End Namespace

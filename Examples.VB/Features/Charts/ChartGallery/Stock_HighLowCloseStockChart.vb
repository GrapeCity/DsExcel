Namespace Features.Charts.ChartGallery
    Public Class Stock_HighLowCloseStockChart
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.StockHLC, 350, 20, 360, 230)
            worksheet.Range("A1:D17").Value = New Object(,)  {
                {Nothing, "High", "Low", "Close"},
                {
                    #9/1/2019#,
                    105.76,
                    92.38,
                    100.94
                },
                {
                    #9/2/2019#,
                    102.45,
                    90.14,
                    93.45
                },
                {
                    #9/3/2019#,
                    102.11,
                    85.01,
                    99.89
                },
                {
                    #9/4/2019#,
                    106.01,
                    94.04,
                    99.45
                },
                {
                    #9/5/2019#,
                    108.23,
                    98.16,
                    104.33
                },
                {
                    #9/8/2019#,
                    107.7,
                    91.02,
                    102.17
                },
                {
                    #9/9/2019#,
                    110.36,
                    101.62,
                    110.07
                },
                {
                    #9/10/2019#,
                    115.97,
                    106.89,
                    112.39
                },
                {
                    #9/11/2019#,
                    120.32,
                    112.15,
                    117.52
                },
                {
                    #9/12/2019#,
                    122.03,
                    114.67,
                    114.75
                },
                {
                    #9/15/2019#,
                    120.46,
                    106.21,
                    116.85
                },
                {
                    #9/16/2019#,
                    118.08,
                    113.55,
                    116.69
                },
                {
                    #9/17/2019#,
                    128.23,
                    110.91,
                    117.25
                },
                {
                    #9/18/2019#,
                    120.55,
                    108.09,
                    112.52
                },
                {
                    #9/19/2019#,
                    112.58,
                    105.42,
                    109.12
                },
                {
                    #9/22/2019#,
                    115.23,
                    97.25,
                    101.56
                }
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D17"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "High-Low-Close Stock Chart"

            Dim valueAxis As IAxis = shape.Chart.Axes.Item(AxisType.Value)
            Dim categoryAxis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
            Dim series_close As ISeries = shape.Chart.SeriesCollection(2)

            'config value axis
            valueAxis.MinimumScale = 80
            valueAxis.MaximumScale = 140
            valueAxis.MajorUnit = 15

            'config category axis
            categoryAxis.CategoryType = CategoryType.CategoryScale
            categoryAxis.MajorTickMark = TickMark.Outside
            categoryAxis.TickLabelSpacingIsAuto = False
            categoryAxis.TickLabelSpacing = 5
            series_close.MarkerStyle = MarkerStyle.Square
        End Sub
    End Class
End Namespace

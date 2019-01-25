Namespace Features.Charts.ChartGallery
    Public Class Stock_VolumeOpenHighLowClose
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.StockVOHLC, 300, 20, 360, 230)
            worksheet.Range("A1:F23").Value = New Object(,)  {
                {Nothing, "Volume", "Open", "High", "Low", "Close"},
                {
                    #9/1/2019#,
                    26085,
                    103.46,
                    105.76,
                    92.38,
                    100.94
                },
                {
                    #9/2/2019#,
                    52314,
                    100.26,
                    102.45,
                    90.14,
                    93.45
                },
                {
                    #9/3/2019#,
                    70308,
                    98.05,
                    102.11,
                    85.01,
                    99.89
                },
                {
                    #9/4/2019#,
                    33401,
                    100.32,
                    106.01,
                    94.04,
                    99.45
                },
                {
                    #9/5/2019#,
                    87500,
                    99.74,
                    108.23,
                    98.16,
                    104.33
                },
                {
                    #9/8/2019#,
                    33756,
                    92.11,
                    107.7,
                    91.02,
                    102.17
                },
                {
                    #9/9/2019#,
                    65737,
                    107.8,
                    110.36,
                    101.62,
                    110.07
                },
                {
                    #9/10/2019#,
                    45668,
                    107.56,
                    115.97,
                    106.89,
                    112.39
                },
                {
                    #9/11/2019#,
                    47815,
                    112.86,
                    120.32,
                    112.15,
                    117.52
                },
                {
                    #9/12/2019#,
                    76759,
                    115.02,
                    122.03,
                    114.67,
                    114.75
                },
                {
                    #9/15/2019#,
                    23492,
                    108.53,
                    120.46,
                    106.21,
                    116.85
                },
                {
                    #9/16/2019#,
                    56127,
                    114.97,
                    118.08,
                    113.55,
                    116.69
                },
                {
                    #9/17/2019#,
                    81142,
                    127.14,
                    128.23,
                    110.91,
                    117.25
                },
                {
                    #9/18/2019#,
                    46384,
                    118.89,
                    120.55,
                    108.09,
                    112.52
                },
                {
                    #9/19/2019#,
                    51005,
                    105.57,
                    112.58,
                    105.42,
                    109.12
                }
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:F23"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Stock Volume-Open-High-Low-Close Chart"

            Dim valueAxis As IAxis = shape.Chart.Axes.Item(AxisType.Value)
            Dim categoryAxis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
            Dim valueSecondaryAxis As IAxis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Secondary)

            valueAxis.MinimumScale = 0
            valueAxis.MaximumScale = 150000
            valueAxis.MajorUnit = 30000
            categoryAxis.CategoryType = CategoryType.CategoryScale
            categoryAxis.TickLabelSpacing = 5
            valueSecondaryAxis.MajorUnit = 40
        End Sub
    End Class
End Namespace

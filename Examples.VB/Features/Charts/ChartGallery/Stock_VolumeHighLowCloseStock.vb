Imports System.Drawing

Namespace Features.Charts.ChartGallery
    Public Class Stock_VolumeHighLowCloseStock
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.StockVHLC, 350, 20, 360, 230)
            worksheet.Range("A1:E17").Value = {
                {Nothing, "Volume", "High", "Low", "Close"},
                {
                    #9/1/2019#,
                    26085,
                    105.76,
                    92.38,
                    100.94
                },
                {
                    #9/2/2019#,
                    52314,
                    102.45,
                    90.14,
                    93.45
                },
                {
                    #9/3/2019#,
                    70308,
                    102.11,
                    85.01,
                    99.89
                },
                {
                    #9/4/2019#,
                    33401,
                    106.01,
                    94.04,
                    99.45
                },
                {
                    #9/5/2019#,
                    87500,
                    108.23,
                    98.16,
                    104.33
                },
                {
                    #9/8/2019#,
                    33756,
                    107.7,
                    91.02,
                    102.17
                },
                {
                    #9/9/2019#,
                    65737,
                    110.36,
                    101.62,
                    110.07
                },
                {
                    #9/10/2019#,
                    45668,
                    115.97,
                    106.89,
                    112.39
                },
                {
                    #9/11/2019#,
                    47815,
                    120.32,
                    112.15,
                    117.52
                },
                {
                    #9/12/2019#,
                    76759,
                    122.03,
                    114.67,
                    114.75
                },
                {
                    #9/15/2019#,
                    23492,
                    120.46,
                    106.21,
                    116.85
                },
                {
                    #9/16/2019#,
                    56127,
                    118.08,
                    113.55,
                    116.69
                },
                {
                    #9/17/2019#,
                    81142,
                    128.23,
                    110.91,
                    117.25
                },
                {
                    #9/18/2019#,
                    46384,
                    120.55,
                    108.09,
                    112.52
                },
                {
                    #9/19/2019#,
                    51005,
                    112.58,
                    105.42,
                    109.12
                },
                {
                    #9/22/2019#,
                    35223,
                    115.23,
                    97.25,
                    101.56
                }
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:E17"), RowCol.Columns)
            shape.Chart.ChartTitle.Text = "Volume-High-Low-Close Stock Chart"
            shape.Chart.LineGroups(0).HiLoLines.Format.Line.Color.RGB = Color.Black

            Dim valueAxis As IAxis = shape.Chart.Axes.Item(AxisType.Value)
            Dim categoryAxis As IAxis = shape.Chart.Axes.Item(AxisType.Category)
            Dim valueSecondaryAxis As IAxis = shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Secondary)
            Dim series_close As ISeries = shape.Chart.SeriesCollection(3)

            'config value axis
            valueAxis.MinimumScale = 0
            valueAxis.MaximumScale = 150000
            valueAxis.MajorUnit = 30000

            'config category axis
            categoryAxis.CategoryType = CategoryType.CategoryScale
            categoryAxis.MajorTickMark = TickMark.Outside
            categoryAxis.TickLabelSpacing = 4

            'config secondary value axis
            valueSecondaryAxis.MinimumScale = 0
            valueSecondaryAxis.MaximumScale = 150
            valueSecondaryAxis.MajorUnit = 30

            'config marker style
            series_close.MarkerFormat.Fill.Color.RGB = Color.Orange
            series_close.MarkerStyle = MarkerStyle.Square
        End Sub
    End Class
End Namespace

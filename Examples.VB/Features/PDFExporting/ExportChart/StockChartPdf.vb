Namespace Features.PDFExporting.ExportChart
    Public Class StockChartPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.StockOHLC, 20, 20, 360, 220)
            worksheet.Range("A20:E36").Value = New Object(,) {
                {Nothing, "Open", "High", "Low", "Close"},
                {#2019-9-1#, 103.46, 105.76, 92.38, 100.94},
                {#2019-9-2#, 100.26, 102.45, 90.14, 93.45},
                {#2019-9-3#, 98.05, 102.11, 85.01, 99.89},
                {#2019-9-4#, 100.32, 106.01, 94.04, 99.45},
                {#2019-9-5#, 99.74, 108.23, 98.16, 104.33},
                {#2019-9-8#, 92.11, 107.7, 91.02, 102.17},
                {#2019-9-9#, 107.8, 110.36, 101.62, 110.07},
                {#2019-9-10#, 107.56, 115.97, 106.89, 112.39},
                {#2019-9-11#, 112.86, 120.32, 112.15, 117.52},
                {#2019-9-12#, 115.02, 122.03, 114.67, 114.75},
                {#2019-9-15#, 108.53, 120.46, 106.21, 116.85},
                {#2019-9-16#, 114.97, 118.08, 113.55, 116.69},
                {#2019-9-17#, 127.14, 128.23, 110.91, 117.25},
                {#2019-9-18#, 118.89, 120.55, 108.09, 112.52},
                {#2019-9-19#, 105.57, 112.58, 105.42, 109.12},
                {#2019-9-22#, 110.23, 115.23, 97.25, 101.56}
            }
            shape.Chart.SeriesCollection.Add(worksheet.Range("A20:E36"), RowCol.Columns)
            'set chart title
            shape.Chart.ChartTitle.Text = "Open-High-Low-Close Stock Chart"

            'config value axis 
            With shape.Chart.Axes.Item(AxisType.Value)
                .MinimumScale = 80
                .MaximumScale = 140
                .MajorUnit = 15
            End With

            'config category axis
            With shape.Chart.Axes.Item(AxisType.Category)
                .CategoryType = CategoryType.CategoryScale
                .MajorTickMark = TickMark.Outside
                .TickMarkSpacing = 5
                .TickLabelSpacing = 5
            End With
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

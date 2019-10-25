Namespace Features.Charts.NewCharts
    Public Class AddBoxWhiskerChart
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:D16").Value = New Object(,) {
                {"Course", "SchoolA", "SchoolB", "SchoolC"},
                {"English", 63, 53, 45},
                {"Physics", 61, 55, 65},
                {"English", 63, 50, 65},
                {"Math", 62, 51, 64},
                {"English", 46, 53, 66},
                {"English", 58, 56, 67},
                {"Math", 60, 51, 67},
                {"Math", 62, 53, 66},
                {"English", 63, 54, 64},
                {"English", 63, 52, 67},
                {"Physics", 60, 56, 64},
                {"English", 60, 56, 67},
                {"Math", 61, 56, 45},
                {"Math", 63, 58, 64},
                {"English", 59, 54, 65}
            }

            'Create a box&whisker chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.BoxWhisker, 300, 20, 300, 200)
            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D16"))

            'Config value axis's scale.
            With shape.Chart.Axes.Item(AxisType.Value, AxisGroup.Primary)
                .MinimumScale = 40
                .MaximumScale = 70
            End With

            'Config the display of box&whisker plot.  
            With shape.Chart.SeriesCollection(0)
                .ShowInnerPoints = True
                .ShowOutlierPoints = False
                .ShowMeanMarkers = False
                .ShowMeanLine = True
                .QuartileCalculationInclusiveMedian = True
            End With
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

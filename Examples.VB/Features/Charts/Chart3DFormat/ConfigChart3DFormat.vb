Namespace Features.Charts.Chart3DFormat
    Public Class ConfigChart3DFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Column3D, 250, 20, 350, 250)

            worksheet.Range("A1:D6").Value = New Object(,)  {
                {Nothing, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            }

            shape.Chart.SeriesCollection.Add(worksheet.Range("A1:D6"), RowCol.Columns, True, True)

            Dim chartarea As IChartArea = shape.Chart.ChartArea

            'config 3D chart format
            chartarea.Format.ThreeD.RotationX = 60
            chartarea.Format.ThreeD.RotationY = 60
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
    End Class
End Namespace

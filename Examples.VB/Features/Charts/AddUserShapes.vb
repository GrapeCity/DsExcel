Namespace Features.Charts
    Public Class AddUserShapes
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:C10").Value = New Object(,) {
                {"Task", "Worker 1", "Worker 2"},
                {"Task 1", 7, 10},
                {"Task 2", 5, 1},
                {"Task 3", 3, 6},
                {"Task 4", 10, 5},
                {"Task 5", 4, 4},
                {"Task 6", 5, 8},
                {"Task 7", 8, 7},
                {"Task 8", 2, 5},
                {"Task 9", 6, 4}
            }

            'Add a line chart.
            Dim shape As IShape = worksheet.Shapes.AddChart(ChartType.Line, 250, 20, 400, 250)
            With shape.Chart
                .SeriesCollection.Add(worksheet.Range("A1:C10"))
                .HasLegend = False
                .ChartTitle.Text = " "
            End With

            'Add a bussiness logo in the line chart area
            Dim stream = GetResourceStream("logo.png")
            shape.Chart.AddPicture(stream, ImageType.PNG, 170, 10, 60, 10)

            'Add shapes in the line chart area 
            With shape.Chart.AddShape(AutoShapeType.Rectangle, 30, 45, 60, 20)
                .Fill.Color.ObjectThemeColor = ThemeColor.Accent2
                .Line.Color.ObjectThemeColor = ThemeColor.Accent2
                .TextFrame.TextRange(0).Text = "Worker 2"
            End With

            With shape.Chart.AddShape(AutoShapeType.Rectangle, 330, 110, 60, 20)
                .Fill.Color.ObjectThemeColor = ThemeColor.Accent1
                .Line.Color.ObjectThemeColor = ThemeColor.Accent1
                .TextFrame.TextRange(0).Text = "Worker 1"
            End With
        End Sub

        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return True
            End Get
        End Property

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

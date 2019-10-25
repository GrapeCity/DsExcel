Namespace Features.Workbook
    Public Class LoadAndSaveJsonWithShapes
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim workbookWithShape As New Excel.Workbook()
            Dim worksheet As IWorksheet = workbookWithShape.Worksheets(0)

            'Add a shape in worksheet
            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 1, 1, 200, 100)
            shape.Line.DashStyle = LineDashStyle.Dash
            shape.Line.Style = LineStyle.Single
            shape.Line.Weight = 2
            shape.Line.Color.ObjectThemeColor = ThemeColor.Accent6
            shape.Line.Transparency = 0.3

            'jsonString contains shapes
            Dim jsonString As String = workbookWithShape.ToJson()

            'GcExcel can load json string contains shapes now
            workbook.FromJson(jsonString)
        End Sub
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

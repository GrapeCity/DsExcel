Namespace Features.Shape
    Public Class ConfigShapeWithPatternFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 1, 1, 200, 100)
            shape.Fill.Patterned(PatternType.Percent10)
            shape.Fill.Color.ObjectThemeColor = ThemeColor.Accent2
            shape.Fill.PatternColor.ObjectThemeColor = ThemeColor.Accent6
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

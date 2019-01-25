Namespace Features.Shape
    Public Class ConfigShapeLine
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 1, 1, 200, 100)
            shape.Line.DashStyle = LineDashStyle.Dash
            shape.Line.Style = LineStyle.Single
            shape.Line.Weight = 2
            shape.Line.Color.ObjectThemeColor = ThemeColor.Accent6
            shape.Line.Transparency = 0.3
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

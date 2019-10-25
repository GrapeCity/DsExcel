Namespace Features.Shape
    Public Class ConfigShapeWithTextureFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 1, 1, 200, 100)
            shape.Fill.PresetTextured(PresetTexture.Canvas)
            shape.Fill.TextureAlignment = TextureAlignment.Center
            shape.Fill.TextureOffsetX = 2.5
            shape.Fill.TextureOffsetY = 3.2
            shape.Fill.TextureHorizontalScale = 0.9
            shape.Fill.TextureVerticalScale = 0.2
            shape.Fill.Transparency = 0.5
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

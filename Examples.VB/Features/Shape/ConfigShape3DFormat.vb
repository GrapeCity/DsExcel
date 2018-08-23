Namespace Features.Shape
    Public Class ConfigShape3DFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 50, 30, 200, 100)
            shape.ThreeD.RotationX = 50
            shape.ThreeD.RotationY = 20
            shape.ThreeD.RotationZ = 30
            shape.ThreeD.Depth = 7
            shape.ThreeD.Z = 20
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

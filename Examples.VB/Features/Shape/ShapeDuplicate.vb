Namespace Features.Shape
    Public Class ShapeDuplicate
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Create shape
            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 50, 50, 200, 200)

            'Duplicate shape
            Dim newShape As IShape = shape.Duplicate()
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

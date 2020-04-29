Namespace Features.Shape
    Public Class ShapeUngroup
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shapes As IShapes = worksheet.Shapes
            Dim pentagon As IShape = shapes.AddShape(AutoShapeType.RegularPentagon, 89.4, 57.0, 153.6, 90.6)
            Dim pie As IShape = shapes.AddShape(AutoShapeType.Pie, 344.4, 156.8, 50.4, 60.0)
            Dim shpRange As IShapeRange = shapes.Range({pentagon.Name, pie.Name})

            ' Group the shape range
            Dim grouped As IShape = shpRange.Group()

            ' Ungroup the group shape
            shpRange = grouped.Ungroup()
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

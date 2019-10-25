Namespace Features.Shape
    Public Class ConnectShapesByConnector
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim ShapeBegin As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 1, 1, 100, 100)
            Dim EndBegin As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 200, 200, 100, 100)
            Dim ConnectorShape As IShape = worksheet.Shapes.AddConnector(ConnectorType.Straight, 1, 1, 101, 101)

            'connect shapes by connector shape.
            ConnectorShape.ConnectorFormat.BeginConnect(ShapeBegin, 3)
            ConnectorShape.ConnectorFormat.EndConnect(EndBegin, 0)
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

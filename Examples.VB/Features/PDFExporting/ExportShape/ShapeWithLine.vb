Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithLine
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            Dim rectangle As IShape = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 50, 200, 200)
            rectangle.Line.DashStyle = LineDashStyle.Dash
            rectangle.Line.Style = LineStyle.Single
            rectangle.Line.Weight = 7
            rectangle.Line.Color.RGB = System.Drawing.Color.Yellow

            Dim donut As IShape = sheet.Shapes.AddShape(AutoShapeType.Donut, 260, 50, 200, 200)
            donut.Line.DashStyle = LineDashStyle.DashDotDot
            donut.Line.Style = LineStyle.Single
            donut.Line.Weight = 7
            donut.Line.Color.RGB = System.Drawing.Color.Red
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

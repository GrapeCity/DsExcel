Namespace Features.PDFExporting.ExportShape
    Public Class SetShapeNotToPrint
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            ' Add a rectangle
            Dim rectangle As IShape = sheet.Shapes.AddShape(AutoShapeType.Rectangle, 20, 15, 100, 100)

            ' Add an oval
            Dim oval As IShape = sheet.Shapes.AddShape(AutoShapeType.Oval, 160, 15, 100, 100)

            'set oval not to print
            oval.IsPrintable = False
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

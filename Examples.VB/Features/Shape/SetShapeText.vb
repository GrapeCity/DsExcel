Imports System.Drawing

Namespace Features.Shape
    Public Class SetShapeText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim shape As IShape = worksheet.Shapes.AddShape(AutoShapeType.Parallelogram, 1, 1, 200, 100)
            shape.Width = 500
            shape.Height = 200
            shape.TextFrame.TextRange.Font.Color.RGB = Color.FromArgb(0, 255, 0)
            shape.TextFrame.TextRange.Font.Bold = True
            shape.TextFrame.TextRange.Font.Italic = True
            shape.TextFrame.TextRange.Font.Size = 20
            shape.TextFrame.TextRange.Font.Strikethrough = True
            shape.TextFrame.TextRange.Paragraphs.Add("This is a parallelogram shape.")
            shape.TextFrame.TextRange.Paragraphs.Add("My name is XXX")
            shape.TextFrame.TextRange.Paragraphs(1).Runs.Add("Hello World!")
            shape.TextFrame.TextRange.Paragraphs(1).Runs(0).Font.Strikethrough = False
            shape.TextFrame.TextRange.Paragraphs(1).Runs(0).Font.Size = 35
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

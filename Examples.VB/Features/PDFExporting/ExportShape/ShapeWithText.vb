Imports System.Drawing

Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithText
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            ' Add a rectangle
            Dim rectangle As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 50, 30, 300, 200)

            ' Add rich text to rectangle
            rectangle.Fill.Color.RGB = Color.White

            ' Add first paragraph
            Dim run1 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs(0).Runs.Add("         Doc")
            run1.Font.Color.RGB = Color.Tomato
            Dim run2 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs(0).Runs.Add("ume")
            run2.Font.Color.RGB = Color.Yellow
            Dim run3 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs(0).Runs.Add("nts")
            run3.Font.Color.RGB = Color.LightBlue
            Dim paragraph0 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs(0)
            paragraph0.Font.Size = 36
            paragraph0.Font.Bold = True

            rectangle.TextFrame.TextRange.Paragraphs.Add(" ")

            ' Add second paragraph
            Dim paragraph1 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs.Add()
            Dim run4 As ITextRange = paragraph1.Runs.Add("Fast, efficient Excel, Word, and PDF APIs for .NET Standard 2.0 and Java")
            run4.Font.Color.RGB = Color.Black
            run4.Font.Size = 20
            run4.Font.Name = "Agency FB"

            rectangle.TextFrame.TextRange.Paragraphs.Add(" ")

            ' Add third paragraph
            Dim paragraph2 As ITextRange = rectangle.TextFrame.TextRange.Paragraphs.Add()
            Dim run5 As ITextRange = paragraph2.Runs.Add("Take total document control with ultra-fast, low-footprint document APIs for enterprise apps")
            run5.Font.Color.RGB = Color.Black
            run5.Font.Size = 16
            run5.Font.Name = "Times New Roman"
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property

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

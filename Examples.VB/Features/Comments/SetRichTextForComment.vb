Imports System.Drawing

Namespace Features.Comments
    Public Class SetRichTextForComment
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim commentC3 As IComment = worksheet.Range!C3.AddComment("This is a rich text comment:" & vbCrLf)

            'config the paragraph's style.
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Font.Bold = True

            'add runs for the paragraph.
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs.Add("Run1 font size is 15.", 1)
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs.Add("Run2 font strikethrough.", 2)
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs.Add("Run3 font italic, green color.")

            'config the first run of the paragraph's style.
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs(1).Font.Size = 15

            'config the second run of the paragraph's style. 
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs(2).Font.Strikethrough = True

            'config the third run of the paragraph's style. 
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs(3).Font.Italic = True
            commentC3.Shape.TextFrame.TextRange.Paragraphs(0).Runs(3).Font.Color.RGB = Color.Green

            'show comment.
            commentC3.Visible = True
            commentC3.Shape.WidthInPixel = 300
            commentC3.Shape.HeightInPixel = 100
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

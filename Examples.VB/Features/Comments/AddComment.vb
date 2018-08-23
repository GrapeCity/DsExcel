Namespace Features.Comments
    Public Class AddComment
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'create comment for range C3.
            Dim comment As IComment = worksheet.Range!C3.AddComment("Range C3's comment.")

            'change comment's text.
            comment.Text = "Range C3's new comment."
        End Sub
    End Class
End Namespace

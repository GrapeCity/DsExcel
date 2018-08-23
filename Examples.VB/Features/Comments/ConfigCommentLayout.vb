Imports System.Drawing

Namespace Features.Comments
    Public Class ConfigCommentLayout
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim commentC3 As IComment = worksheet.Range!C3.AddComment("Range C3's comment.")

            commentC3.Shape.Line.Color.RGB = Color.LightGreen
            commentC3.Shape.Line.Weight = 3
            commentC3.Shape.Line.Style = LineStyle.ThickThin
            commentC3.Shape.Line.DashStyle = LineDashStyle.Solid
            commentC3.Shape.Fill.Color.RGB = Color.Pink
            commentC3.Shape.Width = 100
            commentC3.Shape.Height = 200
            commentC3.Shape.TextFrame.TextRange.Font.Bold = True
            commentC3.Visible = True
        End Sub
    End Class
End Namespace

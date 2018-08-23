Imports System.Drawing

Namespace Features.Formatting.Fonts
    Public Class FontColor
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "font"
            worksheet.Range!A1.Font.Color = Color.Green
        End Sub
    End Class
End Namespace

Imports System.Drawing

Namespace Features.Formatting.Fill
    Public Class SolidFill
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A2.Interior.Color = Color.Green
        End Sub
    End Class
End Namespace

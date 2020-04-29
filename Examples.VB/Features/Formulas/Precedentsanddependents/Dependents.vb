Imports System.Drawing

Namespace Features.Formulas.Precedentsanddependents
    Public Class Dependents
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1").Value = 100
            worksheet.Range("C1").Formula = "=$A$1"
            worksheet.Range("E1:E5").Formula = "=$A$1"

            For Each item In worksheet.Range("A1").GetDependents()
                item.Interior.Color = Color.Azure
            Next
        End Sub

    End Class
End Namespace

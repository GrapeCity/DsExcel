Imports System.Drawing

Namespace Features.Formulas.Precedentsanddependents
    Public Class Precedents
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            With worksheet.Range
                !E2.Formula = "=sum(A1:A2, B4,C1:C3)"
                !A1.Value = 1
                !A2.Value = 2
                !B4.Value = 3
                !C1.Value = 4
                !C2.Value = 5
                !C3.Value = 6
            End With

            For Each item In worksheet.Range("E2").GetPrecedents()
                item.Interior.Color = Color.Pink
            Next
        End Sub
    End Class
End Namespace

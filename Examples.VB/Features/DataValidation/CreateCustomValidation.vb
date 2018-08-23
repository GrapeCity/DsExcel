Imports System.Drawing

Namespace Features.DataValidation
    Public Class CreateCustomValidation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A2.Value = 1
            worksheet.Range!A3.Value = 2
            worksheet.Range!C2.Value = 0

            'create custom validation, if the expression "=$C$2" result is True, the cell's validation will be True, otherwise, it is False.
            'when use custom validation, validationOperator and formula2 parameters will be ignored even if you have given.
            worksheet.Range("A2:A3").Validation.Add(ValidationType.Custom, ValidationAlertStyle.Information, formula1:="=$C$2")

            'judge if Range("A2:A3") has validation.
            For i As Integer = 1 To 2
                If worksheet.Range(i, 0).HasValidation Then
                    'set the range(i, 0)'s interior color.
                    worksheet.Range(i, 0).Interior.Color = Color.LightBlue
                End If
            Next i
        End Sub
    End Class
End Namespace

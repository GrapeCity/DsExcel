Imports System.Drawing

Namespace Features.DataValidation
    Public Class CreateWholeNumberValidation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("C2:E4").Value = New Object(,)  {
                {1, 3, 5},
                {7, 9, 11},
                {13, 15, 17}
            }

            'add whole number validation.
            worksheet.Range("C2:E4").Validation.Add(ValidationType.Whole, ValidationAlertStyle.Stop, ValidationOperator.Between, 1, 8)
            Dim validation As IValidation = worksheet.Range("C2:E4").Validation
            validation.IgnoreBlank = True
            validation.InputTitle = "Tips"
            validation.InputMessage = "Input a value between 1 and 8, please"
            validation.ErrorTitle = "Error"
            validation.ErrorMessage = "input value does not between 1 and 8"
            validation.ShowInputMessage = True
            validation.ShowError = True

            'judge if Range("C2:E4") has validation.
            For i As Integer = 1 To 3
                For j As Integer = 2 To 4
                    If worksheet.Range(i, j).HasValidation Then
                        'set the range(i, j)'s interior color.
                        worksheet.Range(i, j).Interior.Color = Color.LightBlue
                    End If
                Next j
            Next i
        End Sub
    End Class
End Namespace

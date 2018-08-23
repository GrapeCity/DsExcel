Imports System.Drawing

Namespace Features.DataValidation
    Public Class CreateListValidation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A1.Value = "aaa"
            worksheet.Range!A2.Value = "bbb"
            worksheet.Range!A3.Value = "ccc"
            worksheet.Range("C2:E4").Value = {
                {"aaa", "bbb", "ccc"},
                {"aaa1", "bbb1", "ccc1"},
                {"aaa2", "bbb2", "ccc2"}
            }

            'create list validation.
            worksheet.Range("C2:E4").Validation.Add(ValidationType.List, ValidationAlertStyle.Stop, ValidationOperator.Between, "=$a$1:$a$3")
            Dim validation As IValidation = worksheet.Range("C2:E4").Validation
            validation.InCellDropdown = True

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

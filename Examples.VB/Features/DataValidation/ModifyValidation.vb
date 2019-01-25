Imports System.Drawing

Namespace Features.DataValidation
    Public Class ModifyValidation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("C2:E4").Value = New Object(,)  {
                {
                    New TimeSpan(13, 0, 0),
                    New TimeSpan(13, 29, 59),
                    New TimeSpan(13, 30, 0)
                },
                {
                    New TimeSpan(14, 0, 0),
                    New TimeSpan(15, 0, 0),
                    New TimeSpan(16, 30, 0)
                },
                {
                    New TimeSpan(19, 0, 0),
                    New TimeSpan(18, 29, 59),
                    New TimeSpan(18, 30, 0)
                }
            }

            'create time validation.
            worksheet.Range("C2:E4").Validation.Add(ValidationType.Time, ValidationAlertStyle.Stop, ValidationOperator.Between, New TimeSpan(13, 30, 0), New TimeSpan(18, 30, 0))

            'Modify validation.
            worksheet.Range("C2:E4").Validation.Formula1 = New TimeSpan(14, 30, 0)
            worksheet.Range("C2:E4").Validation.Formula2 = New TimeSpan(19, 0, 0)

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

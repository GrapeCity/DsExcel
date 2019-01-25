Imports System.Drawing

Namespace Features.DataValidation
    Public Class CreateTextLength
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("C2:E4").Value = New Object(,)  {
                {"aa", "bb", "aa1"},
                {"ccc", "dddd", "dddd1"},
                {"ff", "ffff", "ffff1"}
            }

            'create text length validation, text length between 2 and 3.
            worksheet.Range("C2:E4").Validation.Add(ValidationType.TextLength, ValidationAlertStyle.Stop, ValidationOperator.Between, 2, 3)

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

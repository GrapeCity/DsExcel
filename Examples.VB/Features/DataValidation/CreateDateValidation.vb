Imports System.Drawing

Namespace Features.DataValidation
    Public Class CreateDateValidation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("C2:E4").Value = {
                {
                    #12/1/2020#,
                    #12/14/2020#,
                    #12/15/2020#
                },
                {
                    #12/18/2020#,
                    #12/19/2020#,
                    #12/30/2020#
                },
                {
                    #12/31/2020#,
                    #12/13/2019#,
                    #12/15/2019#
                }
            }
            'create date validation.
            worksheet.Range("C2:E4").Validation.Add(ValidationType.Date, ValidationAlertStyle.Stop, ValidationOperator.Between, #12/13/2020#, #12/18/2020#)

            'set column width just for export shown.
            worksheet.Range("C:E").EntireColumn.ColumnWidthInPixel = 120

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

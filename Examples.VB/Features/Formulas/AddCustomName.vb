Imports System.Drawing

Namespace Features.Formulas
    Public Class AddCustomName
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet1 As IWorksheet = workbook.Worksheets(0)

            Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()

            worksheet1.Range!C8.NumberFormat = "0.0000"
            worksheet1.Names.Add("test1", "=Sheet1!$A$1")
            worksheet1.Names.Add("test2", "=Sheet1!test1*2")

            workbook.Names.Add("test3", "=Sheet1!$A$1")
            worksheet1.Range!A1.Value = 1

            'C6's value is 1.
            worksheet1.Range!C6.Formula = "=test1"

            'C7's value is 3.
            worksheet1.Range!C7.Formula = "=test1 + test2"

            'C8's value is 6.283185307
            worksheet1.Range!C8.Formula = "=test2*PI()"

            'judge if Range C6:C8 have formula.
            For i As Integer = 5 To 7
                If worksheet1.Range(i, 2).HasFormula Then
                    worksheet1.Range(i, 2).Interior.Color = Color.LightBlue
                End If
            Next i

            'worksheet1 range A2's value is 1.
            worksheet2.Range!A2.Formula = "=test3"

            'judge if Range A2 has formula.
            If worksheet2.Range!A2.HasFormula Then
                worksheet2.Range!A2.Interior.Color = Color.LightBlue
            End If

            'set r1c1 formula.
            worksheet2.Range!A3.FormulaR1C1 = "=R[-1]C"
        End Sub
    End Class
End Namespace

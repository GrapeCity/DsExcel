Imports System.Drawing

Namespace Features.Formulas
    Public Class UseArrayFormula
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("E4:J5").Value = {
                {1, 2, 3},
                {4, 5, 6}
            }
            worksheet.Range("I6:J8").Value = {
                {2, 2},
                {3, 3},
                {4, 4}
            }

            'O     P      Q
            '2     4      #N/A
            '12    15     #N/A
            '#N/A  #N/A   #N/A
            worksheet.Range("O9:Q11").FormulaArray = "=E4:G5*I6:J8"

            'judge if Range O9 has array formula.
            If worksheet.Range!O9.HasArray Then
                'set O9's entire array's interior color.
                Dim currentarray = worksheet.Range!O9.CurrentArray
                currentarray.Interior.Color = Color.Green
            End If
        End Sub
    End Class
End Namespace

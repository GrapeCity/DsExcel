Namespace Features.Formulas
    Public Class DirtyAndCalculation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range!A1.Value = 1
            worksheet.Range!A2.Formula = "=A1"
            worksheet.Range!A3.Formula = "=SUM(A1, A2)"

            'when get value, calc engine will first calculate and cache the result, then returns the cached result.
            Dim value_A2 = worksheet.Range!A2.Value
            Dim value_A3 = worksheet.Range!A3.Value

            'disable calc engine.
            workbook.EnableCalculation = False

            'Dirty() method will clear the cached value of the workbook.
            workbook.Dirty()

            'Calculate() will not work, because of workbook.EnablCalculation is False.
            workbook.Calculate()

            'it returns 0 because of no cache value exist.
            Dim value_A2_1 = worksheet.Range!A2.Value
            Dim value_A3_1 = worksheet.Range!A3.Value
            worksheet.Range!A1.Value = 2

            'enable calc engine.
            workbook.EnableCalculation = True

            'Dirty() method will clear the cached value of Range A2:A3.
            worksheet.Range("A2:A3").Dirty()

            'Calculate() method will calculate and cache the result, it will return the cache value directly when get value later.
            worksheet.Range("A2:A3").Calculate()

            'it returns cache value directly, does not calculate again.
            Dim value_A2_2 = worksheet.Range!A2.Value
            Dim value_A3_2 = worksheet.Range!A3.Value
        End Sub
    End Class
End Namespace

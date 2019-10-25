Namespace Features.Tables
    Public Class PerformCalculationsInTable
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim data As Object(,) = {
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", #6/8/1968#, "Blue", 67, 165},
                {"Nia", "New York", #7/3/1972#, "Brown", 62, 134},
                {"Jared", "New York", #3/2/1964#, "Hazel", 72, 180},
                {"Natalie", "Washington", #8/8/1972#, "Blue", 66, 163},
                {"Damon", "Washington", #2/2/1986#, "Hazel", 76, 176},
                {"Angela", "Washington", #2/15/1993#, "Brown", 68, 145}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:F7").Value = data
            worksheet.Range("A:F").ColumnWidth = 15

            'add table.
            worksheet.Tables.Add(worksheet.Range("A1:F7"), True)

            '"Name" column's total row calculation function will be "=SUBTOTAL(103,[Name])"
            worksheet.Tables(0).Columns(0).TotalsCalculation = TotalsCalculation.Count

            '"Weight" column's total row calculation function will be "=SUBTOTAL(101,[Weight])"
            worksheet.Tables(0).Columns(4).TotalsCalculation = TotalsCalculation.Average
            worksheet.Tables(0).ShowTotals = True
        End Sub
    End Class
End Namespace

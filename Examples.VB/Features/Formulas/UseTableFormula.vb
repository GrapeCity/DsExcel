Imports System.Drawing

Namespace Features.Formulas
    Public Class UseTableFormula
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A:E").ColumnWidth = 15
            worksheet.Range("A1:E3").Value = {
                {"SalesPerson", "Region", "SalesAmount", "ComPct", "ComAmt"},
                {"Joe", "North", 260, 0.1, Nothing},
                {"Nia", "South", 660, 0.15, Nothing}
            }
            worksheet.Tables.Add(worksheet.Range("A1:E3"), True)
            worksheet.Tables(0).Name = "DeptSales"
            worksheet.Tables(0).Columns("ComPct").DataBodyRange.NumberFormat = "0%"

            'Use table formula in table range.
            worksheet.Tables(0).Columns("ComAmt").DataBodyRange.Formula = "=[@ComPct]*[@SalesAmount]"

            'Use table formula out of table range.
            worksheet.Range!F2.Formula = "=SUM(DeptSales[@SalesAmount])"
            worksheet.Range!G2.Formula = "=SUM(DeptSales[[#Data],[SalesAmount]])"
            worksheet.Range!H2.Formula = "=SUM(DeptSales[SalesAmount])"
            worksheet.Range!I2.Formula = "=SUM(DeptSales[@ComPct], DeptSales[@ComAmt])"

            'judge if Range F2:I2 have formula.
            For i As Integer = 5 To 8
                If worksheet.Range(1, i).HasFormula Then
                    worksheet.Range(1, i).Interior.Color = Color.LightBlue
                End If
            Next i
        End Sub
    End Class
End Namespace

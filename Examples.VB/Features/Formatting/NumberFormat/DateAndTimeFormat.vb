Namespace Features.Formatting.NumberFormat
    Public Class DateAndTimeFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A:F").ColumnWidth = 17
            worksheet.Range("A1:F1").Formula = "= Now()"

            'Apply different date formats.
            worksheet.Range!A1.NumberFormat = "m/d/yy"
            worksheet.Range!B1.NumberFormat = "d-mmm-yy"
            worksheet.Range!C1.NumberFormat = "dddd"

            'Apply different time formats.
            worksheet.Range!D1.NumberFormat = "m/d/yy h:mm"
            worksheet.Range!E1.NumberFormat = "h:mm AM/PM"
            worksheet.Range!F1.NumberFormat = "h:mm:ss"
        End Sub
    End Class
End Namespace

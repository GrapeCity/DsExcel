Namespace Features.RangeOperations
    Public Class SetRowHeightColumnWidth
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'set row height for row 1:2.
            worksheet.Range("1:2").RowHeight = 50

            'set column width for column C:D.
            worksheet.Range("C:D").ColumnWidth = 20
        End Sub
    End Class
End Namespace

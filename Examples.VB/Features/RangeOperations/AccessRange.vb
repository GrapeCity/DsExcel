Imports System.Drawing

Namespace Features.RangeOperations
    Public Class AccessRange
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'use index to access cell A1.
            worksheet.Range(0, 0).Interior.Color = Color.LightGreen

            'use index to access range A1:B2
            worksheet.Range(0, 0, 2, 2).Value = 5

            'use string to access range.
            worksheet.Range!A2.Interior.Color = Color.LightYellow
            worksheet.Range("C3:D4").Interior.Color = Color.Tomato
            worksheet.Range("A5:B7, C3, H5:N6").Value = 2

            'use index to access rows
            worksheet.Rows(2).Interior.Color = Color.LightSalmon

            'use string to access rows
            worksheet.Range("4:4").Interior.Color = Color.LightSkyBlue

            'use index to access columns
            worksheet.Columns(2).Interior.Color = Color.LightSalmon

            'use string to access columns
            worksheet.Range("D:D").Interior.Color = Color.LightSkyBlue

            'use Cells to access range.
            worksheet.Cells(5).Interior.Color = Color.LightBlue
            worksheet.Cells(5, 5).Interior.Color = Color.LightYellow

            'access all rows in worksheet
            Dim allRows = worksheet.Rows.ToString()

            'access all columns in worksheet
            Dim allColumns = worksheet.Columns.ToString()

            'access the entire sheet range
            Dim entireSheet = worksheet.Cells.ToString()
        End Sub
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

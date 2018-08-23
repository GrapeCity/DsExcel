Imports System.Drawing

Namespace Features.RangeOperations
    Public Class AccessCellsRowsColumns
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim range = worksheet.Range("A5:B7")

            'set value for cell A7.
            range.Cells(4).Value = "A7"

            'cell is B6
            range.Cells(1, 1).Value = "B6"

            'row count is 3 and range is A6:B6.
            Dim rowCount = range.Rows.Count
            Dim row = range.Rows(1).ToString()

            'set interior color for row range A6:B6.
            range.Rows(1).Interior.Color = Color.LightBlue

            'column count is 2 and range is B5:B7.
            Dim columnCount = range.Columns.Count
            Dim column = range.Columns(1).ToString()

            'set values for column range B5:B7.
            range.Columns(1).Interior.Color = Color.LightSkyBlue

            'entire rows are from row 5 to row 7
            Dim entirerow = range.EntireRow.ToString()

            'entire columns are from column A to column B
            Dim entireColumn = range.EntireColumn.ToString()
        End Sub
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

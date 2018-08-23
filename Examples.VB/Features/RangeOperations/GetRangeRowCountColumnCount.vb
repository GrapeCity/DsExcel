Namespace Features.RangeOperations
    Public Class GetRangeRowCountColumnCount
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim range = worksheet.Range("A5:B7")

            'cell count is 6.
            Dim cellcount = range.Count

            'cell count is 6.
            Dim cellcount1 = range.Cells.Count

            'row count is 3.
            Dim rowcount = range.Rows.Count

            'column count is 2.
            Dim columncount = range.Columns.Count
        End Sub
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

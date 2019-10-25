Namespace Features.Workbook
    Public Class ImportCsvFileToWorkbookWithOptions
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open csv with custom delimiter.
            Dim options As New CsvOpenOptions With {
                .ColumnSeparator = ",",
                .RowSeparator = vbCrLf,
                .CellSeparator = """"c
            }

            'Change the path to the real file path when open.
            workbook.Open(IO.Path.Combine(CurrentDirectory, "source.csv"), options)
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
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

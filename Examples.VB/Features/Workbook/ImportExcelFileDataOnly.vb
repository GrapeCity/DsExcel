Namespace Features.Workbook
    Public Class ImportExcelFileDataOnly
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Use XlsxOpenOptions.ImportFlags to control what you want to import from excel, ImportFlags.Data means only the data will be imported
            'Change the path to the real file path when open.
            Dim options As New XlsxOpenOptions With {
                .ImportFlags = ImportFlags.Data
            }
            workbook.Open(IO.Path.Combine(CurrentDirectory, "source.xlsx"), options)
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
        Public Overrides ReadOnly Property IsUpdate As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

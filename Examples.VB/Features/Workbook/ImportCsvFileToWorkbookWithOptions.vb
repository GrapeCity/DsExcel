Namespace Features.Workbook
    Public Class ImportCsvFileToWorkbookWithOptions
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open csv with more settings.
            Dim options As New CsvOpenOptions
            options.SeparatorString = "-"

            'Change the path to the real file path when open.
            workbook.Open(CurrentDirectory & "source.csv", options)
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

Namespace Features.Workbook
    Public Class ImportExcelFileWithPassword
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Change the path to the real file path when open.
            Dim options As New XlsxOpenOptions
            options.Password = "123456"
            workbook.Open(CurrentDirectory & "source.xlsx", options)
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

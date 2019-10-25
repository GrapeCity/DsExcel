Namespace Features.Workbook
    Public Class ImportExcelFileToWorkbook
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Change the path to the real file path when open.
            workbook.Open(IO.Path.Combine(CurrentDirectory, "source.xlsx"), OpenFileFormat.Xlsx)
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

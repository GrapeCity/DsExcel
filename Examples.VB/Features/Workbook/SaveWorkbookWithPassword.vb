Namespace Features.Workbook
    Public Class SaveWorkbookWithPassword
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Change the path to real export path when save.
            Dim options As New XlsxSaveOptions With {
                .Password = "123456"
            }
            workbook.Save(IO.Path.Combine(CurrentDirectory, "dest.xlsx"), options)
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

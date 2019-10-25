Namespace Features.Workbook
    Public Class SaveWorkbookToExcelFile
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'change the path to real export path when save.
            workbook.Save(IO.Path.Combine(CurrentDirectory, "dest.xlsx"), SaveFileFormat.Xlsx)
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

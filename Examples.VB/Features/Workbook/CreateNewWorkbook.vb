Namespace Features.Workbook
    Public Class CreateNewWorkbook
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Create empty workbook, contains one worksheet default.
            Dim workbookNew As New Excel.Workbook
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

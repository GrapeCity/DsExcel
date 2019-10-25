Namespace Features.Workbook
    Public Class UnprotectWorkbook
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "GrapeCity Documents for Excel"

            'Protects the workbook so that other users cannot view hidden worksheets, add, move, delete, hidie, or rename worksheets.
            'The protection only happens when you open it with an Excel application.
            workbook.Protect()

            'Removes the above protection from the workbook.
            workbook.Unprotect()
        End Sub

        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return True
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

Namespace Features.Worksheets
    Public Class ProtectWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            'protect worksheet, allow insert column.
            worksheet.Protection = True
            worksheet.ProtectionSettings.AllowInsertingColumns = True
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

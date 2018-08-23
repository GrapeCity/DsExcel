Namespace Features.Worksheets
    Public Class AccessWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Use sheet index to get worksheet.
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Use sheet name to get worksheet.
            Dim worksheet1 As IWorksheet = workbook.Worksheets("Sheet1")
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

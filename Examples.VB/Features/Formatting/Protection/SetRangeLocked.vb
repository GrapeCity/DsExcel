Namespace Features.Formatting.Protection
    Public Class SetRangeLocked
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'config range B1's Locked property.
            worksheet.Range!B1.Locked = False

            'protect worksheet, range B1 can be modified in exported xlsx file.
            worksheet.Protection = True
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace

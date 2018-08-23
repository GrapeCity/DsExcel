Namespace Features.Workbook
    Public Class ImportExcelFileWithoutCalculation
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'When XlsxOpenOptions.DoNotRecalculateAfterOpened means GrapeCity Documents for Excel will just read all the cached values without calculating again after
            'opening an Excel file.
            'Change the path to the real file path when open.
            Dim options As New XlsxOpenOptions
            options.DoNotRecalculateAfterOpened = True
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

Namespace Features.Workbook
    Public Class ImportOleObjectsToWorkbookAndExport
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            workbook.Open(GetResourceStream("OleTemplates.xlsx"))
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String() =
            {"xlsx\OleTemplates.xlsx"}
    End Class
End Namespace

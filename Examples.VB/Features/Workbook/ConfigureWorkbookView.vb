Namespace Features.Workbook
    Public Class ConfigureWorkbookView
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Workbook view settings.
            Dim bookView As IWorkbookView = workbook.BookView
            bookView.DisplayVerticalScrollBar = False
            bookView.DisplayWorkbookTabs = True
            bookView.TabRatio = 0.5
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

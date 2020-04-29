Namespace Features.PDFExporting.ExportShape
    Public Class CustomShape
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("CustomShapes.xlsx")
            workbook.Open(fileStream)
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property

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

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\CustomShapes.xlsx"}
            End Get
        End Property
    End Class
End Namespace

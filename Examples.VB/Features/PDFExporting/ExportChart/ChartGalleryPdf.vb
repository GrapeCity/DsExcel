Namespace Features.PDFExporting.ExportChart
    Public Class ChartGalleryPdf
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\Chart Gallery.xlsx")
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
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Chart Gallery.xlsx"}
            End Get
        End Property
    End Class
End Namespace

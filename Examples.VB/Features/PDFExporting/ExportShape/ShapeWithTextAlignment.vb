Namespace Features.PDFExporting.ExportShape
    Public Class ShapeWithTextAlignment
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\WebsiteFlowChart[Template].xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            For Each item In worksheet.Shapes
                'Centers text vertically. 
                item.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle
                'Centers text horizontally.
                item.TextFrame.HorizontalAnchor = HorizontalAnchor.Center
            Next

        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\\WebsiteFlowChart[Template].xlsx"}
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

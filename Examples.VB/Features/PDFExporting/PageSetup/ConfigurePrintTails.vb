Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigurePrintTails
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("RepeatTails.xlsx")
            workbook.Open(fileStream)

            'Repeat the columns "I" at the left of each page when saving worksheet2 to pdf.
            Dim worksheet1 As IWorksheet = workbook.Worksheets(0)
            worksheet1.PageSetup.PrintTailColumns = "$I:$I"

            'Repeat the row 67th at the bottom of each page when saving worksheet1 to pdf.
            Dim worksheet2 As IWorksheet = workbook.Worksheets(1)
            worksheet2.PageSetup.PrintTailRows = "$50:$50"
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "RepeatTails.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\RepeatTails.xlsx"}
            End Get
        End Property
    End Class
End Namespace

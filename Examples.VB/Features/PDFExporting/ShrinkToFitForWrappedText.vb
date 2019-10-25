Imports System.IO

Namespace Features.PDFExporting
    Public Class ShrinkToFitForWrappedText
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.PageSetup.PrintGridlines = True
            worksheet.PageSetup.PrintHeadings = True

            '"A1" is ordinary wrapped text.
            worksheet.Range!A1.WrapText = True
            worksheet.Range!A1.Value = "GrapeCity Documents for Excel"
            worksheet.Range!A1.RowHeight = 42
            worksheet.Range!A1.ColumnWidth = 9

            'The wrapped text "A2" will be sheink to fit.
            'worksheet.Range["A2"].Interior.Color = Color.LightGreen;
            worksheet.Range!A2.WrapText = True
            worksheet.Range!A2.ShrinkToFit = True
            worksheet.Range!A2.Value = "GrapeCity Documents for Excel"
            worksheet.Range!A2.RowHeight = 32

            'You must create a pdfSaveOptions object before using ShrinkToFitSettings.
            Dim pdfSaveOptions As New PdfSaveOptions

            'Shrink the wrapped text within the cell with existing row height/column width, while exporting to PDF. 
            pdfSaveOptions.ShrinkToFitSettings.CanShrinkToFitWrappedText = True

            'Save the workbook into pdf file.
            workbook.Save(outputStream, pdfSaveOptions)

            'Set minimum font size with which the text should shrink.
            'pdfSaveOptions.ShrinkToFitSettings.MinimumFont = 10;
            'If after setting the minimum font size, the text is very long not fully visible, the ellipsis string to show for long text.
            'pdfSaveOptions.ShrinkToFitSettings.Ellipsis = "~";

            'Save the workbook into pdf file.
            workbook.Save(outputStream, pdfSaveOptions)
        End Sub

        Public Overrides ReadOnly Property SavePageInfos As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

Imports System.IO

Namespace Features.PDFExporting
    Public Class SetPickTrayByPDFSizeOptionToPDF
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1").Value = "GrapeCity Documents for Excel"
            worksheet.Range("A1").Font.Size = 25

            'You must create a pdfSaveOptions object before using ViewerPreferences.
            Dim pdfSaveOptions As New PdfSaveOptions

            'The PDF page size is used to select the input paper tray when printing.
            pdfSaveOptions.ViewerPreferences.PickTrayByPDFSize = True

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
    End Class
End Namespace

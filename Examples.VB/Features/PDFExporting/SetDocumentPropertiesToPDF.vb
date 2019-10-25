Imports System.IO

Namespace Features.PDFExporting
    Public Class SetDocumentPropertiesToPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "GrapeCity Documents for Excel"
            worksheet.Range!A1.Font.Size = 25

            Dim documentProperties As New DocumentProperties
            With documentProperties
                'Sets the name of the person that created the PDF document.
                .Author = "Jaime Smith"
                'Sets the title of the  PDF document.
                .Title = "GcPdf Document Info Sample"
                'Do not embed a font.
                .EmbedStandardWindowsFonts = False
                'Set the PDF version.
                .PdfVersion = 1.5F
                'Set the subject of the PDF document.
                .Subject = "GcPdfDocument.DocumentInfo"
                'Set the keyword associated with the PDF document.
                .Keywords = "Keyword1"
                'Set the creation date and time of the PDF document.
                .CreationDate = Date.Now.AddYears(10)
                'Set the date and time the PDF document was most recently modified.
                .ModifyDate = Date.Now.AddYears(11)
                'Set the name of the application that created the original PDF document.
                .Creator = "GcPdfWeb Creator"
                'Set the name of the application that created the PDF document.
                .Producer = "GcPdfWeb Producer"
            End With

            Dim pdfSaveOptions As New PdfSaveOptions
            'Sets the document properties of the pdf.
            pdfSaveOptions.DocumentProperties = documentProperties

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

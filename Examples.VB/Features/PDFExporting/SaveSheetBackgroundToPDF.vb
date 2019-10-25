Imports System.IO

Namespace Features.PDFExporting
    Public Class SaveSheetBackgroundToPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1").Value = "GrapeCity Documents for Excel"
            worksheet.Range("A1").Font.Size = 25

            Dim stream As Stream = GetResourceStream("logo.png")
            Dim imageBytes(CInt(stream.Length) - 1) As Byte
            stream.Read(imageBytes, 0, imageBytes.Length)
            'Set a background image for worksheet
            worksheet.BackgroundPicture = imageBytes

            Dim pdfSaveOptions As New PdfSaveOptions
            'Print the background image when saving pdf.
            'The background image will be centered on every page of the sheet.
            pdfSaveOptions.PrintBackgroundPicture = True

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

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"logo.png"}
            End Get
        End Property
    End Class
End Namespace

Imports System.IO

Namespace Features.PDFExporting
    Public Class SetSecurityOptionsToPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Value = "GrapeCity Documents for Excel"
            worksheet.Range!A1.Font.Size = 25

            'The security settings of pdf when converting excel to pdf.
            Dim securityOptions As New PdfSecurityOptions
            With securityOptions
                'Sets the user password.
                .UserPassword = "user"
                'Sets the owner password.
                .OwnerPassword = "owner"
                'Allow to print pdf document.
                .PrintPermission = True
                'Print the pdf document in high quality.
                .FullQualityPrintPermission = True
                'Allow to copy or extract the content of the pdf document.
                .ExtractContentPermission = True
                'Allow to modify the pdf document.
                .ModifyDocumentPermission = True
                'Allow to insert, rotate, or delete pages and create bookmarks or thumbnail images of the pdf document.
                .AssembleDocumentPermission = True
                'Allow to modify text annotations and fill the form fields of the pdf document.
                .ModifyAnnotationsPermission = True
                'Filling the form fields of the pdf document is not allowed.
                .FillFormsPermission = False
            End With

            Dim pdfSaveOptions As New PdfSaveOptions
            'Sets the secutity settings of the pdf.
            pdfSaveOptions.SecurityOptions = securityOptions

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

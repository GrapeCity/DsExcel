Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class PrintMultipleWorksheetsToOnePage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("Multiple sheets one page.xlsx")
            workbook.Open(fileStream)

            'NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            'Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            ' This page will save datas for multiple pages.
            Dim page As Pdf.Page = doc.NewPage()

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the pagination information of the workbook.
            Dim pages As IList(Of PageInfo) = printManager.Paginate(workbook)

            'Divide the multiple pages into 1 rows and 2 columns and printed them on one page.
            printManager.Draw(page, pages, 1, 2)

            'Save the document into pdf file.
            doc.Save(outputStream)
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Multiple sheets one page.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Multiple sheets one page.xlsx"}
            End Get
        End Property
    End Class
End Namespace

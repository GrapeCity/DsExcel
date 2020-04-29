Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class CustomPageInfos
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("KeepTogether.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Firstly, create a printManager.
            Dim printManager As New Excel.PrintManager

            'Get the natural pagination information of the worksheet.
            'The first page of the natural pagination is "A1:F37", the second page is from row "A38:F73" 
            Dim pages As IList(Of PageInfo) = printManager.Paginate(worksheet)

            'Custom the pageInfos.
            pages(0).PageContent.Range = worksheet.Range("A1:F36") ' The first page is "A1:F36".
            pages(0).PageSettings.CenterHeader = "&KFF0000&20 Budget summary report" ' The center header of the first page will show the text "Budget summary report".
            pages(0).PageSettings.CenterFooter = "&KFF0000&20 Page &P" ' The center footer of the first page will show the page number "1".
            pages(1).PageContent.Range = worksheet.Range("A37:F73") ' The second page is "A37:F73".

            'Save the modified pages into pdf file.
            printManager.SavePDF(outputStream, pages)
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "KeepTogether.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\KeepTogether.xlsx"}
            End Get
        End Property
    End Class
End Namespace

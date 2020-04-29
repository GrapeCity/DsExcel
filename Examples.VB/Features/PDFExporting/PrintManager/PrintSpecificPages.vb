Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class PrintSpecificPages
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("PrintSpecificPDFPages.xlsx")
            workbook.Open(fileStream)

            'Firstly, create a printManager.
            Dim printManager As New Excel.PrintManager

            'Get the natural pagination information of the workbook.
            Dim pages As IList(Of PageInfo) = printManager.Paginate(workbook)

            'Pick some pages to print.
            Dim newPages As IList(Of PageInfo) = New List(Of PageInfo)
            newPages.Add(pages(0))
            newPages.Add(pages(2))

            'Update the page number and the page settings of each page. The page number is continuous.
            printManager.UpdatePageNumberAndPageSettings(newPages)

            'Save the pages into pdf file.
            printManager.SavePDF(outputStream, newPages)
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
                Return "PrintSpecificPDFPages.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\PrintSpecificPDFPages.xlsx"}
            End Get
        End Property
    End Class
End Namespace

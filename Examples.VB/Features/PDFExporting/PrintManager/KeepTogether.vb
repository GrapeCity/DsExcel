Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class KeepTogether
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("KeepTogether.xlsx")
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'The first page of the natural pagination is from row 1th to 37th, the second page is from row 38th to 73th. 
            Dim keepTogetherRanges As IList(Of IRange) = New List(Of IRange)
            'The row 37th and 38th need to keep together. So the pagination results are: the first page if from row 1th to 36th, the second page is from row 37th to 73th.
            keepTogetherRanges.Add(worksheet.Range("37:38"))

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the pagination information of the worksheet.
            Dim pages As IList(Of PageInfo) = printManager.Paginate(worksheet, keepTogetherRanges, Nothing)

            'Save the pages into pdf file.
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

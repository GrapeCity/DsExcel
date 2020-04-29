Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class GetPaginationInfo
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(1)

            Dim printManager As New Excel.PrintManager

            'The columnIndexs is [4, 12, 20], this means that the horizontal direction is split after the column 5th, 13th, and 21th. 
            Dim columnIndexs As IList(Of Integer) = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Horizontal)
            'The rowIndexs is [42, 61], this means that the vertical direction is split after the row 43th and 62th.
            Dim rowIndexs As IList(Of Integer) = printManager.GetPaginationInfo(worksheet, PaginationOrientation.Vertical)

            'Save the modified pages into pdf file.
            Dim pages As IList(Of PageInfo) = printManager.Paginate(worksheet)
            printManager.SavePDF(outputStream, pages)
        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\\Medical office start-up expenses 1.xlsx"}
            End Get
        End Property

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

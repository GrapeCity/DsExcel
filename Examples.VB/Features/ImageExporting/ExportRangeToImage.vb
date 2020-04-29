Imports System.IO

Namespace Features.ImageExporting
    Public Class ExportRangeToImage
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("Personal Monthly Budget.xlsx")

            ' Open a xlsx file
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Export range "B14:E25" to image
            worksheet.Range("B14:E25").ToImage(outputStream, ImageType.PNG)
        End Sub
        Public Overrides ReadOnly Property SaveAsImages As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Personal Monthly Budget.xlsx"}
            End Get
        End Property
    End Class
End Namespace

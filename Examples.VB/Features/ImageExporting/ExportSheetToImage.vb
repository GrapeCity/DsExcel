Imports System.IO

Namespace Features.ImageExporting
    Public Class ExportSheetToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream = GetResourceStream("xlsx\Home inventory.xlsx")
            workbook.Open(fileStream)

            Dim worksheet = workbook.Worksheets(0)

            ' Save the worksheet as image to a stream.
            worksheet.ToImage(outputStream, ImageType.PNG)
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Home inventory.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Home inventory.xlsx"}
            End Get
        End Property
    End Class
End Namespace

Imports System.IO

Namespace Features.ImageExporting
    Public Class ExportShapeToImage
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("GroupShape.xlsx")

            ' Open a xlsx file contains a group shape
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Export the shape to image
            worksheet.Shapes(0).ToImage(outputStream, ImageType.PNG)
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
                Return New String() {"xlsx\GroupShape.xlsx"}
            End Get
        End Property
    End Class
End Namespace


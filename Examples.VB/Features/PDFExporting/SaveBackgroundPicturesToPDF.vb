Imports System.IO

Namespace Features.PDFExporting
    Public Class SaveBackgroundPicturesToPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\To_Do_List.xlsx")
            workbook.Open(fileStream)

            Dim worksheet = workbook.Worksheets(0)

            Dim stream = GetResourceStream("AcmeLogo.png")
            Dim imageBytes(CInt(stream.Length) - 1) As Byte
            stream.Read(imageBytes, 0, imageBytes.Length)

            'Add a background picture for the worksheet, and the background picture will be rendered into the destination rectangle[10, 10, 500, 370].
            Dim picture As IBackgroundPicture = worksheet.BackgroundPictures.AddPictureInPixel(stream, ImageType.PNG, 10, 10, 150, 100)

            'The background picture will be resized to fill the destination dimensions.
            picture.BackgroundImageLayout = ImageLayout.Tile

            'Sets the transparency of the background pictures.
            picture.Transparency = 0.5
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\\To_Do_List.xlsx", "AcmeLogo.png"}
            End Get
        End Property
    End Class
End Namespace

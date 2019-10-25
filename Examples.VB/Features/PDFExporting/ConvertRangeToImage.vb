Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting
    Public Class ConvertRangeToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set value
            worksheet.Range("A1:C1").Value = {
                "Device", "Quantity", "Unit Price"
            }
            worksheet.Range("A2:C5").Value = New Object(,) {
                {"T540p", 12, 9850},
                {"T570", 5, 7460},
                {"Y460", 6, 5400},
                {"Y460F", 8, 6240}
            }

            'Set style
            With worksheet.Range("A1:C1")
                .Font.Bold = True
                .Font.Color = Color.White
                .Interior.Color = Color.LightBlue
            End With
            With worksheet.Range("A2:C5").Borders(BordersIndex.InsideHorizontal)
                .Color = Color.Orange
                .LineStyle = BorderLineStyle.DashDot
            End With

            'NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            'Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            Dim page As Pdf.Page = doc.NewPage()

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the size of the range"A1:C5".
            Dim size As SizeF = printManager.GetSize(worksheet.Range("A1:C5"))

            'Modify the size of the page.
            page.Size = size

            ' Draw the Range"A1:E5" to the specified location on the page. 
            printManager.Draw(page, New PointF(0, 0), worksheet.Range("A1:C5"))

            ' Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, New Pdf.SaveAsImageOptions() With {.Resolution = 72})
        End Sub

        Public Overrides ReadOnly Property SaveAsImages As Boolean
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

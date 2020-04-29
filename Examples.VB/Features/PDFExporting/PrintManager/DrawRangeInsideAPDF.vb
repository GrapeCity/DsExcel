Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class DrawRangeInsideAPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Set value
            worksheet.Range("A4:C4").Value = {"Device", "Quantity", "Unit Price"}
            worksheet.Range("A5:C8").Value = New Object(,) {
                {"T540p", 12, 9850},
                {"T570", 5, 7460},
                {"Y460", 6, 5400},
                {"Y460F", 8, 6240}
            }

            ' Set style
            With worksheet.Range("A4:C4")
                .Font.Bold = True
                .Font.Color = Color.White
                .Interior.Color = Color.LightBlue
            End With

            With worksheet.Range("A5:C8").Borders(BordersIndex.InsideHorizontal)
                .Color = Color.Orange
                .LineStyle = BorderLineStyle.DashDot
            End With

            ' NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            ' Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            Dim page As Pdf.Page = doc.NewPage()
            Dim g As Pdf.GcPdfGraphics = page.Graphics

            ' Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            ' Draw the Range"A4:C8" to the specified location on the page. 
            printManager.Draw(page, New PointF(30, 100), worksheet.Range("A4:C8"))

            ' Save the modified pages into pdf file.
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
    End Class
End Namespace

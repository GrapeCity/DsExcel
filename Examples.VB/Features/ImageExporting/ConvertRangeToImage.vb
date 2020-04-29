Imports System.Drawing
Imports System.IO

Namespace Features.ImageExporting
    Public Class ConvertRangeToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set value
            worksheet.Range("A1:C1").Value = New String() {"Device", "Quantity", "Unit Price"}
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

            ' Save the range "A1:C5" as an image to a stream.
            worksheet.Range("A1:C5").ToImage(outputStream, ImageType.PNG)
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
    End Class
End Namespace

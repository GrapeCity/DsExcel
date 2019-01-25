Imports System.Drawing
Imports System.IO

Namespace Features.Worksheets
    Public Class SetBackgroundImage
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim stream As Stream = GetResourceStream("logo.png")
            Dim imageBytes As Byte() = New Byte(stream.Length - 1) {}
            stream.Read(imageBytes, 0, imageBytes.Length)
            worksheet.BackgroundPicture = imageBytes
        End Sub
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

    End Class
End Namespace

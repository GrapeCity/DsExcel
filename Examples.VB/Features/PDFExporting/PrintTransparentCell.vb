Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting
    Public Class PrintTransparentCell
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Initialize worksheet's values.
            worksheet.Range("A1").Value = "Info from Acme Institute of Health:"
            worksheet.Range("B2").Value = "BLOOD PRESSURE TRACKER"
            worksheet.Range("B4:F13").Value = New Object(,) {
                {"NAME", Nothing, Nothing, Nothing, "JAMES HILL"},
                {Nothing, Nothing, Nothing, Nothing, Nothing},
                {Nothing, Nothing, Nothing, "Systolic", "Diastolic"},
                {"TARGET BLOOD PRESSURE", Nothing, Nothing, 120, 80},
                {Nothing, Nothing, Nothing, Nothing, Nothing},
                {Nothing, Nothing, Nothing, "Systolic", "Diastolic"},
                {"CALL PHYSICIAN IF ABOVE", Nothing, Nothing, 140, 90},
                {Nothing, Nothing, Nothing, Nothing, Nothing},
                {Nothing, Nothing, Nothing, Nothing, Nothing},
                {"PHYSICIAN PHONE NUMBER", Nothing, Nothing, "(001))5104234242", Nothing}
            }
            worksheet.Range("A1").Font.Size = 25

            'Set row height.
            worksheet.StandardHeight = 12.75
            worksheet.StandardWidth = 8.43
            worksheet.Rows(1).RowHeight = 32.25
            worksheet.Rows(2).RowHeight = 13.5
            worksheet.Rows(3).RowHeight = 18.75
            worksheet.Rows(6).RowHeight = 18.75
            worksheet.Rows(9).RowHeight = 18.75
            worksheet.Rows(12).RowHeight = 18.75
            worksheet.Rows(15).RowHeight = 19.5
            worksheet.Rows(16).RowHeight = 13.5
            worksheet.Rows(33).RowHeight = 19.5
            worksheet.Rows(34).RowHeight = 13.5

            'Set column width.
            worksheet.Columns(0).ColumnWidth = 1.7109375
            worksheet.Columns(1).ColumnWidth = 12.140625
            worksheet.Columns(2).ColumnWidth = 12.140625
            worksheet.Columns(3).ColumnWidth = 12.140625
            worksheet.Columns(4).ColumnWidth = 11.85546875
            worksheet.Columns(5).ColumnWidth = 12.7109375

            'Set the transparency value of the background color of range("A1:G20") to 50.
            worksheet.Range("A1:G20").Interior.Color = Color.FromArgb(50, 255, 0, 0)

            'Add a background picture.
            Dim stream As Stream = GetResourceStream("AcmeLogo-vertical-250px.jpg")
            Dim imageBytes(CInt(stream.Length) - 1) As Byte
            stream.Read(imageBytes, 0, imageBytes.Length)
            Dim picture As IBackgroundPicture = worksheet.BackgroundPictures.AddPictureInPixel(stream, ImageType.JPG, 10, 10, 300, 150)

            'You must create a pdfSaveOptions object before using PrintTransparentCell.
            Dim pdfSaveOptions As New PdfSaveOptions
            'Set print the transparency of cell's background color, so the background picture will come out in the back.
            pdfSaveOptions.PrintTransparentCell = True

            'Save the workbook into pdf file.
            workbook.Save(outputStream, pdfSaveOptions)
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

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"AcmeLogo-vertical-250px.jpg"}
            End Get
        End Property
    End Class
End Namespace

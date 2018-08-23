Namespace Features.Picture
    Public Class PictureCut
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim stream As IO.Stream = GetResourceStream("logo.png")

            'Create a shape in worksheet, picture's range is Range("A2:I6")
            Dim picture As IShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 395, 60)

            'Range("A2:I6") must contain picture's range, cut a new picture to Range("J2:R6")
            worksheet.Range("A2:I6").Cut(worksheet.Range!J2)
            'worksheet.Range("A2:I6").Cut(worksheet.Range("J2:R6"))
            'Cross sheet cut, cut a new shape to worksheet2's Range("J2:R6")
            'Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            'worksheet.Range("A2:I6").Cut(worksheet2.Range!J2)
            'worksheet.Range("A2:I6").Cut(worksheet2.Range("J2:R6"))
        End Sub
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

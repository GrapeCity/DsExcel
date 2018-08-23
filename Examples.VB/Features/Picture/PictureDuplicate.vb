Namespace Features.Picture
    Public Class PictureDuplicate
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim stream As IO.Stream = GetResourceStream("logo.png")

            'Create a shape in worksheet
            Dim picture As IShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 395, 60)

            'Duplicate picture
            Dim newPicture As IShape = picture.Duplicate()
        End Sub
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

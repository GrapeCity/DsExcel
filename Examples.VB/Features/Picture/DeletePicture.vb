Namespace Features.Picture
    Public Class DeletePicture
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim stream As IO.Stream = GetResourceStream("logo.png")
            Dim picture As IShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 100, 100)

            'set picture size.
            picture.WidthInPixel = 700
            picture.HeightInPixel = 120

            'config picture layout.
            picture.Fill.Solid()
            picture.Fill.Color.ObjectThemeColor = ThemeColor.Accent1

            'delete picture.
            picture.Delete()
        End Sub
    End Class
End Namespace

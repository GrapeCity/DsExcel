Namespace Features.Picture
    Public Class CreatePicture
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            Dim picture As IShape = worksheet.Shapes.AddPicture(stream, ImageType.PNG, 20, 20, 790, 120)
        End Sub
    End Class
End Namespace

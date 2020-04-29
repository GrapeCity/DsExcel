Namespace Features.RangeOperations
    Public Class GetRangeBoundary
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\Sport sign-up sheet.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Get the absolute location And size of the Range["G1"] in the worksheet.
            Dim range As IRange = worksheet.Range("G1")
            Dim rect As System.Drawing.Rectangle = Excel.CellInfo.GetRangeBoundary(range)
            'Add the image to the Range["G1"]
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            worksheet.Shapes.AddPictureInPixel(stream, ImageType.PNG, rect.X, rect.Y, rect.Width, rect.Height)
        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\\Sport sign-up sheet.xlsx", "logo.png"}
            End Get
        End Property
    End Class
End Namespace

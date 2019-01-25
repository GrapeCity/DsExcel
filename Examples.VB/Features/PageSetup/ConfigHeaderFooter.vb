Namespace Features.PageSetup
    Public Class ConfigHeaderFooter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("PageSetup Demo.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set page headerfooter
            worksheet.PageSetup.LeftHeader = "&""Arial,Italic""LeftHeader"
            worksheet.PageSetup.CenterHeader = "&P"

            'Set page headerfooter's graphic
            worksheet.PageSetup.CenterFooter = "&G"
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            worksheet.PageSetup.CenterFooterPicture.SetGraphicStream(stream, ImageType.PNG)
            'If you have picture resources locally, you can also set graphic in this way.
            'worksheet.PageSetup.CenterFooter = "&G"
            'worksheet.PageSetup.CenterFooterPicture.Filename = "C:\picture.png"
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace

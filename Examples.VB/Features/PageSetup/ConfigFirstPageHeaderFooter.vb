Namespace Features.PageSetup
    Public Class ConfigFirstPageHeaderFooter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set first page headerfooter
            worksheet.PageSetup.DifferentFirstPageHeaderFooter = True
            worksheet.PageSetup.FirstPage.CenterHeader.Text = "&T"
            worksheet.PageSetup.FirstPage.RightFooter.Text = "&D"

            'Set first page headerfooter's graphic
            worksheet.PageSetup.FirstPage.LeftFooter.Text = "&G"

            Dim stream As IO.Stream = GetResourceStream("logo.png")
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.SetGraphicStream(stream, ImageType.PNG)
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.Width = 100
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.Height = 13
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace

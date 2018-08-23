Namespace Features.PageSetup
    Public Class ConfigEvenPageHeaderFooter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetTemplateStream()
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set even page headerfooter
            worksheet.PageSetup.OddAndEvenPagesHeaderFooter = True
            worksheet.PageSetup.EvenPage.CenterHeader.Text = "&T"
            worksheet.PageSetup.EvenPage.RightFooter.Text = "&D"

            'Set even page headerfooter's graphic
            worksheet.PageSetup.EvenPage.LeftFooter.Text = "&G"
            Dim stream As IO.Stream = GetResourceStream("logo.png")
            worksheet.PageSetup.EvenPage.LeftFooter.Picture.SetGraphicStream(stream, ImageType.PNG)
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace

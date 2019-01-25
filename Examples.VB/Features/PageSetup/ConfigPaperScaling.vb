Namespace Features.PageSetup
    Public Class ConfigPaperScaling
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("PageSetup Demo.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Set paper scaling
            'Method 1: Set percent scale 
            worksheet.PageSetup.IsPercentScale = True
            worksheet.PageSetup.Zoom = 150
            'Or Method 2: Fit to page's wide & tall
            'worksheet.PageSetup.IsPercentScale = False
            'worksheet.PageSetup.FitToPagesWide = 3
            'worksheet.PageSetup.FitToPagesTall = 4
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "PageSetup Demo.xlsx"
            End Get
        End Property
    End Class
End Namespace

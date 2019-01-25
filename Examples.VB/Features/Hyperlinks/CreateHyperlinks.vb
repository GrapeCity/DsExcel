Namespace Features.Hyperlinks
    Public Class CreateHyperlinks
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A:A").ColumnWidth = 30

            'add a hyperlink link to web page.
            worksheet.Range("A1:B2").Hyperlinks.Add(worksheet.Range!A1, "http://www.google.com/", Nothing, "open google web site.", "Google")

            'add a hyperlink link to a range in this document.
            worksheet.Range("A3:B4").Hyperlinks.Add(worksheet.Range!A3, Nothing, "Sheet1!$C$3:$E$4", "Go to sheet1 C3:E4")

            'add a hyperlink link to email address.
            worksheet.Range("A5:B6").Hyperlinks.Add(worksheet.Range!A5, "mailto:us.sales@grapecity.com", Nothing, "Send an email to sales", "Send an email to sales")

            'add a hyperlink link to external file.
            'change the path to real picture file path.
            Dim path As String = "external.xlsx"
            worksheet.Range("A7:B8").Hyperlinks.Add(worksheet.Range!A7, path, Nothing, "link to external.xlsx file.", "External.xlsx")
        End Sub
    End Class
End Namespace

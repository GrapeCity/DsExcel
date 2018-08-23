Namespace Features.Theme
    Public Class ChangeWorkbookTheme
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Change workbook's theme to builtin theme.
            workbook.Theme = Themes.Berlin

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!B2.Value = "Major Font:"
            worksheet.Range!B3.Value = "Minor Font:"
            worksheet.Range!C2.Value = "Trebuchet MS"
            worksheet.Range!C3.Value = "Trebuchet MS"

            worksheet.Range!C2.Font.ThemeFont = ThemeFont.Major
            worksheet.Range!C3.Font.ThemeFont = ThemeFont.Minor

            worksheet.Range("E2:E13").Value = {"Light1", "Dark1", "Light2", "Dark2", "Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6", "Hyperlink", "FollowedHyperlink"}

            worksheet.Range!F2.Interior.ThemeColor = ThemeColor.Light1
            worksheet.Range!F3.Interior.ThemeColor = ThemeColor.Dark1
            worksheet.Range!F4.Interior.ThemeColor = ThemeColor.Light2
            worksheet.Range!F5.Interior.ThemeColor = ThemeColor.Dark2
            worksheet.Range!F6.Interior.ThemeColor = ThemeColor.Accent1
            worksheet.Range!F7.Interior.ThemeColor = ThemeColor.Accent2
            worksheet.Range!F8.Interior.ThemeColor = ThemeColor.Accent3
            worksheet.Range!F9.Interior.ThemeColor = ThemeColor.Accent4

            worksheet.Range("F10").Interior.ThemeColor = ThemeColor.Accent5
            worksheet.Range("F11").Interior.ThemeColor = ThemeColor.Accent6
            worksheet.Range("F12").Interior.ThemeColor = ThemeColor.Hyperlink
            worksheet.Range("F13").Interior.ThemeColor = ThemeColor.FollowedHyperlink

            worksheet.Range("B:F").ColumnWidth = 15
        End Sub
    End Class
End Namespace

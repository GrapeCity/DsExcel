Namespace Features.CellRichText
    Public Class AddTextRun
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim b2 As IRange = worksheet.Range!B2

            Dim run1 As ITextRun = b2.RichText.Add("GrapeCity")
            run1.Font.Name = "Agency FB"
            run1.Font.Size = 26
            run1.Font.ThemeColor = ThemeColor.Accent1
            run1.Font.Bold = True

            Dim run2 As ITextRun = b2.RichText.Add(" Documents")
            run2.Font.ThemeColor = ThemeColor.Accent2
            run2.Font.Name = "Arial Black"
            run2.Font.Size = 20

            Dim run3 As ITextRun = b2.RichText.Add(" for ")
            run3.Font.Italic = True

            Dim run4 As ITextRun = b2.RichText.Add("Excel")
            run4.Font.Color = System.Drawing.Color.Blue
            run4.Font.Bold = True
            run4.Font.Size = 26
            run4.Font.Underline = UnderlineType.Double
            b2.EntireRow.RowHeight = 42
        End Sub
    End Class
End Namespace

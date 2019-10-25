Namespace Features.CellRichText
    Public Class CharacterCellText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim b2 As IRange = worksheet.Range!B2
            b2.Value = "GrapeCity Documents for Excel"
            b2.Font.Size = 26
            b2.EntireRow.RowHeight = 42

            Dim run1 As ITextRun = b2.Characters(0, 9)
            run1.Font.Name = "Agency FB"
            run1.Font.ThemeColor = ThemeColor.Accent1
            run1.Font.Bold = True

            Dim run2 As ITextRun = b2.Characters(10, 9)
            run2.Font.ThemeColor = ThemeColor.Accent2
            run2.Font.Name = "Arial Black"
            run2.Font.Underline = UnderlineType.Single

            Dim run3 As ITextRun = b2.Characters(20, 3)
            run3.Font.Italic = True

            Dim run4 As ITextRun = b2.Characters(24, 5)
            run4.Font.Color = System.Drawing.Color.Blue
            run4.Font.Bold = True
        End Sub
    End Class
End Namespace

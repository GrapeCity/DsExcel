Namespace Features.CellRichText
    Public Class ConfigRunFont
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim a2 As IRange = worksheet.Range!A2
            a2.Font.Size = 18
            a2.Font.Bold = True
            a2.VerticalAlignment = VerticalAlignment.Center
            a2.EntireRow.RowHeight = 42
            a2.EntireColumn.ColumnWidth = 40

            a2.Value = "Perfect square trinomial"
            Dim run1 As ITextRun = a2.Characters(8, 7)
            run1.Font.Italic = True
            run1.Font.ThemeColor = ThemeColor.Accent1

            Dim b2 As IRange = worksheet.Range!B2
            b2.Font.Size = 26
            b2.EntireColumn.ColumnWidth = 60

            b2.Value = "(a+b)2 = a2+2ab+b2"
            Dim superRun1 As ITextRun = b2.Characters(5, 1)
            superRun1.Font.Superscript = True
            superRun1.Font.Color = System.Drawing.Color.Red

            Dim superRun2 As ITextRun = b2.Characters(10, 1)
            superRun2.Font.Superscript = True
            superRun2.Font.Color = System.Drawing.Color.Green

            Dim superRun3 As ITextRun = b2.Characters(17, 1)
            superRun3.Font.Superscript = True
            superRun3.Font.Color = System.Drawing.Color.Blue
        End Sub
    End Class
End Namespace

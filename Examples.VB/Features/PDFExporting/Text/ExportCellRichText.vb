Namespace Features.PDFExporting.Text
    Public Class ExportCellRichText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim a1 As IRange = worksheet.Range!A1
            a1.Value = "Perfect square trinomial"

            a1.Font.Size = 26
            a1.Font.Bold = True

            a1.VerticalAlignment = VerticalAlignment.Bottom
            a1.EntireRow.RowHeight = 42
            a1.EntireColumn.ColumnWidth = 50

            Dim run1 As ITextRun = a1.Characters(8, 7)
            run1.Font.Italic = True
            run1.Font.ThemeColor = ThemeColor.Accent1

            Dim b1 As IRange = worksheet.Range!B1
            b1.Font.Size = 22
            b1.EntireColumn.ColumnWidth = 40
            b1.Value = "(a+b)2 = a2+2ab+b2"
            b1.VerticalAlignment = VerticalAlignment.Center

            Dim superRun1 As ITextRun = b1.Characters(5, 1)
            superRun1.Font.Superscript = True
            superRun1.Font.Color = System.Drawing.Color.Red

            Dim superRun2 As ITextRun = b1.Characters(10, 1)
            superRun2.Font.Superscript = True
            superRun2.Font.Color = System.Drawing.Color.Green

            Dim superRun3 As ITextRun = b1.Characters(17, 1)
            superRun3.Font.Superscript = True
            superRun3.Font.Color = System.Drawing.Color.Blue
        End Sub
    End Class
End Namespace

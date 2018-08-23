Imports System.Drawing

Namespace Features.Formatting
    Public Class ApplyStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Change to build in name style.
            worksheet.Range!A1.Value = "Bad"
            worksheet.Range!A1.Style = workbook.Styles("Bad")

            'Change to custom name style.
            'Add custom name style.
            Dim style As IStyle = workbook.Styles.Add("testStyle")

            'Config custom name style settings begin.

            'Border
            style.Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thin
            style.Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thick
            style.Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Double
            style.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Double
            style.Borders.Color = Color.FromArgb(0, 255, 0)

            'Font
            style.Font.ThemeColor = ThemeColor.Accent1
            style.Font.TintAndShade = 0.8
            style.Font.Italic = True
            style.Font.Bold = True
            style.Font.Name = "LiSu"
            style.Font.Size = 28
            style.Font.Strikethrough = True
            style.Font.Subscript = True
            style.Font.Superscript = False
            style.Font.Underline = UnderlineType.Double

            'Protection
            style.FormulaHidden = True
            style.Locked = False

            'Number
            style.NumberFormat = "#,##0_);[Red](#,##0)"

            'Alignment
            style.HorizontalAlignment = HorizontalAlignment.Right
            style.VerticalAlignment = VerticalAlignment.Bottom
            style.WrapText = True
            style.IndentLevel = 5
            style.Orientation = 45

            'Fill
            style.Interior.ColorIndex = 5
            style.Interior.Pattern = Pattern.Down
            style.Interior.PatternColor = Color.FromArgb(0, 0, 255)
            style.IncludeAlignment = False
            style.IncludeBorder = True
            style.IncludeFont = False
            style.IncludeNumber = True
            style.IncludePatterns = False
            style.IncludeProtection = True

            'Config custom name style settings end.
            'Set range's style to custom name style.
            worksheet.Range!A2.Value = "my test style"
            worksheet.Range!A2.Style = worksheet.Workbook.Styles("testStyle")
            worksheet.Range!A2.RowHeightInPixel = 60
            worksheet.Range!A2.ColumnWidthInPixel = 150
        End Sub
    End Class
End Namespace

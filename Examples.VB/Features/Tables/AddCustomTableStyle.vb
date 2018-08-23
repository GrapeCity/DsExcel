Imports System.Drawing

Namespace Features.Tables
    Public Class AddCustomTableStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Add one custom table style.
            Dim style As ITableStyle = workbook.TableStyles.Add("test")

            'Set WholeTable element style.
            style.TableStyleElements(TableStyleElementType.WholeTable).Font.Italic = True
            style.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.White
            style.TableStyleElements(TableStyleElementType.WholeTable).Font.Strikethrough = True
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders.LineStyle = BorderLineStyle.Dotted
            style.TableStyleElements(TableStyleElementType.WholeTable).Borders.Color = Color.FromArgb(0, 193, 213)
            style.TableStyleElements(TableStyleElementType.WholeTable).Interior.Color = Color.FromArgb(59, 92, 170)

            'Set FirstColumnStripe element style.
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).Font.Bold = True
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).Font.Color = Color.FromArgb(255, 0, 0)
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).Borders.LineStyle = BorderLineStyle.Thick
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).Borders.ThemeColor = ThemeColor.Accent5
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).Interior.Color = Color.FromArgb(255, 255, 0)
            style.TableStyleElements(TableStyleElementType.FirstColumnStripe).StripeSize = 2

            'Set SecondColumnStripe element style.
            style.TableStyleElements(TableStyleElementType.SecondColumnStripe).Font.Color = Color.FromArgb(255, 0, 255)
            style.TableStyleElements(TableStyleElementType.SecondColumnStripe).Borders.LineStyle = BorderLineStyle.DashDot
            style.TableStyleElements(TableStyleElementType.SecondColumnStripe).Borders.Color = Color.FromArgb(42, 105, 162)
            style.TableStyleElements(TableStyleElementType.SecondColumnStripe).Interior.Color = Color.FromArgb(204, 204, 255)

            'add table.
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("A1:F7"), True)
            worksheet.Range("A:F").ColumnWidth = 15

            'set custom table style to table.
            table.TableStyle = style
        End Sub
    End Class
End Namespace

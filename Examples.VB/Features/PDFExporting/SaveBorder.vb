Namespace Features.PDFExporting
    Public Class SaveBorder
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Single cell border
            sheet.Range!B2.Borders.ThemeColor = ThemeColor.Accent1
            sheet.Range!B2.Borders.LineStyle = BorderLineStyle.SlantDashDot
            sheet.Range!B2.Borders(BordersIndex.DiagonalUp).ThemeColor = ThemeColor.Accent1
            sheet.Range!B2.Borders(BordersIndex.DiagonalUp).LineStyle = BorderLineStyle.SlantDashDot
            sheet.Range!B2.Borders(BordersIndex.DiagonalDown).ThemeColor = ThemeColor.Accent1
            sheet.Range!B2.Borders(BordersIndex.DiagonalDown).LineStyle = BorderLineStyle.SlantDashDot

            'Range border
            sheet.Range("D2:E3").Borders.ThemeColor = ThemeColor.Accent1
            sheet.Range("D2:E3").Borders.LineStyle = BorderLineStyle.DashDot
            sheet.Range("D2:E3").Borders(BordersIndex.DiagonalDown).ThemeColor = ThemeColor.Accent1
            sheet.Range("D2:E3").Borders(BordersIndex.DiagonalDown).LineStyle = BorderLineStyle.DashDot

            'Merge cell border
            sheet.Range("B6:C7").Merge()
            sheet.Range("B6:C7").Borders.ThemeColor = ThemeColor.Accent1
            sheet.Range("B6:C7").Borders.LineStyle = BorderLineStyle.Double
            sheet.Range("B6:C7").Borders(BordersIndex.DiagonalUp).ThemeColor = ThemeColor.Accent1
            sheet.Range("B6:C7").Borders(BordersIndex.DiagonalUp).LineStyle = BorderLineStyle.Double

            'Border style on table
            Dim table As ITable = sheet.Tables.Add(sheet.Range("B12:G22"), True)

            'Create custom table style
            Dim customTableStyle As ITableStyle = workbook.TableStyles("TableStyleMedium10").Duplicate()

            'Set outline border for "whole table" style
            Dim wholeTableStyle = customTableStyle.TableStyleElements(TableStyleElementType.WholeTable)
            wholeTableStyle.Borders(BordersIndex.EdgeTop).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeRight).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeBottom).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeLeft).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thick

            'Set vertical border for "first row strip" style
            Dim firstRowStripStyle = customTableStyle.TableStyleElements(TableStyleElementType.FirstRowStripe)
            firstRowStripStyle.Borders(BordersIndex.InsideVertical).ThemeColor = ThemeColor.Accent6
            firstRowStripStyle.Borders(BordersIndex.InsideVertical).LineStyle = BorderLineStyle.Dashed

            'Apply custom style to table
            table.TableStyle = customTableStyle
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

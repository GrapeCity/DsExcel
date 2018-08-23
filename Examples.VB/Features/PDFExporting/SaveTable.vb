Namespace Features.PDFExporting
    Public Class SaveTable
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'Add Table
            Dim table As ITable = sheet.Tables.Add(sheet.Range("B5:G16"), True)
            table.ShowTotals = True

            'Set values
            Dim data() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11}
            sheet.Range("C6:C16").Value = data
            sheet.Range("D6:D16").Value = data

            'Set total functions
            table.Columns(1).TotalsCalculation = TotalsCalculation.Average
            table.Columns(2).TotalsCalculation = TotalsCalculation.Sum

            'Create custom table style
            Dim customTableStyle As ITableStyle = workbook.TableStyles("TableStyleMedium10").Duplicate()

            Dim wholeTableStyle = customTableStyle.TableStyleElements(TableStyleElementType.WholeTable)
            wholeTableStyle.Font.Italic = True
            wholeTableStyle.Borders(BordersIndex.EdgeTop).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeTop).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeRight).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeRight).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeBottom).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeBottom).LineStyle = BorderLineStyle.Thick
            wholeTableStyle.Borders(BordersIndex.EdgeLeft).ThemeColor = ThemeColor.Accent1
            wholeTableStyle.Borders(BordersIndex.EdgeLeft).LineStyle = BorderLineStyle.Thick

            Dim firstRowStripStyle = customTableStyle.TableStyleElements(TableStyleElementType.FirstRowStripe)
            firstRowStripStyle.Font.Bold = True

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

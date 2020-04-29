Namespace Features.PDFExporting.Text
    Public Class NumberFormating
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            sheet.Range("B3:B7").Value = 123456.789
            sheet.Range("B9:B13").Value = -123456.789

            sheet.Range("B4, B10").NumberFormat = "0.00;[Red]0.00"
            sheet.Range("B5, B11").NumberFormat = "$#,##0.00;[Red]$#,##0.00"
            sheet.Range("B6, B12").NumberFormat = "0.00E+00"
            sheet.Range("B7, B13").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* "" - ""??_);_(@_)"

            'Other settings
            sheet.Columns(1).ColumnWidthInPixel = 100
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
    End Class
End Namespace

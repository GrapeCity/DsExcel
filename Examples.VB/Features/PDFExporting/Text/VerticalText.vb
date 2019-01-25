Namespace Features.PDFExporting.Text
    Public Class VerticalText
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)
            sheet.Range!A1.Font.Name = "@Meiryo"
            sheet.Range!A1.Value = "日本語（にほんご、にっぽんご）は、主として、日本列島で使用されてきた言語である。GrapeCity"
            sheet.Range!A1.HorizontalAlignment = HorizontalAlignment.Right
            sheet.Range!A1.VerticalAlignment = VerticalAlignment.Top
            sheet.Range!A1.Orientation = -90
            sheet.Range!A1.WrapText = True

            sheet.Range!A1.ColumnWidth = 27
            sheet.Range!A1.RowHeight = 190
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property


    End Class
End Namespace



Namespace Features.Formatting.Alignment
    Public Class WrapText
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeB3 As IRange = worksheet.Range!B3
            rangeB3.Value = "The WrapText property is applied to wrap the text within a cell"
            rangeB3.WrapText = True
            worksheet.Rows(2).RowHeight = 150
        End Sub
    End Class
End Namespace

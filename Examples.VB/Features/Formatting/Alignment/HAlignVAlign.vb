Namespace Features.Formatting.Alignment
    Public Class HAlignVAlign
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Columns(0).ColumnWidth = 17

            Dim rangeA1 As IRange = worksheet.Range!A1
            rangeA1.Value = "Right and top"
            rangeA1.HorizontalAlignment = HorizontalAlignment.Right
            rangeA1.VerticalAlignment = VerticalAlignment.Top

            Dim rangeA2 As IRange = worksheet.Range!A2
            rangeA2.Value = "Center"
            rangeA2.HorizontalAlignment = HorizontalAlignment.Center
            rangeA2.VerticalAlignment = VerticalAlignment.Center

            Dim rangeA3 As IRange = worksheet.Range!A3
            rangeA3.Value = "Left and bottom, indent"
            rangeA3.IndentLevel = 1
        End Sub
    End Class
End Namespace

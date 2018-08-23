Imports System.Drawing

Namespace Features.RangeOperations
    Public Class GetWorksheetUsedRange
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("H6:M7").Value = 1
            worksheet.Range("J9:J10").Merge()

            'set interior color for worksheet usedRange "H6:M10".
            Dim usedrange = worksheet.UsedRange
            usedrange.Interior.Color = Color.LightBlue
        End Sub
    End Class
End Namespace

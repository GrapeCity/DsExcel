Imports System.Drawing

Namespace Features.RangeOperations
    Public Class CutCopyRange
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            worksheet.Range("B3:D12").Value = 5
            worksheet.Range("B3:D12").Interior.Color = Color.LightGreen

            'Copy
            worksheet.Range("B3:D12").Copy(worksheet.Range!E5)

            'Cut
            worksheet.Range("B3:D12").Cut(worksheet.Range("I5:K14"))
            worksheet.Range("I1:K2").Value = 2
            worksheet.Range("I1:K2").Interior.Color = Color.Pink

            'cross sheet cut copy.
            worksheet.Range("I1:K2").Cut(worksheet2.Range!H5)
            worksheet.Range("G4:H5").Copy(worksheet2.Range!A1)
        End Sub
    End Class
End Namespace

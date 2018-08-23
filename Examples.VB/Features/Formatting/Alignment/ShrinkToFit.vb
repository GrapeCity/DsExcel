Namespace Features.Formatting.Alignment
    Public Class ShrinkToFit
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim rangeB1 As IRange = worksheet.Range!B1
            rangeB1.Value = "The ShrinkToFit property is applied"
            rangeB1.ShrinkToFit = True
        End Sub
    End Class
End Namespace

Imports System.Drawing

Namespace Features.RangeOperations
    Public Class GetFeatureRelatedUsedRange
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("A1:B2").Value = New Object(,)  {
                {1, 2},
                {"aaa", "bbb"}
            }
            worksheet.Range("A2:C3").Interior.Color = Color.Green

            'style used range is A2:C3.
            Dim UsedRange_style = worksheet.GetUsedRange(UsedRangeType.Style)
            UsedRange_style.Interior.Color = Color.LightBlue
        End Sub
    End Class
End Namespace

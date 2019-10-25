Namespace Features.Grouping
    Public Class AddSubtotalAdvanced
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim targetRange As IRange = workbook.ActiveSheet.Range("A1:C9")
            ' Set data
            targetRange.Value = New Object(,) {
                {"Grade", "Class", "Score", "Student ID"},
                {1, 1, 93, 1},
                {1, 1, 87, 2},
                {1, 2, 97, 3},
                {1, 2, 95, 4},
                {2, 1, 83, 5},
                {2, 1, 87, 6},
                {2, 2, 96, 7},
                {2, 2, 83, 8}
            }

            ' Group by Grade select Average(Score)
            targetRange.Subtotal(groupBy:=1, ' Grade
                                 subtotalFunction:=ConsolidationFunction.Average,
                                 totalList:={3}, ' Score
                                 replace:=False, pageBreaks:=True)

            ' Group by Class select Average(Score)
            targetRange.Subtotal(groupBy:=2, ' Class
                                 subtotalFunction:=ConsolidationFunction.Average,
                                 totalList:={3}, ' Score
                                 replace:=False)

            targetRange.AutoFit()
        End Sub

    End Class
End Namespace

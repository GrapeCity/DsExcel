Namespace Features.Grouping
    Public Class RemoveSubtotal
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim targetRange As IRange = workbook.ActiveSheet.Range("A1:C9")
            ' Set data
            targetRange.Value = New Object(,) {
                {"Player", "Side", "Commander"},
                {1, "Soviet", "AI"},
                {2, "Soviet", "AI"},
                {3, "Soviet", "Human"},
                {4, "Allied", "Human"},
                {5, "Allied", "Human"},
                {6, "Allied", "AI"},
                {7, "Empire", "AI"},
                {8, "Empire", "AI"}
            }

            ' Subtotal
            targetRange.Subtotal(groupBy:=2, ' Side
                                 subtotalFunction:=ConsolidationFunction.Count,
                                 totalList:={2} ' Side
                                 )

            ' Remove subtotal
            workbook.ActiveSheet.Range("A1:C13").RemoveSubtotal()
        End Sub
    End Class
End Namespace

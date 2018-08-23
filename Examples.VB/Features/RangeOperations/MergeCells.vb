Namespace Features.RangeOperations
    Public Class MergeCells
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'A1:C4 is a single merged cell.
            worksheet.Range("A1:C4").Merge()

            'H5:J5 is a single merged cell.
            'H6:J6 is a single merged cell.
            worksheet.Range("H5:J6").Merge(True)

            'select A1:H5's entire merge area A1:J5, entire merge area is a bounding rectangle.
            Dim entireMergeArea = worksheet.Range("A1:H5").EntireMergeArea
            entireMergeArea.Select()

            'judge if H5 is a merged cell.
            If worksheet.Range!J5.MergeCells Then
                'set value to the top left cell of the merge area.
                worksheet.Range!J5.EntireMergeArea(0, 0).Value = 1
            End If
        End Sub
    End Class
End Namespace

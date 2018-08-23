Namespace Features.RangeOperations
    Public Class InsertCellsDeleteCells
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet1 As IWorksheet = workbook.Worksheets(0)

            Dim worksheet2 As IWorksheet = workbook.Worksheets.Add()
            Dim data = New Object(,) {
                {1, 2, 3},
                {4, 5, 6},
                {7, 8, 9}
            }
            worksheet1.Range("A1:C3").Value = data
            worksheet2.Range("A1:C3").Value = data

            'Insert cells
            worksheet1.Range!A2.Insert() 'Auto
            worksheet1.Range!B2.Insert(InsertShiftDirection.Down)
            worksheet1.Range!C2.Insert(InsertShiftDirection.Right)

            'Delete cells
            worksheet2.Range!A2.Delete() 'Auto
            worksheet2.Range!B2.Delete(DeleteShiftDirection.Left)
            worksheet2.Range!C2.Delete(DeleteShiftDirection.Up)
        End Sub
    End Class
End Namespace

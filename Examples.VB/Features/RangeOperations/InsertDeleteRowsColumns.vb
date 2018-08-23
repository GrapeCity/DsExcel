Namespace Features.RangeOperations
    Public Class InsertDeleteRowsColumns
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

            'Insert rows
            worksheet1.Range("A2:B2").EntireRow.Insert()
            worksheet1.Range("3:5").Insert()

            'Insert columns
            worksheet1.Range("B3:B5").EntireColumn.Insert()
            worksheet1.Range("A:A").Insert()

            'Delete rows
            worksheet2.Range("A3:A5").EntireRow.Delete()
            worksheet2.Range("2:4").Delete()

            'Delete columns
            worksheet2.Range("B3:B5").EntireColumn.Delete()
            worksheet2.Range("A:A").Delete()
        End Sub
    End Class
End Namespace

Namespace Features.Worksheets
    Public Class AddWorksheet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            ' Add a new worksheet to the workbook. The worksheet will be inserted into the end of the existing worksheet collection.
            workbook.Worksheets.Add()

            'Add a new worksheet to the specified position in the collection of worksheets.
            workbook.Worksheets.AddBefore(workbook.Worksheets(0))
            workbook.Worksheets.AddAfter(workbook.Worksheets(1))

            'Set worksheet's name.
            workbook.Worksheets(2).Name = "Product Plan"
        End Sub
    End Class
End Namespace

Namespace Features.Formatting.CellType
    Public Class AddButtonCellType
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim cellType As New ButtonCellType With {
                .Text = "Hello",
                .ButtonBackColor = "Azure",
                .MarginLeft = 10,
                .MarginRight = 10
            }

            worksheet.Range!C5.CellType = cellType
        End Sub
    End Class
End Namespace

Namespace Features.Formatting.CellType
    Public Class AddSheetCellType
        Inherits ExampleBase

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim buttonCellType As New ButtonCellType With {
                .Text = "Button",
                .ButtonBackColor = "Azure",
                .MarginLeft = 10,
                .MarginRight = 10
            }

            worksheet.CellType = buttonCellType
        End Sub
    End Class
End Namespace

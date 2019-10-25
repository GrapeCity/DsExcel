Namespace Features.Formatting.CellType

    Public Class AddRowColumnCellType
        Inherits ExampleBase

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Columns(3).ColumnWidthInPixel = 100

            Dim cellType As New CheckBoxCellType With {
                .Caption = "CheckBox",
                .TextTrue = "True",
                .TextFalse = "False",
                .IsThreeState = True,
                .TextAlign = CheckBoxAlign.Right
            }

            worksheet.Columns(3).CellType = cellType
            worksheet.Range("D1:D10").Value = True

            Dim buttonCellType As New ButtonCellType With {
                .Text = "Button",
                .ButtonBackColor = "Azure",
                .MarginLeft = 10,
                .MarginRight = 10
            }

            worksheet.Rows(3).CellType = buttonCellType
        End Sub
    End Class
End Namespace

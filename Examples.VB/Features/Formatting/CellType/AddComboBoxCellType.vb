Namespace Features.Formatting.CellType
    Public Class AddComboBoxCellType
        Inherits ExampleBase

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim cellType As New ComboBoxCellType With {
                .EditorValueType = EditorValueType.Value
            }

            cellType.Items.Add(New ComboBoxCellItem With {
                .Value = "US",
                .Text = "United States"
            })

            cellType.Items.Add(New ComboBoxCellItem With {
                .Value = "CN",
                .Text = "China"
            })

            cellType.Items.Add(New ComboBoxCellItem With {
                .Value = "JP",
                .Text = "Japan"
            })

            With worksheet.Range!C5
                .CellType = cellType
                .Value = "CN"
            End With
        End Sub
    End Class
End Namespace

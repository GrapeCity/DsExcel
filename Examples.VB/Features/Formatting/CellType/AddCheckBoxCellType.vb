Namespace Features.Formatting.CellType
    Public Class AddCheckBoxCellType
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim cellType As New CheckBoxCellType With {
                .Caption = "Caption",
                .TextTrue = "True",
                .TextFalse = "False",
                .TextIndeterminate = "Indeterminate",
                .IsThreeState = True,
                .TextAlign = CheckBoxAlign.Right
            }

            worksheet.Range("C5:C6").CellType = cellType
            worksheet.Range("C5").Value = True
            worksheet.Range("C6").Value = False
        End Sub
    End Class
End Namespace

Namespace Features.Formatting.CellType
    Public Class AddHyperlinkCellType
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim cellType As New HyperLinkCellType With {
                .Text = "Goole",
                .LinkColor = "Blue",
                .LinkToolTip = "Search by google",
                .VisitedLinkColor = "Green",
                .Target = HyperLinkTargetType.Blank
            }

            With worksheet.Range!C5
                .CellType = cellType
                .Value = "http://www.google.com"
            End With
        End Sub
    End Class
End Namespace

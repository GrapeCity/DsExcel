Imports System.Drawing

Namespace Features.RangeOperations
    Public Class AccessEachAreaOfRange
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim range = worksheet.Range("A5:B7, C3, H5:N6")

            'set interior color for area1 A5:B7.
            Dim area1 = worksheet.Range("A5:B7, C3, H5:N6").Areas(0)

            area1.Interior.Color = Color.Pink

            'set interior color for area2 C3.
            Dim area2 = worksheet.Range("A5:B7, C3, H5:N6").Areas(1)
            area2.Interior.Color = Color.LightGreen

            'set interior color for area3 H5:N6.
            Dim area3 = worksheet.Range("A5:B7, C3, H5:N6").Areas(2)
            area3.Interior.Color = Color.LightBlue
        End Sub
    End Class
End Namespace

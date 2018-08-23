Imports System.Drawing

Namespace Features.RangeOperations
    Public Class CellInfo
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' cell's value B2
            Dim cell As String = Excel.CellInfo.CellIndexToName(1, 1)
            worksheet.Range(cell).Interior.Color = Color.LightBlue
            Dim rowIndex, columnIndex As Integer

            ' rowIndex is 3 and columnIndex is 2
            Excel.CellInfo.CellNameToIndex("C4", rowIndex, columnIndex)
            worksheet.Range(rowIndex, columnIndex).Interior.Color = Color.LightCoral

            ' column is D
            Dim column As String = Excel.CellInfo.ColumnIndexToName(3)
            worksheet.Range(String.Format("{0}:{0}", column)).Interior.Color = Color.LightGreen

            ' columnIndex is 4
            columnIndex = Excel.CellInfo.ColumnNameToIndex("E")
            worksheet.Columns(columnIndex).Interior.Color = Color.LightSalmon

            ' row is 3
            Dim row As String = Excel.CellInfo.RowIndexToName(2)
            worksheet.Range(String.Format("{0}:{0}", row)).Interior.Color = Color.LightSteelBlue

            ' rowIndex is 4
            rowIndex = Excel.CellInfo.RowNameToIndex("5")
            worksheet.Rows(rowIndex).Interior.Color = Color.LightSkyBlue
        End Sub
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

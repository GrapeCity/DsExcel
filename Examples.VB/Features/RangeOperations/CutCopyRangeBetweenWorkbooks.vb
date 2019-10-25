Imports System.Drawing

Namespace Features.RangeOperations
    Public Class CutCopyRangeBetweenWorkbooks
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Load template file Home inventory.xlsx from resource
            Dim source_workbook As New Excel.Workbook
            Dim source_fileStream = GetResourceStream("Home inventory.xlsx")
            source_workbook.Open(source_fileStream)

            'Hide gridline
            workbook.ActiveSheet.SheetView.DisplayGridlines = False

            workbook.ActiveSheet.Range!A1.Value = "Copy content from the first sheet of source workbook"
            workbook.ActiveSheet.Range!A1.Font.Color = Color.Red
            workbook.ActiveSheet.Range!A1.Font.Bold = True

            'Copy content of active sheet from source workbook to the current sheet at A2
            source_workbook.ActiveSheet.GetUsedRange().Copy(workbook.ActiveSheet.Range!A2, PasteType.Default Or PasteType.RowHeights Or PasteType.ColumnWidths)

            workbook.ActiveSheet.Range!C21.Value = "Cut content from the second sheet of source workbook"
            workbook.ActiveSheet.Range!C21.Font.Color = Color.Red
            workbook.ActiveSheet.Range!C21.Font.Bold = True

            'Cut content of second sheet from source workbook to the current sheet at C22
            source_workbook.Worksheets(1).Range("2:15").Cut(workbook.ActiveSheet.Range!C22)

            'Make the theme of two workbooks same
            workbook.Theme = source_workbook.Theme
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Home inventory.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Home inventory.xlsx"}
            End Get
        End Property
    End Class
End Namespace

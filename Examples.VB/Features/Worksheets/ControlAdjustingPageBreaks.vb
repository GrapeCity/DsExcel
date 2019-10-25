Namespace Features.Worksheets
    Public Class ControlAdjustingPageBreaks
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            sheet.Range("A1:E5").Value = New Object(,) {
                {1, 2, 3, 4, 5},
                {6, 7, 8, 9, 10},
                {11, 12, 13, 14, 15},
                {16, 17, 18, 19, 20},
                {21, 22, 23, 24, 25}
            }

            'Add page break
            sheet.HPageBreaks.Add(sheet.Range!D4) 'add a horizontal page break before the fourth row.
            sheet.VPageBreaks.Add(sheet.Range!D4) 'add a vertical page break before the fourth column.

            'delete rows and columns before the page breaks, the page breaks will be adjusted.
            sheet.Range("1:1").Delete() ' the hPageBreak is before the third row.
            sheet.Range("A:A").Delete() ' the vPageBreak is before the third column.

            'set the page breaks are fixed, it will not be adjusted when inserting/deleting rows/columns.
            sheet.FixedPageBreaks = True

            sheet.Range("1:1").Delete() ' the hPageBreak is still before the third row.
            sheet.Range("A:A").Delete() ' the vPageBreak is still before the third column.
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

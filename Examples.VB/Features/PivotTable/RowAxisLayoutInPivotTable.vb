Namespace Features.PivotTable
    Public Class RowAxisLayoutInPivotTable
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sourceData As Object(,) = {
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Carrots", "Vegetables", 4270, #2018-1-6#, "United States"},
                {2, "Broccoli", "Vegetables", 8239, #2018-1-7#, "United Kingdom"},
                {3, "Banana", "Fruit", 617, #2018-1-8#, "United States"},
                {4, "Banana", "Fruit", 8384, #2018-1-10#, "Canada"},
                {5, "Beans", "Vegetables", 2626, #2018-1-10#, "Germany"},
                {6, "Orange", "Fruit", 3610, #2018-1-11#, "United States"},
                {7, "Broccoli", "Vegetables", 9062, #2018-1-11#, "Australia"},
                {8, "Banana", "Fruit", 6906, #2018-1-16#, "New Zealand"},
                {9, "Apple", "Fruit", 2417, #2018-1-16#, "France"},
                {10, "Apple", "Fruit", 7431, #2018-1-16#, "Canada"},
                {11, "Banana", "Fruit", 8250, #2018-1-16#, "Germany"},
                {12, "Broccoli", "Vegetables", 7012, #2018-1-18#, "United States"},
                {13, "Carrots", "Vegetables", 1903, #2018-1-20#, "Germany"},
                {14, "Broccoli", "Vegetables", 2824, #2018-1-22#, "Canada"},
                {15, "Apple", "Fruit", 6946, #2018-1-24#, "France"}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:F16").Value = sourceData
            worksheet.Range("A:F").ColumnWidth = 15

            Dim pivotcache = workbook.PivotCaches.Create(worksheet.Range("A1:F16"))
            Dim pivottable = worksheet.PivotTables.Add(pivotcache, worksheet.Range!H7, "PivotTable1")
            worksheet.Range("D2:D16").NumberFormat = "$#,##0.00"
            worksheet.Range("I9:O11").NumberFormat = "$#,##0.00"
            worksheet.Range("H:O").ColumnWidth = 12

            'config pivot fields
            pivottable.PivotFields!Category.Orientation = PivotFieldOrientation.RowField
            pivottable.PivotFields!Product.Orientation = PivotFieldOrientation.ColumnField
            pivottable.PivotFields!Amount.Orientation = PivotFieldOrientation.DataField
            pivottable.PivotFields!Country.Orientation = PivotFieldOrientation.RowField

            ' Set row axis layout to tabular
            pivottable.SetRowAxisLayout(LayoutRowType.TabularRow)
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property ShowScreenshot As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class

End Namespace

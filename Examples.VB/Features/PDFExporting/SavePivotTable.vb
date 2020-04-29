Namespace Features.PDFExporting
    Public Class SavePivotTable
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sourceData(,) As Object = {
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Broccoli", "Vegetables", 8239, #2018-1-7#, "United Kingdom"},
                {2, "Banana", "Fruit", 617, #2018-1-8#, "United States"},
                {3, "Banana", "Fruit", 8384, #2018-1-10#, "Canada"},
                {4, "Beans", "Vegetables", 2626, #2018-1-10#, "Germany"},
                {5, "Orange", "Fruit", 3610, #2018-1-11#, "United States"},
                {6, "Broccoli", "Vegetables", 9062, #2018-1-11#, "Australia"},
                {7, "Banana", "Fruit", 6906, #2018-1-16#, "New Zealand"},
                {8, "Apple", "Fruit", 2417, #2018-1-16#, "France"},
                {9, "Apple", "Fruit", 7431, #2018-1-16#, "Canada"},
                {10, "Banana", "Fruit", 8250, #2018-1-16#, "Germany"},
                {11, "Broccoli", "Vegetables", 7012, #2018-1-18#, "United States"},
                {12, "Broccoli", "Vegetables", 2824, #2018-1-22#, "Canada"},
                {13, "Apple", "Fruit", 6946, #2018-1-24#, "France"}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("K20:P33").Value = sourceData
            worksheet.Range("K:P").ColumnWidth = 15
            ' Add pivot table
            Dim pivotcache = workbook.PivotCaches.Create(worksheet.Range("K20:P33"))
            Dim pivottable = worksheet.PivotTables.Add(pivotcache, worksheet.Range!A1, "pivottable1")
            worksheet.Range("N21:N35").NumberFormat = "$#,##0.00"
            worksheet.Range("A:G").ColumnWidth = 12

            'config pivot table's fields
            Dim field_Date = pivottable.PivotFields("Date")
            field_Date.Orientation = PivotFieldOrientation.PageField

            Dim field_Category = pivottable.PivotFields("Category")
            field_Category.Orientation = PivotFieldOrientation.RowField

            Dim field_Product = pivottable.PivotFields("Product")
            field_Product.Orientation = PivotFieldOrientation.ColumnField

            Dim field_Amount = pivottable.PivotFields("Amount")
            field_Amount.Orientation = PivotFieldOrientation.DataField
            field_Amount.NumberFormat = "$#,##0.00"

            Dim field_Country = pivottable.PivotFields("Country")
            field_Country.Orientation = PivotFieldOrientation.RowField

            ' Set pivot style
            pivottable.TableStyle = "PivotStyleMedium28"
        End Sub

        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

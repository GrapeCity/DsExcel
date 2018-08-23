Namespace Features.Slicer
    Public Class AddSlicersForPivotTable
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sourceData = {
                {"Order ID", "Product", "Category", "Amount", "Date", "Country"},
                {1, "Carrots", "Vegetables", 4270, #1/6/2018#, "United States"},
                {2, "Broccoli", "Vegetables", 8239, #1/7/2018#, "United Kingdom"},
                {3, "Banana", "Fruit", 617, #1/8/2018#, "United States"},
                {4, "Banana", "Fruit", 8384, #1/10/2018#, "Canada"},
                {5, "Beans", "Vegetables", 2626, #1/10/2018#, "Germany"},
                {6, "Orange", "Fruit", 3610, #1/11/2018#, "United States"},
                {7, "Broccoli", "Vegetables", 9062, #1/11/2018#, "Australia"},
                {8, "Banana", "Fruit", 6906, #1/16/2018#, "New Zealand"},
                {9, "Apple", "Fruit", 2417, #1/16/2018#, "France"},
                {10, "Apple", "Fruit", 7431, #1/16/2018#, "Canada"},
                {11, "Banana", "Fruit", 8250, #1/16/2018#, "Germany"},
                {12, "Broccoli", "Vegetables", 7012, #1/18/2018#, "United States"},
                {13, "Carrots", "Vegetables", 1903, #1/20/2018#, "Germany"},
                {14, "Broccoli", "Vegetables", 2824, #1/22/2018#, "Canada"},
                {15, "Apple", "Fruit", 6946, #1/24/2018#, "France"}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:F16").Value = sourceData
            worksheet.Range("A:F").ColumnWidth = 15

            'Create pivot cache.
            Dim pivotcache As IPivotCache = workbook.PivotCaches.Create(worksheet.Range("A1:F16"))

            'Create pivot tables.
            Dim pivottable1 As IPivotTable = worksheet.PivotTables.Add(pivotcache, worksheet.Range!K5, "pivottable1")
            Dim pivottable2 As IPivotTable = worksheet.PivotTables.Add(pivotcache, worksheet.Range!N3, "pivottable2")
            worksheet.Range("D2:D16").NumberFormat = "$#,##0.00"

            'Config pivot fields
            Dim field_product1 As IPivotField = pivottable1.PivotFields(1)
            field_product1.Orientation = PivotFieldOrientation.RowField

            Dim field_Amount1 As IPivotField = pivottable1.PivotFields(3)
            field_Amount1.Orientation = PivotFieldOrientation.DataField

            Dim field_product2 As IPivotField = pivottable2.PivotFields(5)
            field_product2.Orientation = PivotFieldOrientation.RowField

            Dim field_Amount2 As IPivotField = pivottable2.PivotFields(2)
            field_Amount2.Orientation = PivotFieldOrientation.DataField
            field_Amount2.Function = ConsolidationFunction.Count

            'create slicer cache, the slicers base the slicer cache just control pivot table1.
            Dim cache As ISlicerCache = workbook.SlicerCaches.Add(pivottable1, "Product")
            Dim slicer1 As ISlicer = cache.Slicers.Add(workbook.Worksheets("Sheet1"), "p1", "Product", 30, 550, 100, 200)

            'add pivot table2 for slicer cache, the slicers base the slicer cache will control pivot tabl1 and pivot table2.
            cache.PivotTables.AddPivotTable(pivottable2)
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
    End Class
End Namespace

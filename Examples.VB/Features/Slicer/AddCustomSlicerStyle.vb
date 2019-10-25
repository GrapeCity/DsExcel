Imports System.Drawing

Namespace Features.Slicer
    Public Class AddCustomSlicerStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sourceData As Object(,) = {
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
            worksheet.Range("A:F").ColumnWidth = 15
            worksheet.Range("A1:F16").Value = sourceData

            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("A1:F16"), True)
            table.Columns(3).DataBodyRange.NumberFormat = "$#,##0.00"

            'create slicer cache for table.
            Dim cache As ISlicerCache = workbook.SlicerCaches.Add(table, "Category", "categoryCache")

            'add slicer
            Dim slicer As ISlicer = cache.Slicers.Add(workbook.Worksheets("Sheet1"), "cate2", "Category", 30, 550, 100, 200)

            'create custom slicer style.
            Dim slicerStyle As ITableStyle = workbook.TableStyles.Add("test")

            'set ShowAsAvailableSlicerStyle to True, the style will be treated as slicer style.
            slicerStyle.ShowAsAvailableSlicerStyle = True
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Name = "Arial"
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Bold = False
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Italic = False
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Font.Color = Color.White
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Borders.Color = Color.LightPink
            slicerStyle.TableStyleElements(TableStyleElementType.WholeTable).Interior.Color = Color.LightGreen

            'set slicer style to custom style.
            slicer.Style = slicerStyle
        End Sub
    End Class
End Namespace

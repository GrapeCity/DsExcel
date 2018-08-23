Imports System.Drawing

Namespace Features.Sorting
    Public Class SortRangeByFontColor
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim data = {
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", #6/8/1968#, "Blue", 67, 165},
                {"Nia", "New York", #7/3/1972#, "Brown", 62, 134},
                {"Jared", "New York", #3/2/1964#, "Hazel", 72, 180},
                {"Natalie", "Washington", #8/8/1972#, "Blue", 66, 163},
                {"Damon", "Washington", #2/2/1986#, "Hazel", 76, 176},
                {"Angela", "Washington", #2/15/1993#, "Brown", 68, 145}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("A1:F7").Value = data
            worksheet.Range("A:F").ColumnWidth = 15
            worksheet.Range!F2.Font.Color = Color.Red
            worksheet.Range!F3.Font.Color = Color.Gainsboro
            worksheet.Range!F4.Font.Color = Color.Green
            worksheet.Range!F5.Font.Color = Color.Gainsboro
            worksheet.Range!F6.Font.Color = Color.Gainsboro
            worksheet.Range!F7.Font.Color = Color.Gainsboro

            'The font colors which are same to "F7" cell's color will in the top.
            worksheet.Sort.SortFields.Add(New FontColorSortField(worksheet.Range("F2:F7"), worksheet.Range!F7.Font.Color, SortOrder.Ascending))
            worksheet.Sort.Range = worksheet.Range("A2:F7")
            worksheet.Sort.Orientation = SortOrientation.Columns
            worksheet.Sort.Apply()
        End Sub
    End Class
End Namespace

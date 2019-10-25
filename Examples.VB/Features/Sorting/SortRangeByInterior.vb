Imports System.Drawing

Namespace Features.Sorting
    Public Class SortRangeByInterior
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim data As Object(,) = {
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
            worksheet.Range!F2.Interior.Color = Color.LightPink
            worksheet.Range!F3.Interior.Color = Color.LightGreen
            worksheet.Range!F4.Interior.Color = Color.LightPink
            worksheet.Range!F5.Interior.Color = Color.LightGreen
            worksheet.Range!F6.Interior.Color = Color.LightBlue
            worksheet.Range!F7.Interior.Color = Color.LightPink

            '"F4" will in the top.
            worksheet.Sort.SortFields.Add(New CellColorSortField(worksheet.Range("F2:F7"), worksheet.Range!F4.DisplayFormat.Interior, SortOrder.Ascending))
            worksheet.Sort.Range = worksheet.Range("A2:F7")
            worksheet.Sort.Orientation = SortOrientation.Columns
            worksheet.Sort.Apply()
        End Sub
    End Class
End Namespace

Namespace Features.Filtering.DateFiltering
    Public Class DateFilter
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim data = {
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", #6/8/1968#, "Blue", 67, 165},
                {"Nia", "New York", #7/3/1972#, "Brown", 62, 134},
                {"Jared", "New York", #3/2/1964#, "Hazel", 72, 180},
                {"Natalie", "Washington", #8/8/1972#, "Blue", 66, 163},
                {"Damon", "Washington", #2/2/1986#, "Hazel", 76, 176},
                {"Angela", "Washington", #2/15/1993#, "Brown", 68, 145}
            }

            worksheet.Range("A1:F7").Value = data
            worksheet.Range("A:F").ColumnWidth = 15

            Dim criteria1 = (#7/3/1972#).ToString()
            Dim criteria2 = (#2/15/1993#).ToString()

            'filter date between 1972.7.3 and 1993.2.15
            worksheet.Range("A1:F7").AutoFilter(2, ">=" & criteria1, AutoFilterOperator.And, "<=" & criteria2)
        End Sub
    End Class
End Namespace

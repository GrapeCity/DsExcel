Namespace Features.Filtering.FilteringByIcon
    Public Class FilterByNoIcon
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

            Dim iconset As IIconSetCondition = worksheet.Range("E2:E7").FormatConditions.AddIconSetCondition()
            iconset.IconSet = workbook.IconSets(IconSetType.Icon3TrafficLights1)

            Dim condition As IFormatCondition = TryCast(worksheet.Range("E2:E7").FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, "71", "79"), IFormatCondition)
            condition.StopIfTrue = True

            'filter in the rows which has no icon.
            worksheet.Range("A1:F7").AutoFilter(4, Nothing, AutoFilterOperator.NoIcon)
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

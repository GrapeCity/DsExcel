Namespace Features.ConditionalFormatting
    Public Class CreateMultiRuleForRange
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

            'iconset rule
            Dim iconset As IIconSetCondition = worksheet.Range("E2:E7").FormatConditions.AddIconSetCondition()
            iconset.IconSet = workbook.IconSets(IconSetType.Icon3TrafficLights1)

            'cell value rule added later, it has the highest priority, set StopIfTrue to True, if cell match condition, it will not apply icon set rule.
            Dim cellvalueRule As IFormatCondition = TryCast(worksheet.Range("E2:E7").FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, "66", "70"), IFormatCondition)
            cellvalueRule.Font.ThemeColor = ThemeColor.Accent1
            cellvalueRule.StopIfTrue = True
        End Sub
    End Class
End Namespace

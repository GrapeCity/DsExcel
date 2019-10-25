Imports System.Drawing

Namespace Features.ConditionalFormatting
    Public Class CreateDatabBarRule
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim data As Object(,) = {
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", #6/8/1968#, "Blue", 67, 165},
                {"Nia", "New York", #7/3/1972#, "Brown", 62, 134},
                {"Jared", "New York", #3/2/1964#, "Hazel", 72, 180},
                {"Natalie", "Washington", #8/8/1972#, "Blue", 66, 163},
                {"Damon", "Washington", #2/2/1986#, "Hazel", 76, 176},
                {"Angela", "Washington", #2/15/1993#, "Brown", 68, 145}
            }
            worksheet.Range("A1:F7").Value = data

            'data bar rule.
            Dim dataBar As IDataBar = worksheet.Range("E2:E7").FormatConditions.AddDatabar()
            dataBar.MinPoint.Type = ConditionValueTypes.LowestValue
            dataBar.MinPoint.Value = Nothing
            dataBar.MaxPoint.Type = ConditionValueTypes.HighestValue
            dataBar.MaxPoint.Value = Nothing
            dataBar.BarFillType = DataBarFillType.Gradient
            dataBar.BarColor.Color = Color.Green
            dataBar.Direction = DataBarDirection.Context
            dataBar.AxisColor.Color = Color.Red
            dataBar.AxisPosition = DataBarAxisPosition.Automatic
            dataBar.NegativeBarFormat.BorderColorType = DataBarNegativeColorType.Color
            dataBar.NegativeBarFormat.BorderColor.Color = Color.Blue
            dataBar.NegativeBarFormat.ColorType = DataBarNegativeColorType.Color
            dataBar.NegativeBarFormat.Color.Color = Color.Pink
            dataBar.ShowValue = False
        End Sub
    End Class
End Namespace

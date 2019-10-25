Imports System.Drawing

Namespace Features.ConditionalFormatting
    Public Class CreateThreeColorScaleRule
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

            'three color scale.
            Dim threeColorScaleRule As IColorScale = worksheet.Range("E2:E7").FormatConditions.AddColorScale(ColorScaleType.ThreeColorScale)
            threeColorScaleRule.ColorScaleCriteria(0).Type = ConditionValueTypes.LowestValue
            threeColorScaleRule.ColorScaleCriteria(0).FormatColor.Color = Color.Red
            threeColorScaleRule.ColorScaleCriteria(1).Type = ConditionValueTypes.Percent
            threeColorScaleRule.ColorScaleCriteria(1).Value = 30
            threeColorScaleRule.ColorScaleCriteria(1).FormatColor.Color = Color.Yellow
            threeColorScaleRule.ColorScaleCriteria(2).Type = ConditionValueTypes.HighestValue
            threeColorScaleRule.ColorScaleCriteria(2).FormatColor.Color = Color.Green
        End Sub
    End Class
End Namespace

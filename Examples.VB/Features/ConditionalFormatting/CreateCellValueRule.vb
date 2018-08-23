﻿Imports System.Drawing

Namespace Features.ConditionalFormatting
    Public Class CreateCellValueRule
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            worksheet.Range("B:C").ColumnWidth = 17

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

            'weight between 71 and 79, set its interior color to yellow.
            Dim condition As IFormatCondition = TryCast(worksheet.Range("E2:E7").FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, 66, 70), IFormatCondition)
            condition.Interior.Color = Color.LightGreen
        End Sub
    End Class
End Namespace

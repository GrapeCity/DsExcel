Imports System.Drawing

Namespace Features.ConditionalFormatting
    Public Class CreateUniqueRule
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim data As Object(,) = {
                {"Name", "City", "Birthday", "Eye color", "Weight", "Height"},
                {"Richard", "New York", #6/8/1968#, "Blue", 80, 165},
                {"Nia", "New York", #7/3/1972#, "Brown", 72, 134},
                {"Jared", "New York", #3/2/1964#, "Hazel", 71, 180},
                {"Natalie", "Washington", #8/8/1972#, "Blue", 80, 163},
                {"Damon", "Washington", #2/2/1986#, "Hazel", 58, 176},
                {"Angela", "Washington", #2/15/1993#, "Brown", 71, 145}
            }
            worksheet.Range("A1:F7").Value = data

            'Unique rule.
            Dim condition As IUniqueValues = worksheet.Range("E2:E7").FormatConditions.AddUniqueValues()
            condition.DupeUnique = DupeUnique.Unique
            condition.Font.Name = "Arial"
            condition.Interior.Color = Color.Pink
        End Sub
    End Class
End Namespace

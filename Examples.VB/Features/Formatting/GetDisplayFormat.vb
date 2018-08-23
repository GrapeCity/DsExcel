Imports System.Drawing

Namespace Features.Formatting
    Public Class GetDisplayFormat
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Range A1's displaying color is table style
            worksheet.Tables.Add(worksheet.Range("A1:E5"), True)
            Dim color_A1 = worksheet.Range!A1.DisplayFormat.Interior.Color

            'Range A1's displaying color will be cell style, yellow.
            worksheet.Range!A1.Interior.Color = Color.Yellow
            Dim color_A1_1 = worksheet.Range!A1.DisplayFormat.Interior.Color

            'Range A1's displaying color will be conditional format style, green.
            Dim condition As IFormatCondition = TryCast(worksheet.Range!A1.FormatConditions.Add(FormatConditionType.NoBlanksCondition), IFormatCondition)
            condition.Interior.Color = Color.Green
            Dim color_A1_2 = worksheet.Range!A1.DisplayFormat.Interior.Color
        End Sub
    End Class
End Namespace

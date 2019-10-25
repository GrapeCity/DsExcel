Namespace Features.Formatting
    Public Class CreateStyleBasedOn
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range!A1.Style = workbook.Styles("Good")
            worksheet.Range!A1.Value = "Good"

            ' Create and modify a style based on current existing style
            Dim myGood As IStyle = workbook.Styles.Add("MyGood", workbook.Styles("Good"))
            myGood.Font.Bold = True
            myGood.Font.Italic = True

            worksheet.Range!B1.Style = workbook.Styles("MyGood")
            worksheet.Range!B1.Value = "MyGood"
        End Sub
    End Class
End Namespace

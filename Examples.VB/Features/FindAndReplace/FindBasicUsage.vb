Namespace Features.FindAndReplace
    Public Class FindBasicUsage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            Const CorrectWord = "Macro"
            worksheet.Range("A1:D5").Value = CorrectWord

            Const MisspelledWord = "marco"
            worksheet.Range("A2,C3,D1").Value = MisspelledWord

            ' Find the first misspelled word
            Dim searchRange As IRange = worksheet.Range("A1:D5")
            Dim firstMisspelled = searchRange.Find(MisspelledWord)

            ' Mark it with red foreground
            firstMisspelled.Font.Color = System.Drawing.Color.Red
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

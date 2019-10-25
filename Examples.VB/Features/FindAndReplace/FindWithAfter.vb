Namespace Features.FindAndReplace
    Public Class FindWithAfter
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            Const CorrectWord = "Macro"
            worksheet.Range("A1:D5").Value = CorrectWord

            Const MisspelledWord = "marco"
            worksheet.Range("A2,C3,D1").Value = MisspelledWord

            ' Find all misspelled words and mark them with red background
            Dim searchRange As IRange = worksheet.Range("A1:D5")
            Dim misspelledCell As IRange = Nothing
            Do
                misspelledCell = searchRange.Find(MisspelledWord, misspelledCell)
                If misspelledCell Is Nothing Then
                    Exit Do
                End If
                misspelledCell.Interior.Color = System.Drawing.Color.Red
            Loop
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

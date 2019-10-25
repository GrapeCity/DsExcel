Namespace Features.FindAndReplace
    Public Class FindWithLookIn
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data

            ' Add day to date		
            ' Day    Date	    Result
            ' 1      2019-05-01	2019-05-02
            worksheet.Range("A2:C2").Value = {"Day", "Date", "Result"}
            With worksheet.Range
                !A1.Value = "Add day to date"
                !A3.Value = 1
                !A3.AddComment("Enter the day offset")
                !B3.Formula = "=DATE(2019,5,1)"
                !B3.NumberFormat = "yyyy-mm-dd;@"
                !C3.Formula = "=B3+1"
                !C3.NumberFormat = "yyyy-mm-dd;@"
            End With
            worksheet.UsedRange.AutoFit()

            ' Find the first occurrence of "2019" in the formula bar 
            ' and mark it with green foreground color
            Dim searchRange = worksheet.Range("A1:C3")
            Dim first2019InFormulaBar = searchRange.Find("2019",,
                New FindOptions With {
                   .LookIn = FindLookIn.Formulas
                })
            first2019InFormulaBar.Font.Color = System.Drawing.Color.Green

            ' Find the first occurrence of 1 in text
            ' and mark it with blue foreground 
            Dim firstValue1 = searchRange.Find(1,,
                New FindOptions With {
                    .LookIn = FindLookIn.Texts
                })
            firstValue1.Font.Color = System.Drawing.Color.Blue

            ' Find the first occurrence of "day" in comments
            ' and mark it with yellow background 
            Dim firstDayComments = searchRange.Find("day",,
                New FindOptions With {
                    .LookIn = FindLookIn.Comments
                })
            firstDayComments.Interior.Color = System.Drawing.Color.Yellow

            ' Find the last occurrence of "2019" in the formula property
            ' and mark it with purple foreground
            Dim last2019OnlyFormula = searchRange.Find("2019",
                options:=New FindOptions With {
                    .LookIn = FindLookIn.OnlyFormulas,
                    .SearchDirection = SearchDirection.Previous
                })
            last2019OnlyFormula.Font.Color = System.Drawing.Color.Purple

        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

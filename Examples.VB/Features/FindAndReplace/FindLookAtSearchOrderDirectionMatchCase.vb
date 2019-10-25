Namespace Features.FindAndReplace
    Public Class FindLookAtSearchOrderDirectionMatchCase
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
                !B3.Formula = "=DATE(2019,5,1)"
                !B3.NumberFormat = "yyyy-mm-dd;@"
                !C3.Formula = "=B3+1"
                !C3.NumberFormat = "yyyy-mm-dd;@"
            End With
            worksheet.UsedRange.AutoFit()

            Dim searchRange = worksheet.Range("A1:C3")

            ' Find the last occurrence of 1 in text (match whole word, backward and by columns)
            ' and mark it with blue foreground and bigger font 
            Dim lastValue1 = searchRange.Find(1,,
                New FindOptions With {
                    .LookIn = FindLookIn.Texts,
                    .SearchDirection = SearchDirection.Previous,
                    .LookAt = LookAt.Whole, .SearchOrder = SearchOrder.ByColumns
                })
            lastValue1.Font.Color = System.Drawing.Color.Blue
            lastValue1.Font.Size += 8
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

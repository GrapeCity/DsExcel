Namespace Features.FindAndReplace
    Public Class ReplaceWithOptions
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data

            ' Skew matrix generator		
            ' Input:
            ' DegX    135     
            ' DegY    45      
            ' 
            ' Output:
            ' M11 1	    M12	1
            ' M21 -1	M22	1
            ' M31 0	    M32	0
            worksheet.Range("A2:A4").Value = {"Input:", "DegX", "DegY"}
            worksheet.Range("A7:A9").Value = {"M11", "M21", "M31"}
            worksheet.Range("C7:C9").Value = {"M12", "M22", "M32"}
            With worksheet.Range
                !B1.Value = "Skew matrix generator"
                !B3.Value = 135
                !B4.Value = 45
                !A6.Value = "Output:"
                !B7.Value = 1
                !B8.Formula = "=TAN(B3/180*3.14)"
                !B9.Value = 0
                !D7.Formula = "=TAN(B4/180*3.14)"
                !D8.Value = 1
                !D9.Value = 0
            End With

            ' Replace 3.14 with PI()
            Dim searchRange = worksheet.UsedRange
            searchRange.Replace(3.14, "PI()")

            ' Replace M with m (Match case)
            searchRange.Replace("M", "m", New ReplaceOptions With {.MatchCase = True})

            ' Replace m11 with M11 (Match whole word, match byte)
            searchRange.Replace("m11", "M11",
                New ReplaceOptions With {
                    .LookAt = LookAt.Whole, .MatchByte = True
                })
        End Sub
    End Class
End Namespace

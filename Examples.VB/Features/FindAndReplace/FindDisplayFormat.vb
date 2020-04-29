Namespace Features.FindAndReplace
    Public Class FindDisplayFormat
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            worksheet.Range("A1:C3").Value = "Text"

            With worksheet.Range!B2
                .Interior.Color = System.Drawing.Color.Red
                .Font.Color = System.Drawing.Color.White
                .Value = "B2"
            End With

            With worksheet.Range!A2
                .Interior.Color = System.Drawing.Color.Orange
                .Font.Color = System.Drawing.Color.White
                .Value = "A2"
            End With

            ' Find cells with red background and white foreground,
            ' and highlight them with bold and bigger text

            ' Create a temporary sheet to build a IDisplayFormat
            Dim displayFormatFactoryWorksheet = workbook.Worksheets.Add()
            Dim searchFormat As IDisplayFormat
            With displayFormatFactoryWorksheet.Range!A1
                .Interior.Color = System.Drawing.Color.Red
                .Font.Color = System.Drawing.Color.White
                searchFormat = .DisplayFormat
            End With

            ' Find and bold all occurrences
            Dim searchRange As IRange = worksheet.UsedRange
            Dim options As New FindOptions With {.SearchFormat = searchFormat}
            Dim foundCell = searchRange.Find("*",  , options)

            ' Highlight the found cell
            foundCell.Font.Bold = True
            foundCell.Font.Size += 8

            ' Dispose the temporary sheet
            displayFormatFactoryWorksheet.Delete()
        End Sub
    End Class
End Namespace

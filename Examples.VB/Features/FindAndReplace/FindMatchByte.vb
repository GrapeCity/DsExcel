Namespace Features.FindAndReplace
    Public Class FindMatchByte
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            ' This option is valid when culture is ja-JP or zh-CN.
            workbook.Culture = Globalization.CultureInfo.GetCultureInfo("ja-JP")
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            worksheet.Range("A1:A4").Value = {
                "Mario Games", "スーパーマリオブラザーズ",
                "ﾏﾘｵ&ﾙｲｰｼﾞRPG3 DX", "マリオ＆ルイージRPG1 DX"
            }

            ' Find the first cell that contains "マリオ" (match width) 
            ' and mark it with red foreground.
            Dim searchRange As IRange = worksheet.UsedRange
            Dim matchByteOptions As New FindOptions With {.MatchByte = True}
            Dim marioCell = searchRange.Find("マリオ",, matchByteOptions)
            marioCell.Font.Color = System.Drawing.Color.Red

            ' Find the first cell that contains "ルイージ" (ignore width) 
            ' and mark it with green background.
            Dim luigiCell = searchRange.Find("ルイージ")
            luigiCell.Interior.Color = System.Drawing.Color.Green
        End Sub
    End Class
End Namespace

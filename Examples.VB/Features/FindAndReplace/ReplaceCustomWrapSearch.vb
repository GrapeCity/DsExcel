Namespace Features.FindAndReplace
    Public Class ReplaceCustomWrapSearch
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            worksheet.Range("A1:A8").Value = {
                "Whats new in GcExcel v2 sp2", "Render Excel ranges inside PDF in .NET Core",
                "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                "How to format Pivot table styles in .NET Core (Support Team)",
                "Controlling page breaks when editing Excel files in .NET Core (Support Team)",
                "Combine different workbooks into PDF in .NET Core (Support Team)",
                "Repeating Excel rows/columns on exporting to PDF in .NET Core (Support Team)", "Using GcExcel with Kotlin"
            }

            ' Find ".NET Core" and replace them with ".NET 5", starting after A4
            Dim what = ".NET Core"
            Dim replacement = ".NET 5"
            Dim settings As New FindOptions
            Dim target = worksheet.UsedRange
            Dim after = worksheet.Range!A4

            ' Search start after A4
            Dim cellToReplace As IRange = after
            Do
                cellToReplace = target.Find(what, cellToReplace, settings)
                If cellToReplace Is Nothing Then
                    Exit Do
                End If

                ' Replace
                cellToReplace.Value = cellToReplace.Text.Replace(what, replacement)
            Loop

            ' Search reached the bottom of the range.
            ' Wrap search start at the top-left corner.
            If after IsNot Nothing Then
                Do
                    cellToReplace = target.Find(what, cellToReplace, settings)
                    If cellToReplace Is Nothing Then
                        Exit Do
                    End If

                    ' Replace
                    cellToReplace.Value = cellToReplace.Text.Replace(what, replacement)

                    If cellToReplace.Row = after.Row AndAlso cellToReplace.Column = after.Column Then
                        Exit Do
                    End If
                Loop
            End If
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

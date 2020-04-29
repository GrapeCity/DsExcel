Namespace Features.FindAndReplace
    Public Class ReplaceBasicUsage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Prepare data
            worksheet.Range("A1:A3").Value = {
                "Render Excel ranges inside PDF in .NET Core",
                "Control pagination when printing Excel document to PDF in .NET Core (Support Team)",
                "How to format Pivot table styles in .NET Core (Support Team)"
            }

            ' Replace ".NET Core" with ".NET 5"
            worksheet.UsedRange.Replace(".NET Core", ".NET 5")
        End Sub
    End Class
End Namespace

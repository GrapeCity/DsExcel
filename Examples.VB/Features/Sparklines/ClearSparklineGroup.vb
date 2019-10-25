Namespace Features.Sparklines
    Public Class ClearSparklineGroup
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim data As Object(,) = {
                {"Number", "Date", "Customer", "Description", "Trend", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days", "Amount"},
                {"1001", #5/21/2017#, "Customer A", "Invoice 1001", Nothing, 1200.15, 1916.18, 1105.23, 1806.53, Nothing},
                {"1002", #3/18/2017#, "Customer B", "Invoice 1002", Nothing, 896.23, 1005.53, 1800.56, 1150.49, Nothing},
                {"1003", #6/15/2017#, "Customer C", "Invoice 1003", Nothing, 827.63, 1009.23, 1869.23, 1002.56, Nothing}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("B2:K5").Value = data
            worksheet.Range("B:K").ColumnWidth = 15
            worksheet.Tables.Add(worksheet.Range("B2:K5"), True)
            worksheet.Tables(0).Columns(9).DataBodyRange.Formula = "=SUM(Table1[@[0-30 Days]:[>90 Days]])"

            'create a new group of sparklines.
            worksheet.Range!F3.SparklineGroups.Add(SparkType.Line, "G3:J3")

            'create another new group of sparklines.
            worksheet.Range("F4:F5").SparklineGroups.Add(SparkType.Column, "G4:J5")

            'clear sparkline group, all sparklines in the group will be cleaned.
            worksheet.Range("F3, F4").SparklineGroups.ClearGroups()
        End Sub
    End Class
End Namespace

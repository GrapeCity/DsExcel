Namespace Features.Tables
    Public Class ApplyTableStyle
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'add table.
            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("A1:F7"), True)
            worksheet.Range("A:F").ColumnWidth = 15

            'Add one custom table style.
            Dim style As ITableStyle = workbook.TableStyles.Add("test")

            'set custom table style for table.
            table.TableStyle = style

            'Use table style name get one build in table style.
            Dim tableStyle As ITableStyle = workbook.TableStyles("TableStyleMedium3")

            'set built-in table style for table.
            table.TableStyle = tableStyle
        End Sub
    End Class
End Namespace

Namespace Features.PDFExporting
    Public Class SaveSparklinesToPDF
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim data As Object(,) = {
                {"Customer", "0-30 Days", "30-60 Days", "60-90 Days", ">90 Days"},
                {"Customer A", 1200.15, 1916.18, 1105.23, 1806.53},
                {"Customer B", 896.23, 1005.53, 1800.56, 1150.49},
                {"Customer C", 827.63, 1009.23, 1869.23, 1002.56}
            }

            Dim worksheet As IWorksheet = workbook.Worksheets(0)
            worksheet.Range("B2:E5").Value = data
            worksheet.Range("B:F").ColumnWidth = 15
            worksheet.Range("B:E").HorizontalAlignment = HorizontalAlignment.Center

            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("B2:F5"), True)
            table.TableStyle = workbook.TableStyles("TableStyleMedium3")
            table.Columns(4).Name = "Sparklines"

            'create a new group of sparklines.
            worksheet.Range!F3.SparklineGroups.Add(SparkType.Line, "C3:E3")
            worksheet.Range!F4.SparklineGroups.Add(SparkType.Column, "C4:E4")
            worksheet.Range!F5.SparklineGroups.Add(SparkType.ColumnStacked100, "C5:E5")
        End Sub
        Public Overrides ReadOnly Property SavePdf As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace

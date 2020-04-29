Imports System.IO

Namespace Features.ImageExporting
    Public Class ExportChartToImage
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("Template_SalesTracker_report.xlsx")

            ' Open a xlsx file contains a chart
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Export the chart to image
            worksheet.Shapes("ProductIncomeChart").ToImage(outputStream, ImageType.PNG)
        End Sub
        Public Overrides ReadOnly Property SaveAsImages As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Template_SalesTracker_report.xlsx"}
            End Get
        End Property
    End Class
End Namespace

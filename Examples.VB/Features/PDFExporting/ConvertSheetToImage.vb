Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting
    Public Class ConvertSheetToImage
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("Employee absence schedule.xlsx")
            workbook.Open(fileStream)

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            'Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            Dim page As Pdf.Page = doc.NewPage()

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the first print area of the worksheet.
            Dim printArea As IRange = printManager.GetPrintAreas(worksheet)(0)
            'Get the size of the printArea.
            Dim size As SizeF = printManager.GetSize(printArea)

            'Modify the size of the page.
            page.Size = size

            ' Draw the printArea to the specified location on the page. 
            printManager.Draw(page, New PointF(0, 0), printArea)

            ' Saves the page as an image to a stream.
            page.SaveAsPng(outputStream, New Pdf.SaveAsImageOptions() With {.Resolution = 72})
        End Sub

        Public Overrides ReadOnly Property SaveAsImages As Boolean
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Employee absence schedule.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Employee absence schedule.xlsx"}
            End Get
        End Property
    End Class
End Namespace

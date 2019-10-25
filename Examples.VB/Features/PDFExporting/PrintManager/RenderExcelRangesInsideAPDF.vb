Imports System.Drawing
Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class RenderExcelRangesInsideAPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("FinancialReport.xlsx")
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'NOTE: To use this feature, you should have valid license for GrapeCity Documents for PDF.
            'Create a pdf document.
            Dim doc As New Pdf.GcPdfDocument
            doc.Load(GetResourceStream("Acme-Financial Report 2018.pdf"))

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Draw the contents of the sheet3 to the fourth page. 
            Dim printArea1 As IRange = workbook.Worksheets(2).Range("A3:C24")
            Dim size1 As SizeF = printManager.GetSize(printArea1)
            Dim position1 As RectangleF = doc.FindText(
                New Pdf.FindTextParams("Proposition enhancements are", True, True),
                New Common.OutputRange(4, 4))(0).Bounds.ToRect()
            printManager.Draw(
                doc.Pages(3),
                New RectangleF(position1.X + position1.Width + 70, position1.Y, size1.Width, size1.Height),
                printArea1)

            'Draw the contents of the sheet1 to the fifth page. 
            Dim printArea2 As IRange = workbook.Worksheets(0).Range("A4:E29")
            Dim size2 As SizeF = printManager.GetSize(printArea2)
            Dim position2 As RectangleF = doc.FindText(
                New Pdf.FindTextParams("expenditure, an improvement in working", True, True),
                New Common.OutputRange(5, 5))(0).Bounds.ToRect()
            printManager.Draw(
                doc.Pages(4),
                New RectangleF(position2.X,
                               position2.Y + position2.Height + 20,
                               size2.Width,
                               size2.Height),
                printArea2)

            'Draw the contents of the sheet2 to the sixth page. 
            Dim printArea3 As IRange = workbook.Worksheets(1).Range("A2:E28")
            Dim size3 As SizeF = printManager.GetSize(printArea3)
            Dim position3 As RectangleF = doc.FindText(
                New Pdf.FindTextParams("company will be able to continue", True, True),
                New Common.OutputRange(6, 6))(0).Bounds.ToRect()
            printManager.Draw(
                doc.Pages(5),
                New RectangleF(position3.X,
                               position3.Y + position3.Height + 20,
                               doc.Pages(5).Size.Width - (position3.X * 2) - 10,
                               size3.Height),
                printArea3)

            'Save the modified pages into pdf file.
            doc.Save(outputStream)
        End Sub

        Public Overrides ReadOnly Property SavePageInfos As Boolean

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
                Return "FinancialReport.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\FinancialReport.xlsx", "Acme-Financial Report 2018.pdf"}
            End Get
        End Property
    End Class
End Namespace

Imports System.Drawing

Namespace Features.PDFExporting.PdfPageSetup
    Public Class ConfigurePrintArea
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            Dim row As Integer = 50
            Dim column As Integer = 14
            Dim data(row - 1, column - 1) As Object
            For i As Integer = 0 To row - 1
                For j As Integer = 0 To column - 1
                    data(i, j) = "R" & i & "C" & j
                Next j
            Next i

            'Set data.
            sheet.Range("A1:N50").Value = data
            sheet.Range("C10:H20").Value = "PrintArea"
            sheet.Range("C10:H20").Interior.Color = Color.Yellow
            sheet.PageSetup.PrintHeadings = True

            'Set print area.
            sheet.PageSetup.PrintArea = "$C$10:$H$20"
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

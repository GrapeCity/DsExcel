Imports System.Drawing

Namespace Features.PDFExporting
    Public Class SaveWorkbookToPDF
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim sheet As IWorksheet = workbook.Worksheets(0)

            'set style.
            sheet.Range!A1.Value = "Sheet1"
            sheet.Range!A1.Font.Name = "Wide Latin"
            sheet.Range!A1.Font.Color = Color.Red
            sheet.Range!A1.Interior.Color = Color.Green

            'change the path to real export path when save.
            workbook.Save(IO.Path.Combine(CurrentDirectory, "dest.pdf"), SaveFileFormat.Pdf)
        End Sub
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
